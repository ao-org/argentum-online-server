Attribute VB_Name = "Protocol"

Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR             As String * 1 = vbNullChar

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 255

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer              As New clsByteQueue

Private Enum ServerPacketID

    logged                  ' LOGGED    1
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    EquiteToggle
    CreateRenderText
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM  10
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU
    UserCommerceEnd         ' FINCOMUSUOK
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    NPCKillUser             ' 6
    BlockedWithShieldUser   ' 7
    BlockedWithShieldOther  ' 8
    CharSwing               ' U1          20
    SafeModeOn              ' SEGON
    SafeModeOff             ' SEGOFF
    PartySafeOn
    PartySafeOff
    CantUseWhileMeditating  ' M!
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateExp               ' ASE       30
    ChangeMap               ' CM
    PosUpdate               ' PU
    NPCHitUser              ' N2
    UserHitNPC              ' U2
    UserAttackedSwing       ' U3
    UserHittedByUser        ' N4
    UserHittedUser          ' N5
    ChatOverHead            ' ||        38
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+
    ShowMessageBox          ' !!
    MostrarCuenta
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterMove           ' MP, +, * and _ '
    CharacterChange         ' CP
    ObjectCreate            ' HO
    fxpiso                          '50
    ObjectDelete            ' BO
    BlockPosition           ' BQ
    PlayMidi                ' TM
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    UpdateUserStats         ' EST  60
    WorkRequestTarget       ' T01
    ChangeInventorySlot     ' CSI
    InventoryUnlockSlots
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR   70
    Blind                   ' CEGU
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    AddForumMsg             ' FMSG
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER   80
    DiceRoll                ' DADOS
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    guildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR   90
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    ChangeUserTradeSlot     ' COMUSUINV
    SendNight               ' NOC
    Pong '100
    UpdateTagAndStatus
    FYA
    CerrarleCliente
    Contadores
    
    'GM messages
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    PersonajesDeCuenta  '110
    UserOnline
    ParticleFX
    ParticleFXToFloor
    ParticleFXWithDestino
    ParticleFXWithDestinoXY
    Hora
    light
    AuraToChar
    SpeedTOChar
    LightToFloor  '120
    NieveToggle
    NieblaToggle
    Goliath
    EfectOverHead
    EfectToScreen
    AlquimistaObj
    ShowAlquimiaForm
    Familiar
    SastreObj
    ShowSastreForm
    VelocidadToggle
    MacroTrabajoToggle
    RefreshAllInventorySlot
    BindKeys
    ShowFrmLogear
    ShowFrmMapa
    InmovilizaOK
    BarFx
    SetEscribiendo
    Logros
    TrofeoToggleOn
    TrofeoToggleoff
    LocaleMsg
    ListaCorreo
    ShowPregunta
    DatosGrupo
    Ubicacion
    CorreoPicOn
    DonadorObj
    ExpOverHEad
    OroOverHEad
    ArmaMov
    EscudoMov
    ACTSHOP
    ViajarForm
    Oxigeno
    NadarToggle
    ShowFundarClanForm
    CharUpdateHP
    Ranking
    PosLLamadaDeClan
    QuestDetails
    QuestListSend
    UpdateNPCSimbolo
    ClanSeguro
    Intervals
    UpdateUserKey
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
    BanIP                   '/BANIP
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
    ToggleCentinelActivated '/CENTINELAACTIVADO
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
    SilenciarUser           '/SILENCIAR
    CrearNuevaCuenta
    ValidarCuenta
    IngresarConCuenta
    RevalidarCuenta
    BorrarPJ
    RecuperandoContraseña
    BorrandoCuenta
    NewPacketID
    Desbuggear
    DarLlaveAUsuario
    SacarLlave
    VerLlaves
    UseKey
End Enum

Private Enum NewPacksID

    OfertaInicial
    OfertaDeSubasta
    QuestionGM
    CuentaRegresiva
    PossUser
    Duelo
    NieveToggle
    NieblaToggle
    TransFerGold
    Moveitem
    Genio
    Casarse
    CraftAlquimista
    DropItem
    RequestFamiliar
    FlagTrabajar
    CraftSastre
    MensajeUser
    TraerBoveda
    CompletarAccion
    Escribiendo
    TraerRecompensas
    ReclamarRecompensa
    DecimeLaHora
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

End Enum

Public Type PersonajeCuenta

    nombre As String
    nivel As Byte
    Mapa As Integer
    cuerpo As Integer
    Cabeza As Integer
    Status As Byte
    clase As Byte
    Arma As Integer
    Escudo As Integer
    Casco As Integer
    ClanIndex As Integer

End Type

''
' Handles incoming data.
'
' @param    UserIndex The index of the user sending the message.

Public Function HandleIncomingData(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/09/07
    '
    '***************************************************

    On Error Resume Next

    ' Dim packetID As Byte
    
    '  PaquetesCount = PaquetesCount + 1
    ' frmMain.paquetesRecibidos = PaquetesCount
    
    ' packetID = UserList(UserIndex).incomingData.PeekByte()
    
    Dim packetID As Long
    packetID = CLng(UserList(UserIndex).incomingData.PeekByte())

    'frmMain.listaDePaquetes.AddItem "Paq:" & PaquetesCount & ": " & packetID
    
    ' Debug.Print "Llego paquete ní" & packetID & " pesa: " & UserList(UserIndex).incomingData.length & "Bytes"
    
    'Does the packet requires a logged user??
    If Not (packetID = ClientPacketID.LoginExistingChar Or packetID = ClientPacketID.LoginNewChar Or packetID = ClientPacketID.CrearNuevaCuenta Or packetID = ClientPacketID.IngresarConCuenta Or packetID = ClientPacketID.RevalidarCuenta Or packetID = ClientPacketID.BorrarPJ Or packetID = ClientPacketID.RecuperandoContraseña Or packetID = ClientPacketID.BorrandoCuenta Or packetID = ClientPacketID.ValidarCuenta Or packetID = ClientPacketID.ThrowDice) Then
        
        'Is the user actually logged?
        If Not UserList(UserIndex).flags.UserLogged Then
            Call CloseSocket(UserIndex)
            Exit Function
        
            'He is logged. Reset idle counter if id is valid.
        ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
            UserList(UserIndex).Counters.IdleCount = 0

        End If

    Else
    
        UserList(UserIndex).Counters.IdleCount = 0
        
        ' Envió el primer paquete
        UserList(UserIndex).flags.FirstPacket = True

    End If
    
    Select Case packetID
        
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
            
        Case ClientPacketID.CrearNuevaCuenta
            Call HandleCrearCuenta(UserIndex)
        
        Case ClientPacketID.IngresarConCuenta
            Call HandleIngresarConCuenta(UserIndex)
            
        Case ClientPacketID.ValidarCuenta
            Call HandleValidarCuenta(UserIndex)
            
        Case ClientPacketID.RevalidarCuenta
            Call HandleReValidarCuenta(UserIndex)
            
        Case ClientPacketID.BorrarPJ
            Call HandleBorrarPJ(UserIndex)
            
        Case ClientPacketID.RecuperandoContraseña
            Call HandleRecuperandoContraseña(UserIndex)
        
        Case ClientPacketID.BorrandoCuenta
         
            Call HandleBorrandoCuenta(UserIndex)
        
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
        
        Case ClientPacketID.GrupoMsg               '/ACOMPAíAR
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
            
        Case ClientPacketID.SilenciarUser               '/BAN
            Call HandleSilenciarUser(UserIndex)
            
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
        
        Case ClientPacketID.BanIP                   '/BANIP
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
        
        Case ClientPacketID.ToggleCentinelActivated '/CENTINELAACTIVADO
            Call HandleToggleCentinelActivated(UserIndex)
        
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
            
        Case ClientPacketID.GlobalMessage     '/CONSOLA
            Call HandleGlobalMessage(UserIndex)
        
        Case ClientPacketID.GlobalOnOff        '/GLOBAL
            Call HandleGlobalOnOff(UserIndex)
        
        Case ClientPacketID.NewPacketID    'Los Nuevos Packs ID
            Call HandleIncomingDataNewPacks(UserIndex)

        Case Else
            'ERROR : Abort!
            Call CloseSocket(UserIndex)

    End Select

    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(UserIndex).incomingData.length > 0 And Err.Number = 0 Then
        HandleIncomingData = True
  
    ElseIf Err.Number <> 0 And Not Err.Number = UserList(UserIndex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & Err.Number & " [" & Err.description & "] " & " Source: " & Err.source & vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & vbTab & " LastDllError: " & Err.LastDllError & vbTab & " - UserIndex: " & UserIndex & " - producido al manejar el paquete: " & CStr(packetID))
        Call CloseSocket(UserIndex)
  
        HandleIncomingData = False
    Else
        'Flush buffer - send everything that has been written
        
        HandleIncomingData = False

    End If

End Function

Public Sub HandleIncomingDataNewPacks(ByVal UserIndex As Integer)
        
        On Error GoTo HandleIncomingDataNewPacks_Err
        

        '***************************************************
        'Los nuevos Pack ID
        'Creado por Ladder con gran ayuda de Maraxus
        '04.12.08
        '***************************************************
        Dim packetID As Integer
    
100     packetID = UserList(UserIndex).incomingData.PeekInteger() \ &H100
    
102     Select Case packetID

            Case NewPacksID.OfertaInicial
104             Call HandleOfertaInicial(UserIndex)
    
106         Case NewPacksID.OfertaDeSubasta
108             Call HandleOfertaDeSubasta(UserIndex)
        
110         Case NewPacksID.CuentaRegresiva
112             Call HandleCuentaRegresiva(UserIndex)

114         Case NewPacksID.QuestionGM
116             Call HandleQuestionGM(UserIndex)

118         Case NewPacksID.PossUser
120             Call HandlePossUser(UserIndex)

122         Case NewPacksID.Duelo
124             'Call HandleDuelo(UserIndex)

126         Case NewPacksID.NieveToggle
128             Call HandleNieveToggle(UserIndex)

130         Case NewPacksID.NieblaToggle
132             Call HandleNieblaToggle(UserIndex)

134         Case NewPacksID.TransFerGold
136             Call HandleTransFerGold(UserIndex)

138         Case NewPacksID.Moveitem
140             Call HandleMoveItem(UserIndex)

142         Case NewPacksID.LlamadadeClan
144             Call HandleLlamadadeClan(UserIndex)

146         Case NewPacksID.QuieroFundarClan
148             Call HandleQuieroFundarClan(UserIndex)

150         Case NewPacksID.BovedaMoveItem
152             Call HandleBovedaMoveItem(UserIndex)

154         Case NewPacksID.Genio
156             Call HandleGenio(UserIndex)

158         Case NewPacksID.Casarse
160             Call HandleCasamiento(UserIndex)

162         Case NewPacksID.EnviarCodigo
164             Call HandleEnviarCodigo(UserIndex)

166         Case NewPacksID.CrearTorneo
168             Call HandleCrearTorneo(UserIndex)
            
170         Case NewPacksID.ComenzarTorneo
172             Call HandleComenzarTorneo(UserIndex)
            
174         Case NewPacksID.CancelarTorneo
176             Call HandleCancelarTorneo(UserIndex)

178         Case NewPacksID.BusquedaTesoro
180             Call HandleBusquedaTesoro(UserIndex)
            
182         Case NewPacksID.CrearEvento
184             Call HandleCrearEvento(UserIndex)

186         Case NewPacksID.CraftAlquimista
188             Call HandleCraftAlquimia(UserIndex)

190         Case NewPacksID.DropItem
192             Call HandleDropItem(UserIndex)

194         Case NewPacksID.RequestFamiliar
196             Call HandleRequestFamiliar(UserIndex)

198         Case NewPacksID.FlagTrabajar
200             Call HandleFlagTrabajar(UserIndex)

202         Case NewPacksID.CraftSastre
204             Call HandleCraftSastre(UserIndex)

206         Case NewPacksID.MensajeUser
208             Call HandleMensajeUser(UserIndex)

210         Case NewPacksID.TraerBoveda
212             Call HandleTraerBoveda(UserIndex)

214         Case NewPacksID.CompletarAccion
216             Call HandleCompletarAccion(UserIndex)

218         Case NewPacksID.Escribiendo
220             Call HandleEscribiendo(UserIndex)

222         Case NewPacksID.TraerRecompensas
224             Call HandleTraerRecompensas(UserIndex)

226         Case NewPacksID.ReclamarRecompensa
228             Call HandleReclamarRecompensa(UserIndex)

230         Case NewPacksID.DecimeLaHora
232             Call HandleDecimeLaHora(UserIndex)

234         Case NewPacksID.Correo
236             Call HandleCorreo(UserIndex)

238         Case NewPacksID.SendCorreo ' ok
240             Call HandleSendCorreo(UserIndex)

242         Case NewPacksID.RetirarItemCorreo ' ok
244             Call HandleRetirarItemCorreo(UserIndex)

246         Case NewPacksID.BorrarCorreo
248             Call HandleBorrarCorreo(UserIndex) 'ok

250         Case NewPacksID.InvitarGrupo
252             Call HandleInvitarGrupo(UserIndex) 'ok

254         Case NewPacksID.MarcaDeClanPack
256             Call HandleMarcaDeClan(UserIndex)

258         Case NewPacksID.MarcaDeGMPack
260             Call HandleMarcaDeGM(UserIndex)

262         Case NewPacksID.ResponderPregunta 'ok
264             Call HandleResponderPregunta(UserIndex)

266         Case NewPacksID.RequestGrupo
268             Call HandleRequestGrupo(UserIndex) 'ok

270         Case NewPacksID.AbandonarGrupo
272             Call HandleAbandonarGrupo(UserIndex) ' ok

274         Case NewPacksID.HecharDeGrupo
276             Call HandleHecharDeGrupo(UserIndex) 'ok

278         Case NewPacksID.MacroPossent
280             Call HandleMacroPos(UserIndex)

282         Case NewPacksID.SubastaInfo
284             Call HandleSubastaInfo(UserIndex)

286         Case NewPacksID.EventoInfo
288             Call HandleEventoInfo(UserIndex)

290         Case NewPacksID.CrearEvento
292             Call HandleCrearEvento(UserIndex)

294         Case NewPacksID.BanCuenta
296             Call HandleBanCuenta(UserIndex)
            
298         Case NewPacksID.unBanCuenta
300             Call HandleUnBanCuenta(UserIndex)
            
302         Case NewPacksID.BanSerial
304             Call HandleBanSerial(UserIndex)
        
306         Case NewPacksID.unBanSerial
308             Call HandleUnBanSerial(UserIndex)
            
310         Case NewPacksID.CerrarCliente
312             Call HandleCerrarCliente(UserIndex)
            
314         Case NewPacksID.BanTemporal
316             Call HandleBanTemporal(UserIndex)

318         Case NewPacksID.Traershop
320             Call HandleTraerShop(UserIndex)

322         Case NewPacksID.TraerRanking
324             Call HandleTraerRanking(UserIndex)

326         Case NewPacksID.Pareja
328             Call HandlePareja(UserIndex)
            
330         Case NewPacksID.ComprarItem
332             Call HandleComprarItem(UserIndex)
            
334         Case NewPacksID.CompletarViaje
336             Call HandleCompletarViaje(UserIndex)
            
338         Case NewPacksID.ScrollInfo
340             Call HandleScrollInfo(UserIndex)

342         Case NewPacksID.CancelarExit
344             Call HandleCancelarExit(UserIndex)
            
346         Case NewPacksID.Quest
348             Call HandleQuest(UserIndex)
            
350         Case NewPacksID.QuestAccept
352             Call HandleQuestAccept(UserIndex)
        
354         Case NewPacksID.QuestListRequest
356             Call HandleQuestListRequest(UserIndex)
        
358         Case NewPacksID.QuestDetailsRequest
360             Call HandleQuestDetailsRequest(UserIndex)
        
362         Case NewPacksID.QuestAbandon
364             Call HandleQuestAbandon(UserIndex)
            
366         Case NewPacksID.SeguroClan
368             Call HandleSeguroClan(UserIndex)
            
370         Case Else
                'ERROR : Abort!
372             Call CloseSocket(UserIndex)
            
        End Select
    
374     If UserList(UserIndex).incomingData.length > 0 And Err.Number = 0 Then
376         Err.Clear
378         Call HandleIncomingData(UserIndex)
    
380     ElseIf Err.Number <> 0 And Not Err.Number = UserList(UserIndex).incomingData.NotEnoughDataErrCode Then
            'An error ocurred, log it and kick player.
382         Call LogError("Error: " & Err.Number & " [" & Err.description & "] " & " Source: " & Err.source & vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & vbTab & " LastDllError: " & Err.LastDllError & vbTab & " - UserIndex: " & UserIndex & " - producido al manejar el paquete: " & CStr(packetID))
384         Call CloseSocket(UserIndex)
    
        Else
            'Flush buffer - send everything that has been written
        

        End If

        
        Exit Sub

HandleIncomingDataNewPacks_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleIncomingDataNewPacks", Erl)
        Resume Next
        
End Sub

''
' Handles the "LoginExistingChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    ''Last Modification: 01/12/08 Ladder
    '***************************************************
    If UserList(UserIndex).incomingData.length < 16 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName    As String

    Dim CuentaEmail As String

    Dim Password    As String

    Dim Version     As String

    Dim MacAddress  As String

    Dim HDserial    As Long

    CuentaEmail = buffer.ReadASCIIString()
    Password = buffer.ReadASCIIString()
    Version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    UserName = buffer.ReadASCIIString()
    MacAddress = buffer.ReadASCIIString()
    HDserial = buffer.ReadLong()
    
    If Not VersionOK(Version) Then
        Call WriteShowMessageBox(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If

    If Not EntrarCuenta(UserIndex, CuentaEmail, Password, MacAddress, HDserial) Then
        
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

        If BanNick = "" Then
            BanNick = "*Error en la base de datos*"

        End If
        
        If BaneoMotivo = "" Then
            BaneoMotivo = "*No se registra el motivo del baneo.*"

        End If
        
        Call WriteShowMessageBox(UserIndex, "Se te ha prohibido la entrada al juego debido a " & BaneoMotivo & ". Esta decisión fue tomada por " & BanNick & ".")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
        
    Call ConnectUser(UserIndex, UserName, CuentaEmail)

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    If UserList(UserIndex).incomingData.length < 21 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName As String

    Dim race     As eRaza

    Dim gender   As eGenero

    Dim Class As eClass

    Dim Head        As Integer

    Dim CuentaEmail As String

    Dim Password    As String

    Dim MacAddress  As String

    Dim HDserial    As Long

    Dim Version     As String
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(UserIndex, "La creacion de personajes en este servidor se ha deshabilitado.")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If

    If aClon.MaxPersonajes(UserList(UserIndex).ip) Then
        Call WriteErrorMsg(UserIndex, "Has creado demasiados personajes.")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    If ObtenerCantidadDePersonajesByUserIndex(UserIndex) >= MAX_PERSONAJES Then
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    CuentaEmail = buffer.ReadASCIIString()
    Password = buffer.ReadASCIIString()
    Version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    UserName = buffer.ReadASCIIString()
    race = buffer.ReadByte()
    gender = buffer.ReadByte()
    Class = buffer.ReadByte()
    Head = buffer.ReadInteger()
    MacAddress = buffer.ReadASCIIString()
    HDserial = buffer.ReadLong()
    
    If Not VersionOK(Version) Then
        Call WriteShowMessageBox(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If

    If Not EntrarCuenta(UserIndex, CuentaEmail, Password, MacAddress, HDserial) Then
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If

    Call ConnectNewUser(UserIndex, UserName, race, gender, Class, Head, CuentaEmail)

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleThrowDice(ByVal UserIndex As Integer)
        'Remove packet ID
        
        On Error GoTo HandleThrowDice_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
102     With UserList(UserIndex).Stats
104         .UserAtributos(eAtributos.Fuerza) = RandomNumber(DiceMinimum, DiceMaximum)
106         .UserAtributos(eAtributos.Agilidad) = RandomNumber(DiceMinimum, DiceMaximum)
108         .UserAtributos(eAtributos.Inteligencia) = RandomNumber(DiceMinimum, DiceMaximum)
            .UserAtributos(eAtributos.Carisma) = RandomNumber(DiceMinimum, DiceMaximum)
110         .UserAtributos(eAtributos.Constitucion) = RandomNumber(DiceMinimum, DiceMaximum)

        End With
    
112     Call WriteDiceRoll(UserIndex)

        
        Exit Sub

HandleThrowDice_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleThrowDice", Erl)
        Resume Next
        
End Sub

''
' Handles the "Talk" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        
        chat = buffer.ReadASCIIString()
        
        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.name, "Dijo: " & chat)

        End If
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0

            If .flags.invisible = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                'Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(UserIndex, "307", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
       
        If .flags.Silenciado = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Los administradores te han impedido hablar durante los proximos " & .flags.MinutosRestantes & " minutos debido a tu comportamiento.", FontTypeNames.FONTTYPE_VENENO)
            Call WriteLocaleMsg(UserIndex, "110", FontTypeNames.FONTTYPE_INFO, .flags.MinutosRestantes)
        Else

            If LenB(chat) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(chat)
                
                If .flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR, UserList(UserIndex).name))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, .flags.ChatColor, UserList(UserIndex).name))

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        
        chat = buffer.ReadASCIIString()
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Los muertos no pueden comunicarse con el mundo de los vivos.", FontTypeNames.FONTTYPE_INFO)
        Else

            '[Consejeros & GMs]
            If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                Call LogGM(.name, "Grito: " & chat)

            End If
            
            'I see you....
            If .flags.Oculto > 0 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0

                If .flags.invisible = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                End If

            End If
            
            If .flags.Silenciado = 1 Then
                Call WriteLocaleMsg(UserIndex, "110", FontTypeNames.FONTTYPE_VENENO, .flags.MinutosRestantes)
        
                'Call WriteConsoleMsg(UserIndex, "Los administradores te han impedido hablar durante los proximos " & .flags.MinutosRestantes & " minutos debido a tu comportamiento.", FontTypeNames.FONTTYPE_VENENO)
            Else

                If LenB(chat) <> 0 Then
                    'Analize chat...
                    Call Statistics.ParseChat(chat)

                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, vbRed, UserList(UserIndex).name))
               
                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat            As String

        Dim targetCharIndex As String

        Dim targetUserIndex As Integer

        Dim rank            As Integer
        
        rank = UserList(UserIndex).flags.Privilegios

        targetCharIndex = buffer.ReadASCIIString()
        chat = buffer.ReadASCIIString()
        
        targetUserIndex = NameIndex(targetCharIndex)

        If targetUserIndex <= 0 Then 'existe el usuario destino?
            Call WriteConsoleMsg(UserIndex, "Usuario offline o inexistente.", FontTypeNames.FONTTYPE_INFO)
        Else
        
            If rank = 1 And (UserList(targetUserIndex).flags.Privilegios) > 1 Then
                Call WriteConsoleMsg(UserIndex, "No podes hablar por privado con administradores del juego.", FontTypeNames.FONTTYPE_WARNING)
            Else

                If EstaPCarea(UserIndex, targetUserIndex) Then
                    If LenB(chat) <> 0 Then
                        'Analize chat...
                        Call Statistics.ParseChat(chat)
            
                        Call SendData(SendTarget.ToSuperiores, UserIndex, PrepareMessageChatOverHead(chat, .Char.CharIndex, RGB(157, 226, 20)))
                        
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

        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
        

        Dim demora      As Long

        Dim demorafinal As Long

100     demora = timeGetTime

102     If UserList(UserIndex).incomingData.length < 2 Then
104         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
        Dim dummy    As Long

        Dim TempTick As Long

        Dim heading  As eHeading
    
106     With UserList(UserIndex)
            'Remove packet ID
108         Call .incomingData.ReadByte
        
110         heading = .incomingData.ReadByte()
        
112         If .flags.Paralizado = 0 Or .flags.Inmovilizado = 0 Then

114             If .flags.Meditando Then
                    'Stop meditating, next action will start movement.
116                 .flags.Meditando = False
120                 UserList(UserIndex).Char.FX = 0
122                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))
                End If
            
124             'If IntervaloPermiteCaminar(UserIndex) Then
            
                    'Move user
126                 Call MoveUserChar(UserIndex, heading)
                
128                 If UserList(UserIndex).Grupo.EnGrupo = True Then
130                     Call CompartirUbicacion(UserIndex)
                    End If

                    'Stop resting if needed
132                 If .flags.Descansar Then
134                     .flags.Descansar = False
                    
136                     Call WriteRestOK(UserIndex)
                        'Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
138                     Call WriteLocaleMsg(UserIndex, "178", FontTypeNames.FONTTYPE_INFO)

                    End If

                    Dim TiempoDeWalk As Byte

140                 If .flags.Montado = 1 Then
142                     TiempoDeWalk = 37
144                 ElseIf .flags.Muerto = 1 Then
146                     TiempoDeWalk = 40
                    Else
148                     TiempoDeWalk = 34

                    End If
                    
                    'Prevent SpeedHack
150                 If .flags.TimesWalk >= TiempoDeWalk Then
152                     TempTick = GetTickCount And &H7FFFFFFF
154                     dummy = (TempTick - .flags.StartWalk)
                        
                        ' 5800 is actually less than what would be needed in perfect conditions to take 30 steps
                        '(it's about 193 ms per step against the over 200 needed in perfect conditions)
156                     If dummy < 5200 Then
158                         If TempTick - .flags.CountSH > 30000 Then
160                             .flags.CountSH = 0

                            End If
                            
162                         If Not .flags.CountSH = 0 Then
164                             If dummy <> 0 Then dummy = 126000 \ dummy
                               
166                             Call LogHackAttemp("Tramposo SH: " & .name & " , " & dummy)
168                             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha sido echado por el servidor por posible uso de SpeedHack.", FontTypeNames.FONTTYPE_SERVER))
170                             Call CloseSocket(UserIndex)
                                
                                Exit Sub
                            Else
172                             .flags.CountSH = TempTick

                            End If

                        End If

174                     .flags.StartWalk = TempTick
176                     .flags.TimesWalk = 0

                    End If
                    
178                 .flags.TimesWalk = .flags.TimesWalk + 1
                
180                 Call CancelExit(UserIndex)

                'End If

            Else    'paralized

182             If Not .flags.UltimoMensaje = 1 Then
184                 .flags.UltimoMensaje = 1
                    'Call WriteConsoleMsg(UserIndex, "No podes moverte porque estas paralizado.", FontTypeNames.FONTTYPE_INFO)
186                 Call WriteLocaleMsg(UserIndex, "54", FontTypeNames.FONTTYPE_INFO)

                End If
            
188             .flags.CountSH = 0

            End If
        
190         If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
192             .flags.Oculto = 0
194             .Counters.TiempoOculto = 0
                
                'If not under a spell effect, show char
196             If .flags.invisible = 0 Then
198                 Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
200                 Call WriteLocaleMsg(UserIndex, "307", FontTypeNames.FONTTYPE_INFO)
202                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                End If
            End If

        End With
    
204     demorafinal = timeGetTime - demora

        
        Exit Sub

HandleWalk_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleWalk", Erl)
        Resume Next
        
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
        'Remove packet ID
        
        On Error GoTo HandleRequestPositionUpdate_Err
        
100     UserList(UserIndex).incomingData.ReadByte
    
102     Call WritePosUpdate(UserIndex)

        
        Exit Sub

HandleRequestPositionUpdate_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestPositionUpdate", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'If dead, can't attack
104         If .flags.Muerto = 1 Then
106             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "í¡No podes atacar a nadie porque estas muerto!!.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'If equiped weapon is ranged, can't attack this way
108         If .Invent.WeaponEqpObjIndex > 0 Then
110             If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
112                 Call WriteConsoleMsg(UserIndex, "No podés usar así esta arma.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If

            End If
        
114         If .Invent.HerramientaEqpObjIndex > 0 Then
116             Call WriteConsoleMsg(UserIndex, "Para atacar debes desequipar la herramienta.", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If
        
118         If UserList(UserIndex).flags.Meditando Then
120             UserList(UserIndex).flags.Meditando = False
124             UserList(UserIndex).Char.FX = 0
126             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.CharIndex, 0))
            End If
        
            'If exiting, cancel
128         Call CancelExit(UserIndex)
        
            'Attack!
130         Call UsuarioAtaca(UserIndex)
        
            'I see you...
132         If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
134             .flags.Oculto = 0
136             .Counters.TiempoOculto = 0

138             If .flags.invisible = 0 Then
140                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    'Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
142                 Call WriteLocaleMsg(UserIndex, "307", FontTypeNames.FONTTYPE_INFOIAO)

                End If

            End If

        End With

        
        Exit Sub

HandleAttack_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleAttack", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'If dead, it can't pick up objects
104         If .flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Los muertos no pueden tomar objetos.", FontTypeNames.FONTTYPE_INFO)
106             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            
                Exit Sub

            End If
        
            'Lower rank administrators can't pick up items
108         If .flags.Privilegios And PlayerType.Consejero Then
110             If Not .flags.Privilegios And PlayerType.RoleMaster Then
112                 Call WriteConsoleMsg(UserIndex, "No podés tomar ningun objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            End If
        
114         Call GetObj(UserIndex)

        End With

        
        Exit Sub

HandlePickUp_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandlePickUp", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Seguro Then
106             Call WriteSafeModeOff(UserIndex)
            Else
108             Call WriteSafeModeOn(UserIndex)

            End If
        
110         .flags.Seguro = Not .flags.Seguro

        End With

        
        Exit Sub

HandleSafeToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleSafeToggle", Erl)
        Resume Next
        
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
102         Call .incomingData.ReadByte
        
104         .flags.SeguroParty = Not .flags.SeguroParty
        
106         If .flags.SeguroParty Then
108             Call WritePartySafeOn(UserIndex)
            Else
110             Call WritePartySafeOff(UserIndex)

            End If

        End With

        
        Exit Sub

HandlePartyToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandlePartyToggle", Erl)
        Resume Next
        
End Sub

Private Sub HandleSeguroClan(ByVal UserIndex As Integer)
        
        On Error GoTo HandleSeguroClan_Err
        

        '***************************************************
        'Author: Ladder
        'Date: 31/10/20
        '***************************************************
100     With UserList(UserIndex)
102         Call .incomingData.ReadInteger 'Leemos paquete
                
104         .flags.SeguroClan = Not .flags.SeguroClan

106         Call WriteClanSeguro(UserIndex, .flags.SeguroClan)

        End With

        
        Exit Sub

HandleSeguroClan_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleSeguroClan", Erl)
        Resume Next
        
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
        'Remove packet ID
        
        On Error GoTo HandleRequestGuildLeaderInfo_Err
        
100     UserList(UserIndex).incomingData.ReadByte
    
102     Call modGuilds.SendGuildLeaderInfo(UserIndex)

        
        Exit Sub

HandleRequestGuildLeaderInfo_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestGuildLeaderInfo", Erl)
        Resume Next
        
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
        'Remove packet ID
        
        On Error GoTo HandleRequestAtributes_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
102     Call WriteAttributes(UserIndex)

        
        Exit Sub

HandleRequestAtributes_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestAtributes", Erl)
        Resume Next
        
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
        'Remove packet ID
        
        On Error GoTo HandleRequestSkills_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
102     Call WriteSendSkills(UserIndex)

        
        Exit Sub

HandleRequestSkills_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestSkills", Erl)
        Resume Next
        
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
        'Remove packet ID
        
        On Error GoTo HandleRequestMiniStats_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
102     Call WriteMiniStats(UserIndex)

        
        Exit Sub

HandleRequestMiniStats_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestMiniStats", Erl)
        Resume Next
        
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
        'Remove packet ID
        
        On Error GoTo HandleCommerceEnd_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
        'User quits commerce mode
102     If UserList(UserIndex).flags.TargetNPC <> 0 Then
104         If Npclist(UserList(UserIndex).flags.TargetNPC).SoundClose <> 0 Then
106             Call WritePlayWave(UserIndex, Npclist(UserList(UserIndex).flags.TargetNPC).SoundClose, NO_3D_SOUND, NO_3D_SOUND)

            End If

        End If

108     UserList(UserIndex).flags.Comerciando = False
110     Call WriteCommerceEnd(UserIndex)

        
        Exit Sub

HandleCommerceEnd_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCommerceEnd", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Quits commerce mode with user
104         If .ComUsu.DestUsu > 0 And UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
106             Call WriteConsoleMsg(.ComUsu.DestUsu, .name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
108             Call FinComerciarUsu(.ComUsu.DestUsu)
            
                'Send data in the outgoing buffer of the other user
            

            End If
        
110         Call FinComerciarUsu(UserIndex)

        End With

        
        Exit Sub

HandleUserCommerceEnd_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleUserCommerceEnd", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'User exits banking mode
104         .flags.Comerciando = False
        
106         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("171", UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
108         Call WriteBankEnd(UserIndex)

        End With

        
        Exit Sub

HandleBankEnd_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleBankEnd", Erl)
        Resume Next
        
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
        'Remove packet ID
        
        On Error GoTo HandleUserCommerceOk_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
        'Trade accepted
102     Call AceptarComercioUsu(UserIndex)

        
        Exit Sub

HandleUserCommerceOk_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleUserCommerceOk", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         otherUser = .ComUsu.DestUsu
        
            'Offer rejected
106         If otherUser > 0 Then
108             If UserList(otherUser).flags.UserLogged Then
110                 Call WriteConsoleMsg(otherUser, .name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
112                 Call FinComerciarUsu(otherUser)
                
                    'Send data in the outgoing buffer of the other user
                

                End If

            End If
        
114         Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
116         Call FinComerciarUsu(UserIndex)

        End With

        
        Exit Sub

HandleUserCommerceReject_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleUserCommerceReject", Erl)
        Resume Next
        
End Sub

''
' Handles the "Drop" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)
        
        On Error GoTo HandleDrop_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(UserIndex).incomingData.length < 4 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
        Dim slot   As Byte

        Dim Amount As Long
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte

108         slot = .incomingData.ReadByte()
110         Amount = .incomingData.ReadLong()

112         If Not IntervaloPermiteTirar(UserIndex) Then Exit Sub

            'low rank admins can't drop item. Neither can the dead nor those sailing.
114         If .flags.Muerto = 1 Or ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub
        
            'Are we dropping gold or other items??
116         If slot = FLAGORO Then

118             Call TirarOro(Amount, UserIndex)
            
120             Call WriteUpdateGold(UserIndex)
            Else
        
                '04-05-08 Ladder
122             If (.flags.Privilegios And PlayerType.Admin) <> 16 Then
124                 If ObjData(.Invent.Object(slot).ObjIndex).Newbie = 1 Then
126                     Call WriteConsoleMsg(UserIndex, "No se pueden tirar los objetos Newbies.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
128                 If ObjData(.Invent.Object(slot).ObjIndex).Instransferible = 1 Then
130                     Call WriteConsoleMsg(UserIndex, "Acción no permitida.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
132                 If ObjData(.Invent.Object(slot).ObjIndex).Intirable = 1 Then
134                     Call WriteConsoleMsg(UserIndex, "Acción no permitida.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
136                 If UserList(UserIndex).flags.BattleModo = 1 Then
138                     Call WriteConsoleMsg(UserIndex, "No podes tirar items en este mapa.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                End If
        
140             If ObjData(.Invent.Object(slot).ObjIndex).OBJType = eOBJType.otBarcos And UserList(UserIndex).flags.Navegando Then
142                 Call WriteConsoleMsg(UserIndex, "Para tirar la barca deberias estar en tierra firme.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
144             If ObjData(.Invent.Object(slot).ObjIndex).OBJType = eOBJType.otMonturas And UserList(UserIndex).flags.Montado Then
146                 Call WriteConsoleMsg(UserIndex, "Para tirar tu montura deberias descender de ella.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                '04-05-08 Ladder
        
                'Only drop valid slots
148             If slot <= UserList(UserIndex).CurrentInventorySlots And slot > 0 Then
150                 If .Invent.Object(slot).ObjIndex = 0 Then
                        Exit Sub

                    End If
                
152                 Call DropObj(UserIndex, slot, Amount, .Pos.Map, .Pos.x, .Pos.Y)

                End If

            End If

        End With

        
        Exit Sub

HandleDrop_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleDrop", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Spell As Byte
        
108         Spell = .incomingData.ReadByte()
        
110         If .flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
112             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         .flags.Hechizo = Spell
        
116         If .flags.Hechizo < 1 Then
118             .flags.Hechizo = 0
120         ElseIf .flags.Hechizo > MAXUSERHECHIZOS Then
122             .flags.Hechizo = 0

            End If
        
124         If .flags.Hechizo <> 0 Then

                Dim uh As Integer
            
126             uh = UserList(UserIndex).Stats.UserHechizos(Spell)

128             If Hechizos(uh).AutoLanzar = 1 Then
130                 UserList(UserIndex).flags.TargetUser = UserIndex
132                 Call LanzarHechizo(.flags.Hechizo, UserIndex)
                Else
134                 Call WriteWorkRequestTarget(UserIndex, eSkill.magia)

                End If

            End If
        
        End With

        
        Exit Sub

HandleCastSpell_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCastSpell", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 3 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex).incomingData
            'Remove packet ID
106         Call .ReadByte
        
            Dim x As Byte

            Dim Y As Byte
        
108         x = .ReadByte()
110         Y = .ReadByte()
        
112         Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, Y)

        End With

        
        Exit Sub

HandleLeftClick_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleLeftClick", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 3 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex).incomingData
            'Remove packet ID
106         Call .ReadByte
        
            Dim x As Byte

            Dim Y As Byte
        
108         x = .ReadByte()
110         Y = .ReadByte()
        
112         Call accion(UserIndex, UserList(UserIndex).Pos.Map, x, Y)

        End With

        
        Exit Sub

HandleDoubleClick_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleDoubleClick", Erl)
        Resume Next
        
End Sub

''
' Handles the "Work" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)
        
        On Error GoTo HandleWork_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Skill As eSkill
        
108         Skill = .incomingData.ReadByte()
        
110         If UserList(UserIndex).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
112             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'If exiting, cancel
114         Call CancelExit(UserIndex)
        
116         Select Case Skill

                Case Robar, magia, Domar
118                 Call WriteWorkRequestTarget(UserIndex, Skill)

120             Case Ocultarse

122                 If .flags.Navegando = 1 Then

                        '[CDT 17-02-2004]
124                     If Not .flags.UltimoMensaje = 3 Then
                            'Call WriteConsoleMsg(UserIndex, "No podés ocultarte si estás navegando.", FontTypeNames.FONTTYPE_INFO)
126                         Call WriteLocaleMsg(UserIndex, "56", FontTypeNames.FONTTYPE_INFO)
128                         .flags.UltimoMensaje = 3

                        End If

                        '[/CDT]
                        Exit Sub

                    End If
                
130                 If .flags.Montado = 1 Then

                        '[CDT 17-02-2004]
132                     If Not .flags.UltimoMensaje = 3 Then
134                         Call WriteConsoleMsg(UserIndex, "No podés ocultarte si estás montado.", FontTypeNames.FONTTYPE_INFO)
136                         .flags.UltimoMensaje = 3

                        End If

                        '[/CDT]
                        Exit Sub

                    End If

138                 If .flags.Oculto = 1 Then

                        '[CDT 17-02-2004]
140                     If Not .flags.UltimoMensaje = 2 Then
142                         Call WriteLocaleMsg(UserIndex, "55", FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
144                         .flags.UltimoMensaje = 2

                        End If

                        '[/CDT]
                        Exit Sub

                    End If
                
146                 Call DoOcultarse(UserIndex)

            End Select

        End With

        
        Exit Sub

HandleWork_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleWork", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
104         Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.name & " fue expulsado por Anti-macro de hechizos", FontTypeNames.FONTTYPE_VENENO))
106         Call WriteErrorMsg(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros")
        
108         Call CloseSocket(UserIndex)

        End With

        
        Exit Sub

HandleUseSpellMacro_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleUseSpellMacro", Erl)
        Resume Next
        
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
        

100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim slot As Byte
        
108         slot = .incomingData.ReadByte()
        
110         If slot <= UserList(UserIndex).CurrentInventorySlots And slot > 0 Then
112             If .Invent.Object(slot).ObjIndex = 0 Then Exit Sub

114             Call UseInvItem(UserIndex, slot)

            End If


        End With

        
        Exit Sub

HandleUseItem_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleUseItem", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 3 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex).incomingData
            'Remove packet ID
106         Call .ReadByte
        
            Dim Item As Integer
        
108         Item = .ReadInteger()
        
110         If Item < 1 Then Exit Sub
        
            ' If ObjData(Item).SkHerreria = 0 Then Exit Sub
        
112         Call HerreroConstruirItem(UserIndex, Item)

        End With

        
        Exit Sub

HandleCraftBlacksmith_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCraftBlacksmith", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 3 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex).incomingData
            'Remove packet ID
106         Call .ReadByte
        
            Dim Item As Integer
        
108         Item = .ReadInteger()
        
110         If Item < 1 Then Exit Sub
        
            'If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        
112         Call CarpinteroConstruirItem(UserIndex, Item)

        End With

        
        Exit Sub

HandleCraftCarpenter_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCraftCarpenter", Erl)
        Resume Next
        
End Sub

Private Sub HandleCraftAlquimia(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCraftAlquimia_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(UserIndex).incomingData.length < 4 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex).incomingData
            'Remove packet ID
106         Call .ReadInteger
        
            Dim Item As Integer
        
108         Item = .ReadInteger()
        
110         If Item < 1 Then Exit Sub

112         Call AlquimistaConstruirItem(UserIndex, Item)

        End With

        
        Exit Sub

HandleCraftAlquimia_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCraftAlquimia", Erl)
        Resume Next
        
End Sub

Private Sub HandleCraftSastre(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCraftSastre_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(UserIndex).incomingData.length < 4 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex).incomingData
            'Remove packet ID
106         Call .ReadInteger
        
            Dim Item As Integer
        
108         Item = .ReadInteger()
        
110         If Item < 1 Then Exit Sub
            'If ObjData(Item).SkMAGOria = 0 Then Exit Sub

112         Call SastreConstruirItem(UserIndex, Item)

        End With

        
        Exit Sub

HandleCraftSastre_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCraftSastre", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 4 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim x        As Byte
            Dim Y        As Byte

            Dim Skill    As eSkill
            Dim DummyInt As Integer

            Dim tU       As Integer   'Target user
            Dim tN       As Integer   'Target NPC
        
108         x = .incomingData.ReadByte()
110         Y = .incomingData.ReadByte()
        
112         Skill = .incomingData.ReadByte()
            
            'No te dejo trabajar si tenes el inventario lleno.
            If .Invent.NroItems = .CurrentInventorySlots Then
                Call WriteConsoleMsg(UserIndex, "No podés trabajar con el inventario lleno.", FONTTYPE_INFO)
                Call WriteWorkRequestTarget(UserIndex, 0)
                Exit Sub
            End If
            
114         If .flags.Muerto = 1 Or .flags.Descansar Or Not InMapBounds(.Pos.Map, x, Y) Then Exit Sub

        
116         If Not InRangoVision(UserIndex, x, Y) Then
118             Call WritePosUpdate(UserIndex)
                Exit Sub

            End If
            
            If .flags.Meditando Then
                .flags.Meditando = False
                .Char.FX = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.CharIndex, 0))
            End If
        
            'If exiting, cancel
120         Call CancelExit(UserIndex)
        
122         Select Case Skill

                Dim fallo As Boolean

                Case eSkill.Proyectiles
            
                    'Check attack interval
124                 If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Sub

                    'Check Magic interval
126                 If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub

                    'Check bow's interval
128                 If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                    'Make sure the item is valid and there is ammo equipped.
130                 With .Invent

132                     If .WeaponEqpObjIndex = 0 Then
134                         DummyInt = 1
136                     ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > UserList(UserIndex).CurrentInventorySlots Then
138                         DummyInt = 1
140                     ElseIf .MunicionEqpSlot < 1 Or .MunicionEqpSlot > UserList(UserIndex).CurrentInventorySlots Then
142                         DummyInt = 1
144                     ElseIf .MunicionEqpObjIndex = 0 Then
146                         DummyInt = 1
148                     ElseIf ObjData(.WeaponEqpObjIndex).proyectil <> 1 Then
150                         DummyInt = 2
152                     ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
154                         DummyInt = 1
156                     ElseIf .Object(.MunicionEqpSlot).Amount < 1 Then
158                         DummyInt = 1

                        End If
                    
160                     If DummyInt <> 0 Then
162                         If DummyInt = 1 Then
164                             Call WriteConsoleMsg(UserIndex, "No tenés municiones.", FontTypeNames.FONTTYPE_INFO)
                            
                                'Call Desequipar(UserIndex, .WeaponEqpSlot)
166                             Call WriteWorkRequestTarget(UserIndex, 0)

                            End If
                        
168                         Call Desequipar(UserIndex, .MunicionEqpSlot)
170                         Call WriteWorkRequestTarget(UserIndex, 0)
                            Exit Sub

                        End If

                    End With
                
                    'Quitamos stamina
172                 If .Stats.MinSta >= 10 Then
174                     Call QuitarSta(UserIndex, RandomNumber(1, 10))
176                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(UserList(UserIndex).Char.CharIndex))
                    Else
178                     Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                        ' Call WriteConsoleMsg(UserIndex, "Estís muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
180                     Call WriteWorkRequestTarget(UserIndex, 0)
                        Exit Sub

                    End If
                
182                 Call LookatTile(UserIndex, .Pos.Map, x, Y)
                
184                 tU = .flags.TargetUser
186                 tN = .flags.TargetNPC
188                 fallo = True

                    'Validate target
190                 If tU > 0 Then

                        'Only allow to atack if the other one can retaliate (can see us)
192                     If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
194                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
196                         Call WriteWorkRequestTarget(UserIndex, 0)
                            Exit Sub

                        End If
                    
                        'Prevent from hitting self
198                     If tU = UserIndex Then
200                         Call WriteConsoleMsg(UserIndex, "¡No podés atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
202                         Call WriteWorkRequestTarget(UserIndex, 0)
                            Exit Sub

                        End If
                    
                        'Attack!
204                     If Not PuedeAtacar(UserIndex, tU) Then Exit Sub 'TODO: Por ahora pongo esto para solucionar lo anterior.
                    
                        Dim backup    As Byte

                        Dim envie     As Boolean

                        Dim Particula As Integer

                        Dim Tiempo    As Long
                    
206                     Select Case ObjData(.Invent.MunicionEqpObjIndex).Subtipo

                            Case 1 'Paraliza
208                             backup = UserList(UserIndex).flags.Paraliza
210                             UserList(UserIndex).flags.Paraliza = 1

212                         Case 2 ' Incinera
214                             backup = UserList(UserIndex).flags.incinera
216                             UserList(UserIndex).flags.incinera = 1

218                         Case 3 ' envenena
220                             backup = UserList(UserIndex).flags.Envenena
222                             UserList(UserIndex).flags.Envenena = 1

224                         Case 4 ' Explosiva

                        End Select

226                     Call UsuarioAtacaUsuario(UserIndex, tU)
                    
228                     Select Case ObjData(.Invent.MunicionEqpObjIndex).Subtipo

                            Case 0

230                         Case 1 'Paraliza
232                             UserList(UserIndex).flags.Paraliza = backup

234                         Case 2 ' Incinera
236                             UserList(UserIndex).flags.incinera = backup

238                         Case 3 ' envenena
240                             UserList(UserIndex).flags.Envenena = backup

242                         Case 4 ' Explosiva

                        End Select
                    
244                     If ObjData(.Invent.MunicionEqpObjIndex).CreaFX <> 0 Then
246                         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageCreateFX(UserList(tU).Char.CharIndex, ObjData(.Invent.MunicionEqpObjIndex).CreaFX, 0))

                        End If
                    
248                     If ObjData(.Invent.MunicionEqpObjIndex).CreaParticula <> "" Then
                    
250                         Particula = val(ReadField(1, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
252                         Tiempo = val(ReadField(2, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
254                         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, Particula, Tiempo, False))

                        End If
                    
256                     fallo = False
                    
258                 ElseIf tN > 0 Then

                        'Only allow to atack if the other one can retaliate (can see us)
260                     If Abs(Npclist(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(tN).Pos.x - .Pos.x) > RANGO_VISION_X Then
262                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
264                         Call WriteWorkRequestTarget(UserIndex, 0)
                            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                            Exit Sub

                        End If
                    
                        'Is it attackable???
266                     If Npclist(tN).Attackable <> 0 Then
                    
268                         fallo = False
                        
                            'Attack!
                        
270                         Select Case UsuarioAtacaNpcFunction(UserIndex, tN)
                        
                                Case 0 ' no se puede pegar
                            
272                             Case 1 ' le pego
                
274                                 If Npclist(tN).flags.Snd2 > 0 Then
276                                     Call SendData(SendTarget.ToNPCArea, tN, PrepareMessagePlayWave(Npclist(tN).flags.Snd2, Npclist(tN).Pos.x, Npclist(tN).Pos.Y))
                                    Else
278                                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(tN).Pos.x, Npclist(tN).Pos.Y))

                                    End If
                                
280                                 If ObjData(.Invent.MunicionEqpObjIndex).Subtipo = 1 And UserList(UserIndex).flags.TargetNPC > 0 Then
282                                     If Npclist(tN).flags.Paralizado = 0 Then

                                            Dim Probabilidad As Byte

284                                         Probabilidad = RandomNumber(1, 2)

286                                         If Probabilidad = 1 Then
288                                             If Npclist(tN).flags.AfectaParalisis = 0 Then
290                                                 Npclist(tN).flags.Paralizado = 1
                                                
292                                                 Npclist(tN).Contadores.Paralisis = IntervaloParalizado

294                                                 If UserList(UserIndex).ChatCombate = 1 Then
                                                        'Call WriteConsoleMsg(UserIndex, "Tu golpe a paralizado a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
296                                                     Call WriteLocaleMsg(UserIndex, "136", FontTypeNames.FONTTYPE_FIGHT)

                                                    End If

298                                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(Npclist(tN).Char.CharIndex, 8, 0))
300                                                 envie = True
                                                Else

302                                                 If UserList(UserIndex).ChatCombate = 1 Then
                                                        'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
304                                                     Call WriteLocaleMsg(UserIndex, "381", FontTypeNames.FONTTYPE_INFO)

                                                    End If

                                                End If

                                            End If

                                        End If

                                    End If
                                
306                                 If ObjData(.Invent.MunicionEqpObjIndex).CreaFX <> 0 Then
308                                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(Npclist(tN).Char.CharIndex, ObjData(.Invent.MunicionEqpObjIndex).CreaFX, 0))

                                    End If
                    
310                                 If ObjData(.Invent.MunicionEqpObjIndex).CreaParticula <> "" Then
312                                     Particula = val(ReadField(1, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
314                                     Tiempo = val(ReadField(2, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
316                                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(Npclist(tN).Char.CharIndex, Particula, Tiempo, False))

                                    End If
                                
318                             Case 2 ' Fallo
                            
                            End Select
                        
                        End If

                    End If
                
320                 With .Invent
322                     DummyInt = .MunicionEqpSlot
                    
                        'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                    
324                     If Not fallo Then
326                         Call QuitarUserInvItem(UserIndex, DummyInt, 1)

                        End If
                    
                        'Call DropObj(UserIndex, .MunicionEqpSlot, 1, UserList(UserIndex).Pos.Map, x, y)
                   
                        ' If fallo And MapData(UserList(UserIndex).Pos.Map, x, Y).Blocked = 0 Then
                        '  Dim flecha As obj
                        ' flecha.Amount = 1
                        'flecha.ObjIndex = .MunicionEqpObjIndex
                        ' Call MakeObj(flecha, UserList(UserIndex).Pos.Map, x, Y)
                        ' End If
                    
328                     If .Object(DummyInt).Amount > 0 Then
                            'QuitarUserInvItem unequipps the ammo, so we equip it again
330                         .MunicionEqpSlot = DummyInt
332                         .MunicionEqpObjIndex = .Object(DummyInt).ObjIndex
334                         .Object(DummyInt).Equipped = 1
                        Else
336                         .MunicionEqpSlot = 0
338                         .MunicionEqpObjIndex = 0

                        End If

340                     Call UpdateUserInv(False, UserIndex, DummyInt)

                    End With

                    '-----------------------------------
            
342             Case eSkill.magia
                    'Check the map allows spells to be casted.
                    '  If MapInfo(.Pos.map).MagiaSinEfecto > 0 Then
                    ' Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energía", FontTypeNames.FONTTYPE_FIGHT)
                    '  Exit Sub
                    ' End If
                
                    'Target whatever is in that tile
344                 Call LookatTile(UserIndex, .Pos.Map, x, Y)
                
                    'If it's outside range log it and exit
346                 If Abs(.Pos.x - x) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
348                     Call LogCheating("Ataque fuera de rango de " & .name & "(" & .Pos.Map & "/" & .Pos.x & "/" & .Pos.Y & ") ip: " & .ip & " a la posicion (" & .Pos.Map & "/" & x & "/" & Y & ")")
                        Exit Sub

                    End If
                
                    'Check bow's interval
350                 If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                    'Check attack-spell interval
352                 If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
                
                    'Check Magic interval
354                 If Not IntervaloPermiteLanzarSpell(UserIndex) Then Exit Sub
                
                    'Check intervals and cast
356                 If .flags.Hechizo > 0 Then
358                     Call LanzarHechizo(.flags.Hechizo, UserIndex)
360                     .flags.Hechizo = 0
                    Else
362                     Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)

                    End If
            
364             Case eSkill.Pescar
                
366                 If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                
                    'Check interval
368                 If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
370                 Select Case .Invent.HerramientaEqpObjIndex
                
                        Case CAÑA_PESCA, CAÑA_PESCA_DORADA

372                         If HayAgua(.Pos.Map, x, Y) Then
374                             Call DoPescar(UserIndex, False, .Invent.HerramientaEqpObjIndex = CAÑA_PESCA_DORADA)
376                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.x, .Pos.Y))
                            Else
378                             Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
380                             Call WriteMacroTrabajoToggle(UserIndex, False)
    
                            End If
                    
382                     Case RED_PESCA
    
384                         If HayAgua(.Pos.Map, x, Y) Then
                            
386                             If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 8 Then
388                                 Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
390                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub
    
                                End If
                                
392                             If UserList(UserIndex).Stats.UserSkills(eSkill.Pescar) < 80 Then
394                                 Call WriteConsoleMsg(UserIndex, "Para utilizar la red de pesca debes tener 80 skills en recoleccion.", FontTypeNames.FONTTYPE_INFO)
396                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub
    
                                End If
                                    
398                             If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
400                                 Call WriteConsoleMsg(UserIndex, "Esta prohibida la pesca masiva en las ciudades.", FontTypeNames.FONTTYPE_INFO)
402                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub
    
                                End If
                                    
404                             If UserList(UserIndex).flags.Navegando = 0 Then
406                                 Call WriteConsoleMsg(UserIndex, "Necesitas estar sobre tu barca para utilizar la red de pesca.", FontTypeNames.FONTTYPE_INFO)
408                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub
    
                                End If
                                    
410                             Call DoPescar(UserIndex, True, True)
412                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.x, .Pos.Y))
                        
                            Else
                        
414                             Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
416                             Call WriteWorkRequestTarget(UserIndex, 0)
    
                            End If
                
                    End Select
                
                    
418             Case eSkill.Talar
            
420                 If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
            
                    'Check interval
422                 If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

424                 Select Case .Invent.HerramientaEqpObjIndex
                
                        Case HACHA_LEÑADOR, HACHA_LEÑADOR_DORADA
                        
                            'Target whatever is in the tile
426                         Call LookatTile(UserIndex, .Pos.Map, x, Y)

                            ' Ahora se puede talar en la ciudad
                            'If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
                            '    Call WriteConsoleMsg(UserIndex, "Esta prohibido talar arboles en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                            '    Call WriteWorkRequestTarget(UserIndex, 0)
                            '    Exit Sub
                            'End If
                            
428                         DummyInt = MapData(.Pos.Map, x, Y).ObjInfo.ObjIndex
                            
430                         If DummyInt > 0 Then
432                             If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 1 Then
434                                 Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
436                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                
438                             If .Pos.x = x And .Pos.Y = Y Then
440                                 Call WriteConsoleMsg(UserIndex, "No podés talar desde allí.", FontTypeNames.FONTTYPE_INFO)
442                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                
444                             If MapData(.Pos.Map, x, Y).ObjInfo.Amount <= 0 Then
446                                 Call WriteConsoleMsg(UserIndex, "El árbol ya no te puede entregar mas leña.", FontTypeNames.FONTTYPE_INFO)
448                                 Call WriteWorkRequestTarget(UserIndex, 0)
450                                 Call WriteMacroTrabajoToggle(UserIndex, False)
                                    Exit Sub

                                End If

                                '¡Hay un arbol donde clickeo?
452                             If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
454                                 Call DoTalar(UserIndex, x, Y, .Invent.HerramientaEqpObjIndex = HACHA_LEÑADOR_DORADA)

                                End If

                            Else
456                             Call WriteConsoleMsg(UserIndex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
458                             Call WriteWorkRequestTarget(UserIndex, 0)

460                             If UserList(UserIndex).Counters.Trabajando > 1 Then
462                                 Call WriteMacroTrabajoToggle(UserIndex, False)

                                End If

                            End If
                
                    End Select
            
464             Case eSkill.Alquimia
            
466                 If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                
                    'Check interval
468                 If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

470                 Select Case .Invent.HerramientaEqpObjIndex
                
                        Case TIJERAS, TIJERAS_DORADAS

472                         If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
474                             Call WriteWorkRequestTarget(UserIndex, 0)
476                             Call WriteConsoleMsg(UserIndex, "Esta prohibido cortar raices en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If
                            
478                         If MapData(.Pos.Map, x, Y).ObjInfo.Amount <= 0 Then
480                             Call WriteConsoleMsg(UserIndex, "El árbol ya no te puede entregar mas raices.", FontTypeNames.FONTTYPE_INFO)
482                             Call WriteWorkRequestTarget(UserIndex, 0)
484                             Call WriteMacroTrabajoToggle(UserIndex, False)
                                Exit Sub

                            End If
                
486                         DummyInt = MapData(.Pos.Map, x, Y).ObjInfo.ObjIndex
                            
488                         If DummyInt > 0 Then
                            
490                             If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 2 Then
492                                 Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
494                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                
496                             If .Pos.x = x And .Pos.Y = Y Then
498                                 Call WriteConsoleMsg(UserIndex, "No podés quitar raices allí.", FontTypeNames.FONTTYPE_INFO)
500                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                
                                '¡Hay un arbol donde clickeo?
502                             If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
504                                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_TIJERAS, .Pos.x, .Pos.Y))
506                                 Call DoRaices(UserIndex, x, Y)

                                End If

                            Else
508                             Call WriteConsoleMsg(UserIndex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
510                             Call WriteWorkRequestTarget(UserIndex, 0)
512                             Call WriteMacroTrabajoToggle(UserIndex, False)

                            End If
                
                    End Select
                
514             Case eSkill.Mineria
            
516                 If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                
                    'Check interval
518                 If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub

520                 Select Case .Invent.HerramientaEqpObjIndex
                
                        Case PIQUETE_MINERO, PIQUETE_MINERO_DORADA
                
                            'Target whatever is in the tile
522                         Call LookatTile(UserIndex, .Pos.Map, x, Y)
                            
524                         DummyInt = MapData(.Pos.Map, x, Y).ObjInfo.ObjIndex
                            
526                         If DummyInt > 0 Then

                                'Check distance
528                             If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 2 Then
530                                 Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
532                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                
534                             If MapData(.Pos.Map, x, Y).ObjInfo.Amount <= 0 Then
536                                 Call WriteConsoleMsg(UserIndex, "Este yacimiento no tiene mas minerales para entregar.", FontTypeNames.FONTTYPE_INFO)
538                                 Call WriteWorkRequestTarget(UserIndex, 0)
540                                 Call WriteMacroTrabajoToggle(UserIndex, False)
                                    Exit Sub

                                End If
                                
542                             DummyInt = MapData(.Pos.Map, x, Y).ObjInfo.ObjIndex 'CHECK

                                '¡Hay un yacimiento donde clickeo?
544                             If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
546                                 Call DoMineria(UserIndex, x, Y, .Invent.HerramientaEqpObjIndex = PIQUETE_MINERO_DORADA)
                                Else
548                                 Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
550                                 Call WriteWorkRequestTarget(UserIndex, 0)

                                End If

                            Else
552                             Call WriteConsoleMsg(UserIndex, "Ahí no hay ningun yacimiento.", FontTypeNames.FONTTYPE_INFO)
554                             Call WriteWorkRequestTarget(UserIndex, 0)

                            End If

                    End Select

556             Case eSkill.Robar

                    'Does the map allow us to steal here?
558                 If MapInfo(.Pos.Map).Seguro = 0 Then
                    
                        'Check interval
560                     If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                        'Target whatever is in that tile
562                     Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, Y)
                    
564                     tU = .flags.TargetUser
                    
566                     If tU > 0 And tU <> UserIndex Then

                            'Can't steal administrative players
568                         If UserList(tU).flags.Privilegios And PlayerType.user Then
570                             If UserList(tU).flags.Muerto = 0 Then
572                                 If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 2 Then
574                                     Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                        'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
576                                     Call WriteWorkRequestTarget(UserIndex, 0)
                                        Exit Sub

                                    End If
                                 
                                    '17/09/02
                                    'Check the trigger
578                                 If MapData(UserList(tU).Pos.Map, x, Y).trigger = eTrigger.ZONASEGURA Then
580                                     Call WriteConsoleMsg(UserIndex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
582                                     Call WriteWorkRequestTarget(UserIndex, 0)
                                        Exit Sub

                                    End If
                                 
584                                 If MapData(.Pos.Map, .Pos.x, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
586                                     Call WriteConsoleMsg(UserIndex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
588                                     Call WriteWorkRequestTarget(UserIndex, 0)
                                        Exit Sub

                                    End If
                                 
590                                 Call DoRobar(UserIndex, tU)

                                End If

                            End If

                        Else
592                         Call WriteConsoleMsg(UserIndex, "No a quien robarle!", FontTypeNames.FONTTYPE_INFO)
594                         Call WriteWorkRequestTarget(UserIndex, 0)

                        End If

                    Else
596                     Call WriteConsoleMsg(UserIndex, "¡No podés robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)
598                     Call WriteWorkRequestTarget(UserIndex, 0)

                    End If
                    
                Case eSkill.Domar
                    'Modificado 25/11/02
                    'Optimizado y solucionado el bug de la doma de
                    'criaturas hostiles.
                    
                    'Target whatever is that tile
                    Call LookatTile(UserIndex, .Pos.Map, x, Y)
                    tN = .flags.TargetNPC
                    
                    If tN > 0 Then
                        If Npclist(tN).flags.Domable > 0 Then
                            If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
    
                            End If
                            
                            If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
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
               
600             Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
            
                    'Check interval
602                 If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
604                 Call LookatTile(UserIndex, .Pos.Map, x, Y)
                
                    'Check there is a proper item there
606                 If .flags.TargetObj > 0 Then
608                     If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then

                            'Validate other items
610                         If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > UserList(UserIndex).CurrentInventorySlots Then
                                Exit Sub

                            End If
                        
                            ''chequeamos que no se zarpe duplicando oro
612                         If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
614                             If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
616                                 Call WriteConsoleMsg(UserIndex, "No tienes más minerales", FontTypeNames.FONTTYPE_INFO)
618                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                            
                                ''FUISTE
620                             Call WriteErrorMsg(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                            
622                             Call CloseSocket(UserIndex)
                                Exit Sub

                            End If
                        
624                         Call FundirMineral(UserIndex)
                        Else
626                         Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
628                         Call WriteWorkRequestTarget(UserIndex, 0)

630                         If UserList(UserIndex).Counters.Trabajando > 1 Then
632                             Call WriteMacroTrabajoToggle(UserIndex, False)

                            End If

                        End If

                    Else
634                     Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
636                     Call WriteWorkRequestTarget(UserIndex, 0)

638                     If UserList(UserIndex).Counters.Trabajando > 1 Then
640                         Call WriteMacroTrabajoToggle(UserIndex, False)

                        End If

                    End If

642             Case eSkill.Grupo
                    'If UserList(UserIndex).Grupo.EnGrupo = False Then
                    'Target whatever is in that tile
644                 Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, Y)
                    
646                 tU = .flags.TargetUser
                    
                    'Call WritePreguntaBox(UserIndex, UserList(UserIndex).name & " te invitó a unirte a su grupo. ¿Deseas unirte?")
                    
648                 If tU > 0 And tU <> UserIndex Then

                        'Can't steal administrative players
650                     If UserList(UserIndex).Grupo.EnGrupo = False Then
652                         If UserList(tU).flags.Muerto = 0 Then
654                             If Abs(.Pos.x - x) + Abs(.Pos.Y - Y) > 8 Then
656                                 Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
658                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                         
660                             If UserList(UserIndex).Grupo.CantidadMiembros = 0 Then
662                                 UserList(UserIndex).Grupo.Lider = UserIndex
664                                 UserList(UserIndex).Grupo.Miembros(1) = UserIndex
666                                 UserList(UserIndex).Grupo.CantidadMiembros = 1
668                                 Call InvitarMiembro(UserIndex, tU)
                                Else
670                                 UserList(UserIndex).Grupo.Lider = UserIndex
672                                 Call InvitarMiembro(UserIndex, tU)

                                End If
                                         
                            Else
674                             Call WriteLocaleMsg(UserIndex, "7", FontTypeNames.FONTTYPE_INFO)
                                'Call WriteConsoleMsg(UserIndex, "El usuario esta muerto.", FontTypeNames.FONTTYPE_INFOIAO)
676                             Call WriteWorkRequestTarget(UserIndex, 0)

                            End If

                        Else

678                         If UserList(UserIndex).Grupo.Lider = UserIndex Then
680                             Call InvitarMiembro(UserIndex, tU)
                            Else
682                             Call WriteConsoleMsg(UserIndex, "Tu no podés invitar usuarios, debe hacerlo " & UserList(UserList(UserIndex).Grupo.Lider).name & ".", FontTypeNames.FONTTYPE_INFOIAO)
684                             Call WriteWorkRequestTarget(UserIndex, 0)

                            End If

                        End If

                    Else
686                     Call WriteLocaleMsg(UserIndex, "261", FontTypeNames.FONTTYPE_INFO)

                    End If

                    ' End If
688             Case eSkill.MarcaDeClan

                    'If UserList(UserIndex).Grupo.EnGrupo = False Then
                    'Target whatever is in that tile
                    Dim clan_nivel As Byte
                
690                 If UserList(UserIndex).GuildIndex = 0 Then
692                     Call WriteConsoleMsg(UserIndex, "Servidor> No Perteneces a ningun clan.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
                
694                 clan_nivel = modGuilds.NivelDeClan(UserList(UserIndex).GuildIndex)

696                 If clan_nivel < 4 Then
698                     Call WriteConsoleMsg(UserIndex, "Servidor> El nivel de tu clan debe ser 4 para utilizar esta opción.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
                                
700                 Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, Y)
                    
702                 tU = .flags.TargetUser

704                 If tU = 0 Then Exit Sub
                    
706                 If UserList(UserIndex).GuildIndex = UserList(tU).GuildIndex Then
708                     Call WriteConsoleMsg(UserIndex, "Servidor> No podes marcar a un miembro de tu clan.", FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If
                    
                    'Call WritePreguntaBox(UserIndex, UserList(UserIndex).name & " te invitó a unirte a su grupo. ¿Deseas unirte?")
                    
710                 If tU > 0 And tU <> UserIndex Then

                        'Can't steal administrative players
712                     If UserList(tU).flags.Muerto = 0 Then
                            'call marcar
714                         Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, 210, 700, False))
716                         Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageConsoleMsg("Clan> [" & UserList(UserIndex).name & "] marco a " & UserList(tU).name & ".", FontTypeNames.FONTTYPE_GUILD))
                        Else
718                         Call WriteLocaleMsg(UserIndex, "7", FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "El usuario esta muerto.", FontTypeNames.FONTTYPE_INFOIAO)
720                         Call WriteWorkRequestTarget(UserIndex, 0)

                        End If

                    Else
722                     Call WriteLocaleMsg(UserIndex, "261", FontTypeNames.FONTTYPE_INFO)

                    End If

724             Case eSkill.MarcaDeGM
726                 Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, x, Y)
                    
728                 tU = .flags.TargetUser

730                 If tU > 0 Then
732                     Call WriteConsoleMsg(UserIndex, "Servidor> [" & UserList(tU).name & "] seleccionado.", FontTypeNames.FONTTYPE_SERVER)
                    Else
734                     Call WriteLocaleMsg(UserIndex, "261", FontTypeNames.FONTTYPE_INFO)

                    End If
                    
            End Select

        End With

        
        Exit Sub

HandleWorkLeftClick_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleWorkLeftClick", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Desc       As String

        Dim GuildName  As String

        Dim errorStr   As String

        Dim Alineacion As Byte
        
        Desc = buffer.ReadASCIIString()
        GuildName = buffer.ReadASCIIString()
        Alineacion = buffer.ReadByte()
        
        If modGuilds.CrearNuevoClan(UserIndex, Desc, GuildName, Alineacion, errorStr) Then

            Call QuitarObjetos(407, 1, UserIndex)
            Call QuitarObjetos(408, 1, UserIndex)
            Call QuitarObjetos(409, 1, UserIndex)
            Call QuitarObjetos(411, 1, UserIndex)
            
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(.name & " fundó el clan " & GuildName & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
            'Update tag
            Call RefreshCharStatus(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim spellSlot As Byte

            Dim Spell     As Integer
        
108         spellSlot = .incomingData.ReadByte()
        
            'Validate slot
110         If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
112             Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate spell in the slot
114         Spell = .Stats.UserHechizos(spellSlot)

116         If Spell > 0 And Spell < NumeroHechizos + 1 Then

118             With Hechizos(Spell)
                    'Send information
120                 Call WriteConsoleMsg(UserIndex, "HECINF*" & Spell, FontTypeNames.FONTTYPE_INFO)

                End With

            End If

        End With

        
        Exit Sub

HandleSpellInfo_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleSpellInfo", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim itemSlot As Byte
        
108         itemSlot = .incomingData.ReadByte()
        
            'Dead users can't equip items
110         If .flags.Muerto = 1 Then
112             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Sólo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate item slot
114         If itemSlot > UserList(UserIndex).CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
116         If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
118         Call EquiparInvItem(UserIndex, itemSlot)

        End With

        
        Exit Sub

HandleEquipItem_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleEquipItem", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim heading As eHeading
        
108         heading = .incomingData.ReadByte()
        
            'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
110         If heading > 0 And heading < 5 Then
112             .Char.heading = heading
114             Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

            End If

        End With

        
        Exit Sub

HandleChangeHeading_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleChangeHeading", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 1 + NUMSKILLS Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim i                      As Long

            Dim Count                  As Integer

            Dim points(1 To NUMSKILLS) As Byte
        
            'Codigo para prevenir el hackeo de los skills
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
108         For i = 1 To NUMSKILLS
110             points(i) = .incomingData.ReadByte()
            
112             If points(i) < 0 Then
114                 Call LogHackAttemp(.name & " IP:" & .ip & " trató de hackear los skills.")
116                 .Stats.SkillPts = 0
118                 Call CloseSocket(UserIndex)
                    Exit Sub

                End If
            
120             Count = Count + points(i)
122         Next i
        
124         If Count > .Stats.SkillPts Then
126             Call LogHackAttemp(.name & " IP:" & .ip & " trató de hackear los skills.")
128             Call CloseSocket(UserIndex)
                Exit Sub

            End If

            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
130         With .Stats

132             For i = 1 To NUMSKILLS
134                 .SkillPts = .SkillPts - points(i)
136                 .UserSkills(i) = .UserSkills(i) + points(i)
                
                    'Client should prevent this, but just in case...
138                 If .UserSkills(i) > 100 Then
140                     .SkillPts = .SkillPts + .UserSkills(i) - 100
142                     .UserSkills(i) = 100

                    End If

144             Next i

            End With

        End With

        
        Exit Sub

HandleModifySkills_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleModifySkills", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim SpawnedNpc As Integer

            Dim PetIndex   As Byte
        
108         PetIndex = .incomingData.ReadByte()
        
110         If .flags.TargetNPC = 0 Then Exit Sub
        
112         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
114         If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
116             If PetIndex > 0 And PetIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                    'Create the creature
118                 SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(PetIndex).NpcIndex, Npclist(.flags.TargetNPC).Pos, True, False)
                
120                 If SpawnedNpc > 0 Then
122                     Npclist(SpawnedNpc).MaestroNPC = .flags.TargetNPC
124                     Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1

                    End If

                End If

            Else
126             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))

            End If

        End With

        
        Exit Sub

HandleTrain_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleTrain", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 4 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim slot   As Byte

            Dim Amount As Integer
        
108         slot = .incomingData.ReadByte()
110         Amount = .incomingData.ReadInteger()
        
            'Dead people can't commerce...
112         If .flags.Muerto = 1 Then
114             Call WriteConsoleMsg(UserIndex, "¡¡Estís muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'íEl target es un NPC valido?
116         If .flags.TargetNPC < 1 Then Exit Sub
            
            'íEl NPC puede comerciar?
118         If Npclist(.flags.TargetNPC).Comercia = 0 Then
120             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
                Exit Sub

            End If
        
            'Only if in commerce mode....
122         If Not .flags.Comerciando Then
124             Call WriteConsoleMsg(UserIndex, "No estás comerciando", FontTypeNames.FONTTYPE_INFO)
126             Call WriteCommerceEnd(UserIndex)
                Exit Sub

            End If
        
            'User compra el item
128         Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, slot, Amount)

        End With

        
        Exit Sub

HandleCommerceBuy_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCommerceBuy", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 5 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim slot        As Byte

            Dim slotdestino As Byte

            Dim Amount      As Integer
        
108         slot = .incomingData.ReadByte()
110         Amount = .incomingData.ReadInteger()
        
112         slotdestino = .incomingData.ReadByte()
        
            'Dead people can't commerce
114         If .flags.Muerto = 1 Then
116             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            '¿El target es un NPC valido?
118         If .flags.TargetNPC < 1 Then Exit Sub
        
            '¿Es el banquero?
120         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
                Exit Sub

            End If
        
            'User retira el item del slot
122         Call UserRetiraItem(UserIndex, slot, Amount, slotdestino)

        End With

        
        Exit Sub

HandleBankExtractItem_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleBankExtractItem", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 4 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim slot   As Byte

            Dim Amount As Integer
        
108         slot = .incomingData.ReadByte()
110         Amount = .incomingData.ReadInteger()
        
            'Dead people can't commerce...
112         If .flags.Muerto = 1 Then
114             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'íEl target es un NPC valido?
116         If .flags.TargetNPC < 1 Then Exit Sub
        
            'íEl NPC puede comerciar?
118         If Npclist(.flags.TargetNPC).Comercia = 0 Then
120             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
                Exit Sub

            End If
        
            'User compra el item del slot
122         Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, slot, Amount)

        End With

        
        Exit Sub

HandleCommerceSell_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCommerceSell", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 5 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim slot        As Byte

            Dim slotdestino As Byte

            Dim Amount      As Integer
        
108         slot = .incomingData.ReadByte()
110         Amount = .incomingData.ReadInteger()
112         slotdestino = .incomingData.ReadByte()
        
            'Dead people can't commerce...
114         If .flags.Muerto = 1 Then
116             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'íEl target es un NPC valido?
118         If .flags.TargetNPC < 1 Then Exit Sub
        
            'íEl NPC puede comerciar?
120         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
                Exit Sub

            End If
        
            'User deposita el item del slot rdata
122         Call UserDepositaItem(UserIndex, slot, Amount, slotdestino)

        End With

        
        Exit Sub

HandleBankDeposit_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleBankDeposit", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim File     As String

        Dim title    As String

        Dim msg      As String

        Dim postFile As String
        
        Dim handle   As Integer

        Dim i        As Long

        Dim Count    As Integer
        
        title = buffer.ReadASCIIString()
        msg = buffer.ReadASCIIString()
        
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 3 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex).incomingData
            'Remove packet ID
106         Call .ReadByte
        
            Dim dir As Integer
        
108         If .ReadBoolean() Then
110             dir = 1
            Else
112             dir = -1

            End If
        
114         Call DesplazarHechizo(UserIndex, dir, .ReadByte())

        End With

        
        Exit Sub

HandleMoveSpell_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleMoveSpell", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Desc As String
        
        Desc = buffer.ReadASCIIString()
        
        Call modGuilds.ChangeCodexAndDesc(Desc, .GuildIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 6 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Amount As Long

            Dim slot   As Byte

            Dim tUser  As Integer
        
108         slot = .incomingData.ReadByte()
110         Amount = .incomingData.ReadLong()
        
            'Get the other player
112         tUser = .ComUsu.DestUsu
        
            'If Amount is invalid, or slot is invalid and it's not gold, then ignore it.
114         If ((slot < 1 Or slot > UserList(UserIndex).CurrentInventorySlots) And slot <> FLAGORO) Or Amount <= 0 Then Exit Sub
        
            'Is the other player valid??
116         If tUser < 1 Or tUser > MaxUsers Then Exit Sub
        
            'Is the commerce attempt valid??
118         If UserList(tUser).ComUsu.DestUsu <> UserIndex Then
120             Call FinComerciarUsu(UserIndex)
                Exit Sub

            End If
        
            'Is he still logged??
122         If Not UserList(tUser).flags.UserLogged Then
124             Call FinComerciarUsu(UserIndex)
                Exit Sub
            Else

                'Is he alive??
126             If UserList(tUser).flags.Muerto = 1 Then
128                 Call FinComerciarUsu(UserIndex)
                    Exit Sub

                End If
            
                'Has he got enough??
130             If slot = FLAGORO Then

                    'gold
132                 If Amount > .Stats.GLD Then
134                     Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                        Exit Sub

                    End If

                Else

                    'inventory
136                 If Amount > .Invent.Object(slot).Amount Then
138                     Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                        Exit Sub

                    End If

                End If
            
                'Prevent offer changes (otherwise people would ripp off other players)
140             If .ComUsu.Objeto > 0 Then
142                 Call WriteConsoleMsg(UserIndex, "No podés cambiar tu oferta.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub

                End If
            
                'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
144             If .flags.Navegando = 1 Then
146                 If .Invent.BarcoSlot = slot Then
148                     Call WriteConsoleMsg(UserIndex, "No podés vender tu barco mientras lo estás usando.", FontTypeNames.FONTTYPE_TALK)
                        Exit Sub

                    End If

                End If
            
150             If .flags.Montado = 1 Then
152                 If .Invent.MonturaSlot = slot Then
154                     Call WriteConsoleMsg(UserIndex, "No podés vender tu montura mientras la estás usando.", FontTypeNames.FONTTYPE_TALK)
                        Exit Sub

                    End If

                End If
            
156             .ComUsu.Objeto = slot
158             .ComUsu.cant = Amount
            
                'If the other one had accepted, we turn that back and inform of the new offer (just to be cautious).
160             If UserList(tUser).ComUsu.Acepto = True Then
162                 UserList(tUser).ComUsu.Acepto = False
164                 Call WriteConsoleMsg(tUser, .name & " ha cambiado su oferta.", FontTypeNames.FONTTYPE_TALK)

                End If
            
166             Call EnviarObjetoTransaccion(tUser)

            End If

        End With

        
        Exit Sub

HandleUserCommerceOffer_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleUserCommerceOffer", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild    As String

        Dim proposal As String

        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de paz enviada", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild    As String

        Dim proposal As String

        Dim errorStr As String
        
        guild = buffer.ReadASCIIString()
        proposal = buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de alianza enviada", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild    As String

        Dim errorStr As String

        Dim details  As String
        
        guild = buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.ALIADOS, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild    As String

        Dim errorStr As String

        Dim details  As String
        
        guild = buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, guild, RELACIONES_GUILD.PAZ, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim user    As String

        Dim details As String
        
        user = buffer.ReadASCIIString()
        
        details = modGuilds.a_DetallesAspirante(UserIndex, user)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje no ha mandado solicitud, o no estás habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(UserIndex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
        'Remove packet ID
        
        On Error GoTo HandleGuildAlliancePropList_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
102     Call WriteAlianceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.ALIADOS))

        
        Exit Sub

HandleGuildAlliancePropList_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleGuildAlliancePropList", Erl)
        Resume Next
        
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
        'Remove packet ID
        
        On Error GoTo HandleGuildPeacePropList_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
102     Call WritePeaceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.PAZ))

        
        Exit Sub

HandleGuildPeacePropList_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleGuildPeacePropList", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild           As String

        Dim errorStr        As String

        Dim otherGuildIndex As Integer
        
        guild = buffer.ReadASCIIString()
        
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarWebSite(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String

        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If Not modGuilds.a_AceptarAspirante(UserIndex, UserName, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(UserName)

            If tUser > 0 And UserList(tUser).flags.BattleModo = 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
                Call RefreshCharStatus(tUser)

            End If
            
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("[" & UserName & "] ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String

        Dim UserName As String

        Dim Reason   As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName   As String

        Dim GuildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        
        GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
        
        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        Else
            Call WriteConsoleMsg(UserIndex, "No podés expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarNoticias(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendDetallesPersonaje(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            Dim Error As String
        
104         If Not modGuilds.v_AbrirElecciones(UserIndex, Error) Then
106             Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)
            Else
108             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .name, FontTypeNames.FONTTYPE_GUILD))

            End If

        End With

        
        Exit Sub

HandleGuildOpenElections_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleGuildOpenElections", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild       As String

        Dim application As String

        Dim errorStr    As String
        
        guild = buffer.ReadASCIIString()
        application = buffer.ReadASCIIString()
        
        If Not modGuilds.a_NuevoAspirante(UserIndex, guild, application, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & guild & ".", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendGuildDetails(UserIndex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

''
' Handles the "Online" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal UserIndex As Integer)
        
        On Error GoTo HandleOnline_Err
        

        '***************************************************
        Dim i     As Long

        Dim Count As Long
    
100     With UserList(UserIndex)
            'Remove packet ID
102         Call .incomingData.ReadByte

            Dim nombres As String
        
104         For i = 1 To LastUser

106             If UserList(i).flags.UserLogged Then
                    If UserList(i).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then
                        nombres = nombres & " - " & UserList(i).name
                    End If
108                 Count = Count + 1

                End If

110         Next i


            If .flags.Privilegios And PlayerType.user Then
112             Call WriteConsoleMsg(UserIndex, "Número de usuarios: " & CStr(Count) & " conectados.", FontTypeNames.FONTTYPE_INFOIAO)
            Else
                Call WriteConsoleMsg(UserIndex, "Número de usuarios: " & CStr(Count) & " conectados: " & nombres & ".", FontTypeNames.FONTTYPE_INFOIAO)
            End If

        End With

        
        Exit Sub

HandleOnline_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleOnline", Erl)
        Resume Next
        
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
        '04/15/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
        '***************************************************
        Dim tUser        As Integer

        Dim isNotVisible As Boolean
    
100     With UserList(UserIndex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Paralizado = 1 Then
106             Call WriteConsoleMsg(UserIndex, "No podés salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
        
            'exit secure commerce
108         If .ComUsu.DestUsu > 0 Then
110             tUser = .ComUsu.DestUsu
            
112             If UserList(tUser).flags.UserLogged Then
114                 If UserList(tUser).ComUsu.DestUsu = UserIndex Then
116                     Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
118                     Call FinComerciarUsu(tUser)

                    End If

                End If
            
120             Call WriteConsoleMsg(UserIndex, "Comercio cancelado. ", FontTypeNames.FONTTYPE_TALK)
122             Call FinComerciarUsu(UserIndex)

            End If
        
124         isNotVisible = (.flags.Oculto Or .flags.invisible)

126         If isNotVisible Then
128             .flags.Oculto = 0
130             .flags.invisible = 0
    
132             .Counters.Invisibilidad = 0
134             .Counters.TiempoOculto = 0
                'Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
136             Call WriteLocaleMsg(UserIndex, "307", FontTypeNames.FONTTYPE_INFO)
138             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))

            End If
        
            Rem   Call WritePersonajesDeCuenta(UserIndex, .Cuenta)
140         Call Cerrar_Usuario(UserIndex)

        End With

        
        Exit Sub

HandleQuit_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleQuit", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'obtengo el guildindex
104         GuildIndex = m_EcharMiembroDeClan(UserIndex, .name)
        
106         If GuildIndex > 0 Then
108             Call WriteConsoleMsg(UserIndex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
110             Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
            Else
112             Call WriteConsoleMsg(UserIndex, "Tu no podés salir de ningún clan.", FontTypeNames.FONTTYPE_GUILD)

            End If

        End With

        
        Exit Sub

HandleGuildLeave_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleGuildLeave", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Dead people can't check their accounts
104         If .flags.Muerto = 1 Then
106             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
108         If .flags.TargetNPC = 0 Then
110             Call WriteConsoleMsg(UserIndex, "Primero tenes que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
112         If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
114             Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
116         Select Case Npclist(.flags.TargetNPC).NPCtype

                Case eNPCType.Banquero
118                 Call WriteChatOverHead(UserIndex, "Tenes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
120             Case eNPCType.Timbero

122                 If Not .flags.Privilegios And PlayerType.user Then
124                     earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
126                     If earnings >= 0 And Apuestas.Ganancias <> 0 Then
128                         percentage = Int(earnings * 100 / Apuestas.Ganancias)

                        End If
                    
130                     If earnings < 0 And Apuestas.Perdidas <> 0 Then
132                         percentage = Int(earnings * 100 / Apuestas.Perdidas)

                        End If
                    
134                     Call WriteConsoleMsg(UserIndex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)

                    End If

            End Select

        End With

        
        Exit Sub

HandleRequestAccountState_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestAccountState", Erl)
        Resume Next
        
End Sub

''
' Handles the "PetStand" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal UserIndex As Integer)
        
        On Error GoTo HandlePetStand_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Dead people can't use pets
104         If .flags.Muerto = 1 Then
106             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
108         If .flags.TargetNPC = 0 Then
110             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make sure it's close enough
112         If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
114             Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Do it!
116         Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO
        
118         Call Expresar(.flags.TargetNPC, UserIndex)

        End With

        
        Exit Sub

HandlePetStand_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandlePetStand", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        
        chat = buffer.ReadASCIIString()
        
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Dead users can't use pets
104         If .flags.Muerto = 1 Then
106             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
108         If .flags.TargetNPC = 0 Then
110             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make sure it's close enough
112         If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
114             Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make sure it's the trainer
116         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
118         Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC)

        End With

        
        Exit Sub

HandleTrainList_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleTrainList", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Dead users can't use pets
104         If .flags.Muerto = 1 Then
106             Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Solo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If HayOBJarea(.Pos, FOGATA) Then
110             Call WriteRestOK(UserIndex)
            
112             If Not .flags.Descansar Then
114                 Call WriteConsoleMsg(UserIndex, "Te acomodás junto a la fogata y comenzís a descansar.", FontTypeNames.FONTTYPE_INFO)
                Else
116                 Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)

                End If
            
118             .flags.Descansar = Not .flags.Descansar
            Else

120             If .flags.Descansar Then
122                 Call WriteRestOK(UserIndex)
124                 Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
126                 .flags.Descansar = False
                    Exit Sub

                End If
            
128             Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        
        Exit Sub

HandleRest_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRest", Erl)
        Resume Next
        
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

            'Remove packet ID
102         Call .incomingData.ReadByte
            
            'Si ya tiene el mana completo, no lo dejamos meditar.
104         If .Stats.MinMAN = .Stats.MaxMAN Then Exit Sub
                           
            'Las clases NO MAGICAS no meditan...
106         If .clase = eClass.Hunter Or _
               .clase = eClass.Trabajador Or _
               .clase = eClass.Warrior Then Exit Sub

108         If .flags.Muerto = 1 Then
110             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
112         If .flags.Montado = 1 Then
114             Call WriteConsoleMsg(UserIndex, "No podes meditar estando montado.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

116         .flags.Meditando = Not .flags.Meditando

118         If .flags.Meditando Then

                .Counters.InicioMeditar = GetTickCount And &H7FFFFFFF

120             Select Case .Stats.ELV

                    Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14
122                     .Char.FX = Meditaciones.MeditarInicial

124                 Case 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29
126                     .Char.FX = Meditaciones.MeditarMayor15

128                 Case 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44
130                     .Char.FX = Meditaciones.MeditarMayor30

132                 Case Else
134                     .Char.FX = Meditaciones.MeditarMayor45

                End Select
            Else
136             .Char.FX = 0
                'Call WriteLocaleMsg(UserIndex, "123", FontTypeNames.FONTTYPE_INFO)
            End If

140         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.CharIndex, .Char.FX))

        End With

        
        Exit Sub

HandleMeditate_Err:
142     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleMeditate", Erl)
144     Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Se asegura que el target es un npc
104         If .flags.TargetNPC = 0 Then
106             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate NPC and make sure player is dead
108         If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 0 Then Exit Sub
        
            'Make sure it's close enough
110         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
112             Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         Call RevivirUsuario(UserIndex)
116         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Curar, 100, False))
118         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("104", UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
120         Call WriteConsoleMsg(UserIndex, "ííHís sido resucitado!!", FontTypeNames.FONTTYPE_INFO)

        End With

        
        Exit Sub

HandleResucitate_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleResucitate", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Se asegura que el target es un npc
104         If .flags.TargetNPC = 0 Then
106             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie) Or .flags.Muerto <> 0 Then Exit Sub
        
110         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
112             Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         .Stats.MinHp = .Stats.MaxHp
        
116         Call WriteUpdateHP(UserIndex)
        
118         Call WriteConsoleMsg(UserIndex, "ííHís sido curado!!", FontTypeNames.FONTTYPE_INFO)

        End With

        
        Exit Sub

HandleHeal_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleHeal", Erl)
        Resume Next
        
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
        'Remove packet ID
        
        On Error GoTo HandleRequestStats_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
102     Call SendUserStatsTxt(UserIndex, UserIndex)

        
        Exit Sub

HandleRequestStats_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestStats", Erl)
        Resume Next
        
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
        'Remove packet ID
        
        On Error GoTo HandleHelp_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
102     Call SendHelp(UserIndex)

        
        Exit Sub

HandleHelp_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleHelp", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Dead people can't commerce
104         If .flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
106             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Is it already in commerce mode??
108         If .flags.Comerciando Then
110             Call WriteConsoleMsg(UserIndex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
112         If .flags.TargetNPC > 0 Then

                'Does the NPC want to trade??
114             If Npclist(.flags.TargetNPC).Comercia = 0 Then
116                 If LenB(Npclist(.flags.TargetNPC).Desc) <> 0 Then
118                     Call WriteChatOverHead(UserIndex, "No tengo ningún interís en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

                    End If
                
                    Exit Sub

                End If
            
120             If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
122                 Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Start commerce....
124             Call IniciarComercioNPC(UserIndex)
                '[Alejo]
126         ElseIf .flags.TargetUser > 0 Then

                'User commerce...
                'Can he commerce??
128             If .flags.Privilegios And PlayerType.Consejero Then
130                 Call WriteConsoleMsg(UserIndex, "No podés vender items.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub

                End If
            
                'Is the other one dead??
132             If UserList(.flags.TargetUser).flags.Muerto = 1 Then
134                 Call WriteConsoleMsg(UserIndex, "í¡No podés comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Is it me??
136             If .flags.TargetUser = UserIndex Then
138                 Call WriteConsoleMsg(UserIndex, "No podés comerciar con vos mismo...", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Check distance
140             If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
142                 Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Is he already trading?? is it with me or someone else??
144             If UserList(.flags.TargetUser).flags.Comerciando = True And UserList(.flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
146                 Call WriteConsoleMsg(UserIndex, "No podés comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Initialize some variables...
148             .ComUsu.DestUsu = .flags.TargetUser
150             .ComUsu.DestNick = UserList(.flags.TargetUser).name
152             .ComUsu.cant = 0
154             .ComUsu.Objeto = 0
156             .ComUsu.Acepto = False
            
                'Rutina para comerciar con otro usuario
158             Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)
            Else
160             Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        
        Exit Sub

HandleCommerceStart_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCommerceStart", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Dead people can't commerce
104         If .flags.Muerto = 1 Then
106             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If .flags.Comerciando Then
110             Call WriteConsoleMsg(UserIndex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
112         If .flags.TargetNPC > 0 Then
114             If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 6 Then
116                 Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'If it's the banker....
118             If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
120                 Call IniciarDeposito(UserIndex)

                End If

            Else
122             Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        
        Exit Sub

HandleBankStart_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleBankStart", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Validate target NPC
104         If .flags.TargetNPC = 0 Then
106             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
110         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
112             Call WriteConsoleMsg(UserIndex, "Debes acercarte mís.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
116             Call EnlistarArmadaReal(UserIndex)
            Else
118             Call EnlistarCaos(UserIndex)

            End If

        End With

        
        Exit Sub

HandleEnlist_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleEnlist", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Validate target NPC
104         If .flags.TargetNPC = 0 Then
106             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
110         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
112             Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
116             If .Faccion.ArmadaReal = 0 Then
118                 Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

120             Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, cada 100 criminales que derrotes te darí una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else

122             If .Faccion.FuerzasCaos = 0 Then
124                 Call WriteChatOverHead(UserIndex, "No perteneces a la legiín oscura!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

126             Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te darí una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

            End If

        End With

        
        Exit Sub

HandleInformation_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleInformation", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Validate target NPC
104         If .flags.TargetNPC = 0 Then
106             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
110         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
112             Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
116             If .Faccion.ArmadaReal = 0 Then
118                 Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

120             Call RecompensaArmadaReal(UserIndex)
            Else

122             If .Faccion.FuerzasCaos = 0 Then
124                 Call WriteChatOverHead(UserIndex, "No perteneces a la legiín oscura!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

126             Call RecompensaCaos(UserIndex)

            End If

        End With

        
        Exit Sub

HandleReward_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleReward", Erl)
        Resume Next
        
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
        'Remove packet ID
        
        On Error GoTo HandleRequestMOTD_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
102     Call SendMOTD(UserIndex)

        
        Exit Sub

HandleRequestMOTD_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestMOTD", Erl)
        Resume Next
        
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
        'Remove packet ID
        
        On Error GoTo HandleUpTime_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
        Dim Time      As Long

        Dim UpTimeStr As String
    
        'Get total time in seconds
102     Time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000
    
        'Get times in dd:hh:mm:ss format
104     UpTimeStr = (Time Mod 60) & " segundos."
106     Time = Time \ 60
    
108     UpTimeStr = (Time Mod 60) & " minutos, " & UpTimeStr
110     Time = Time \ 60
    
112     UpTimeStr = (Time Mod 24) & " horas, " & UpTimeStr
114     Time = Time \ 24
    
116     If Time = 1 Then
118         UpTimeStr = Time & " día, " & UpTimeStr
        Else
120         UpTimeStr = Time & " días, " & UpTimeStr

        End If
    
122     Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)

        
        Exit Sub

HandleUpTime_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleUpTime", Erl)
        Resume Next
        
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
        'Remove packet ID
        
        On Error GoTo HandleInquiry_Err
        
100     Call UserList(UserIndex).incomingData.ReadByte
    
102     ConsultaPopular.SendInfoEncuesta (UserIndex)

        
        Exit Sub

HandleInquiry_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleInquiry", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        
        chat = buffer.ReadASCIIString()
        
        If LenB(chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(chat)
            
            If .GuildIndex > 0 Then
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.name & "> " & chat))

                'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
                'Call SendData(SendTarget.ToClanArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "í< " & rData & " >í" & CStr(UserList(UserIndex).Char.CharIndex))
            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 3 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
108         Call CentinelaCheckClave(UserIndex, .incomingData.ReadInteger())

        End With

        
        Exit Sub

HandleCentinelReport_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCentinelReport", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            Dim onlineList As String
        
104         onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .GuildIndex)
        
106         If .GuildIndex <> 0 Then
108             Call WriteConsoleMsg(UserIndex, "Compaíeros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
            Else
110             Call WriteConsoleMsg(UserIndex, "No pertences a ningún clan.", FontTypeNames.FONTTYPE_GUILDMSG)

            End If

        End With

        
        Exit Sub

HandleGuildOnline_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleGuildOnline", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        
        chat = buffer.ReadASCIIString()
        
        If LenB(chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(chat)
            
            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim request As String
        
        request = buffer.ReadASCIIString()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     With UserList(UserIndex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If Not Ayuda.Existe(.name) Then
106             Call WriteConsoleMsg(UserIndex, "El mensaje ha sido entregado, ahora sílo debes esperar que se desocupe algín GM.", FontTypeNames.FONTTYPE_INFO)
                'Call Ayuda.Push(.name)
            Else
                'Call Ayuda.Quitar(.name)
                'Call Ayuda.Push(.name)
108             Call WriteConsoleMsg(UserIndex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        
        Exit Sub

HandleGMRequest_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleGMRequest", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim description As String
        
        description = buffer.ReadASCIIString()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No podés cambiar la descripciín estando muerto.", FontTypeNames.FONTTYPE_INFOIAO)
        Else

            If Not AsciiValidos(description) Then
                Call WriteConsoleMsg(UserIndex, "La descripciín tiene caractíres invílidos.", FontTypeNames.FONTTYPE_INFOIAO)
            Else
                .Desc = Trim$(description)
                Call WriteConsoleMsg(UserIndex, "La descripciín a cambiado.", FontTypeNames.FONTTYPE_INFOIAO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim vote     As String

        Dim errorStr As String
        
        vote = buffer.ReadASCIIString()
        
        If Not modGuilds.v_UsuarioVota(UserIndex, vote, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim name  As String

        Dim Count As Integer
        
        name = buffer.ReadASCIIString()
        
        If LenB(name) <> 0 Then
            If (InStrB(name, "\") <> 0) Then
                name = Replace(name, "\", "")

            End If

            If (InStrB(name, "/") <> 0) Then
                name = Replace(name, "/", "")

            End If

            If (InStrB(name, ":") <> 0) Then
                name = Replace(name, ":", "")

            End If

            If (InStrB(name, "|") <> 0) Then
                name = Replace(name, "|", "")

            End If
            
            If FileExist(CharPath & name & ".chr", vbNormal) Then
                Count = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))

                If Count = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                Else

                    While Count > 0

                        Call WriteConsoleMsg(UserIndex, Count & " - " & GetVar(CharPath & name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
                        Count = Count - 1
                    Wend

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "Personaje """ & name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        Dim oldPass  As String

        Dim newPass  As String

        Dim oldPass2 As String
        
        'Remove packet ID
        Call buffer.ReadByte
        
        oldPass = buffer.ReadASCIIString()
        newPass = buffer.ReadASCIIString()

        If Database_Enabled Then
            Call ChangePasswordDatabase(UserIndex, SDesencriptar(oldPass), SDesencriptar(newPass))
        
        Else

            If LenB(SDesencriptar(newPass)) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Debe especificar una contraseña nueva, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Else
                oldPass2 = GetVar(CuentasPath & UserList(UserIndex).Cuenta & ".act", "INIT", "PASSWORD")
                
                If SDesencriptar(oldPass2) <> SDesencriptar(oldPass) Then
                    Call WriteConsoleMsg(UserIndex, "La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteVar(CuentasPath & UserList(UserIndex).Cuenta & ".act", "INIT", "PASSWORD", newPass)
                    Call WriteConsoleMsg(UserIndex, "La contraseña de su cuenta fue cambiada con éxito.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 3 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Amount As Integer
        
108         Amount = .incomingData.ReadInteger()
        
110         If .flags.Muerto = 1 Then
112             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
114         ElseIf .flags.TargetNPC = 0 Then
                'Validate target NPC
116             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
118         ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
120             Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                ' Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
122         ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
124             Call WriteChatOverHead(UserIndex, "No tengo ningún interís en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
126         ElseIf Amount < 1 Then
128             Call WriteChatOverHead(UserIndex, "El mínimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
130         ElseIf Amount > 10000 Then
132             Call WriteChatOverHead(UserIndex, "El míximo de apuesta es 10000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
134         ElseIf .Stats.GLD < Amount Then
136             Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else

138             If RandomNumber(1, 100) <= 45 Then
140                 .Stats.GLD = .Stats.GLD + Amount
142                 Call WriteChatOverHead(UserIndex, "Felicidades! Has ganado " & CStr(Amount) & " monedas de oro!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
144                 Apuestas.Perdidas = Apuestas.Perdidas + Amount
146                 Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
148                 .Stats.GLD = .Stats.GLD - Amount
150                 Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
152                 Apuestas.Ganancias = Apuestas.Ganancias + Amount
154                 Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))

                End If
            
156             Apuestas.Jugadas = Apuestas.Jugadas + 1
            
158             Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
160             Call WriteUpdateGold(UserIndex)

            End If

        End With

        
        Exit Sub

HandleGamble_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleGamble", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim opt As Byte
        
108         opt = .incomingData.ReadByte()
        
110         Call WriteConsoleMsg(UserIndex, ConsultaPopular.doVotar(UserIndex, opt), FontTypeNames.FONTTYPE_GUILD)

        End With

        
        Exit Sub

HandleInquiryVote_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleInquiryVote", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 5 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Amount As Long
        
108         Amount = .incomingData.ReadLong()
        
            'Dead people can't leave a faction.. they can't talk...
110         If .flags.Muerto = 1 Then
112             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
114         If .flags.TargetNPC = 0 Then
116             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
118         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
120         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
122             Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
124         If Amount > 0 And Amount <= .Stats.Banco Then
126             .Stats.Banco = .Stats.Banco - Amount
128             .Stats.GLD = .Stats.GLD + Amount
130             'Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

                Call WriteUpdateGold(UserIndex)
                Call WriteGoliathInit(UserIndex)
            Else
132             Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

            End If

        End With

        
        Exit Sub

HandleBankExtractGold_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleBankExtractGold", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Dead people can't leave a faction.. they can't talk...
104         If .flags.Muerto = 1 Then
106             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If .Faccion.ArmadaReal = 0 And .Faccion.FuerzasCaos = 0 Then
110             If .Faccion.Status = 1 Then
112                 Call VolverCriminal(UserIndex)
114                 Call WriteConsoleMsg(UserIndex, "Ahora sos un criminal.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If

            Else

                ' Call WriteConsoleMsg(UserIndex, "Ya sos un criminal.", FontTypeNames.FONTTYPE_INFOIAO)
                ' Exit Sub
            End If
        
            'Validate target NPC
116         If .flags.TargetNPC = 0 Then
118             If .Faccion.ArmadaReal = 1 Then
120                 Call WriteConsoleMsg(UserIndex, "Para salir del ejercito debes ir a visitar al rey.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
122             ElseIf .Faccion.FuerzasCaos = 1 Then
124                 Call WriteConsoleMsg(UserIndex, "Para salir de la legion debes ir a visitar al diablo.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If

            End If
        
126         If .flags.TargetNPC = 0 Then Exit Sub
        
128         If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Enlistador Then

                'Quit the Royal Army?
130             If .Faccion.ArmadaReal = 1 Then
132                 If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
134                     Call ExpulsarFaccionReal(UserIndex)
136                     Call WriteChatOverHead(UserIndex, "Serís bienvenido a las fuerzas imperiales si deseas regresar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                        Exit Sub
                    Else
138                     Call WriteChatOverHead(UserIndex, "¡¡¡Sal de aquí bufón!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                   
                    End If

                    'Quit the Chaos Legion??
140             ElseIf .Faccion.FuerzasCaos = 1 Then

142                 If Npclist(.flags.TargetNPC).flags.Faccion = 1 Then
144                     Call ExpulsarFaccionCaos(UserIndex)
146                     Call WriteChatOverHead(UserIndex, "Ya volverís arrastrandote.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Else
148                     Call WriteChatOverHead(UserIndex, "Sal de aquí maldito criminal", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

                    End If

                Else
150                 Call WriteChatOverHead(UserIndex, "¡No perteneces a ninguna facción!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

                End If

            End If
    
        End With

        
        Exit Sub

HandleLeaveFaction_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleLeaveFaction", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 5 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Amount As Long
        
108         Amount = .incomingData.ReadLong()
        
            'Dead people can't leave a faction.. they can't talk...
110         If .flags.Muerto = 1 Then
112             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
114         If .flags.TargetNPC = 0 Then
116             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
118         If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
120             Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
122         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
124         If Amount > 0 And Amount <= .Stats.GLD Then
126             .Stats.Banco = .Stats.Banco + Amount
128             .Stats.GLD = .Stats.GLD - Amount
130             'Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
132             Call WriteUpdateGold(UserIndex)
                Call WriteGoliathInit(UserIndex)
            Else
134             Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

            End If

        End With

        
        Exit Sub

HandleBankDepositGold_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleBankDepositGold", Erl)
        Resume Next
        
End Sub

''
' Handles the "Denounce" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)
        
        On Error GoTo HandleDenounce_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)
            'Remove packet ID
102         Call .incomingData.ReadByte

104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
106         If EventoActivo Then
108             Call FinalizarEvento
            Else
110             Call WriteConsoleMsg(UserIndex, "No hay ningun evento activo.", FontTypeNames.FONTTYPE_New_Eventos)
        
            End If
        
        End With

        
        Exit Sub

HandleDenounce_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleDenounce", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild       As String

        Dim memberCount As Integer

        Dim i           As Long

        Dim UserName    As String
        
        guild = buffer.ReadASCIIString()
        
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String
        
        Message = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.user Then
            Call LogGM(.name, "Mensaje a Gms:" & Message)
        
            If LenB(Message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(Message)
            
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & "> " & Message, FontTypeNames.FONTTYPE_GMMSG))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
106             .showName = Not .showName 'Show / Hide the name
            
108             Call RefreshCharStatus(UserIndex)

            End If

        End With

        
        Exit Sub

HandleShowName_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleShowName", Erl)
        Resume Next
        
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
            'Remove packet ID
102         .incomingData.ReadByte
        
104         If .flags.Privilegios And PlayerType.user Then Exit Sub
    
            Dim i    As Long

            Dim list As String

106         For i = 1 To LastUser

108             If UserList(i).ConnID <> -1 Then
110                 If UserList(i).Faccion.ArmadaReal = 1 Then
112                     If UserList(i).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Or .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
114                         list = list & UserList(i).name & ", "

                        End If

                    End If

                End If

116         Next i

        End With
    
118     If Len(list) > 0 Then
120         Call WriteConsoleMsg(UserIndex, "Armadas conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
        Else
122         Call WriteConsoleMsg(UserIndex, "No hay Armadas conectados", FontTypeNames.FONTTYPE_INFO)

        End If

        
        Exit Sub

HandleOnlineRoyalArmy_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleOnlineRoyalArmy", Erl)
        Resume Next
        
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
            'Remove packet ID
102         .incomingData.ReadByte
        
104         If .flags.Privilegios And PlayerType.user Then Exit Sub
    
            Dim i    As Long

            Dim list As String

106         For i = 1 To LastUser

108             If UserList(i).ConnID <> -1 Then
110                 If UserList(i).Faccion.FuerzasCaos = 1 Then
112                     If UserList(i).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Or .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
114                         list = list & UserList(i).name & ", "

                        End If

                    End If

                End If

116         Next i

        End With

118     If Len(list) > 0 Then
120         Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
        Else
122         Call WriteConsoleMsg(UserIndex, "No hay Caos conectados", FontTypeNames.FONTTYPE_INFO)

        End If

        
        Exit Sub

HandleOnlineChaosLegion_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleOnlineChaosLegion", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        
        UserName = buffer.ReadASCIIString()
        
        Dim tIndex As Integer

        Dim x      As Long

        Dim Y      As Long

        Dim i      As Long

        Dim found  As Boolean
        
        tIndex = NameIndex(UserName)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then

            'Si es dios o Admins no podemos salvo que nosotros tambiín lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex <= 0 Then 'existe el usuario destino?
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else

                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For x = UserList(tIndex).Pos.x - i To UserList(tIndex).Pos.x + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i

                                If MapData(UserList(tIndex).Pos.Map, x, Y).UserIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.Map, x, Y, True, True) Then
                                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, x, Y, True)
                                        found = True
                                        Exit For

                                    End If

                                End If

                            Next Y
                            
                            If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next x
                        
                        If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next i
                    
                    'No space found??
                    If Not found Then
                        Call WriteConsoleMsg(UserIndex, "Todos los lugares estín ocupados.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim comment As String

        comment = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.user Then
            Call LogGM(.name, "Comentario: " & comment)
            Call WriteConsoleMsg(UserIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
            'Remove packet ID
102         Call .incomingData.ReadByte
    
104         If .flags.Privilegios And PlayerType.user Then Exit Sub
    
106         Call LogGM(.name, "Hora.")

        End With
    
108     Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & Time & " " & Date, FontTypeNames.FONTTYPE_INFO))

        
        Exit Sub

HandleServerTime_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleServerTime", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.user Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                If (UserList(tUser).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(UserIndex, "Ubicaciín  " & UserName & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.x & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "/Donde " & UserName)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 3 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Map As Integer

            Dim i, j As Long

            Dim NPCcount1, NPCcount2 As Integer

            Dim NPCcant1() As Integer

            Dim NPCcant2() As Integer

            Dim List1()    As String

            Dim List2()    As String
        
108         Map = .incomingData.ReadInteger()
        
110         If .flags.Privilegios And PlayerType.user Then Exit Sub
        
112         If MapaValido(Map) Then

114             For i = 1 To LastNPC

                    'VB isn't lazzy, so we put more restrictive condition first to speed up the process
116                 If Npclist(i).Pos.Map = Map Then

                        'íesta vivo?
118                     If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).Stats.Alineacion = 2 Then
120                         If NPCcount1 = 0 Then
122                             ReDim List1(0) As String
124                             ReDim NPCcant1(0) As Integer
126                             NPCcount1 = 1
128                             List1(0) = Npclist(i).name & ": (" & Npclist(i).Pos.x & "," & Npclist(i).Pos.Y & ")"
130                             NPCcant1(0) = 1
                            Else

132                             For j = 0 To NPCcount1 - 1

134                                 If Left$(List1(j), Len(Npclist(i).name)) = Npclist(i).name Then
136                                     List1(j) = List1(j) & ", (" & Npclist(i).Pos.x & "," & Npclist(i).Pos.Y & ")"
138                                     NPCcant1(j) = NPCcant1(j) + 1
                                        Exit For

                                    End If

140                             Next j

142                             If j = NPCcount1 Then
144                                 ReDim Preserve List1(0 To NPCcount1) As String
146                                 ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
148                                 NPCcount1 = NPCcount1 + 1
150                                 List1(j) = Npclist(i).name & ": (" & Npclist(i).Pos.x & "," & Npclist(i).Pos.Y & ")"
152                                 NPCcant1(j) = 1

                                End If

                            End If

                        Else

154                         If NPCcount2 = 0 Then
156                             ReDim List2(0) As String
158                             ReDim NPCcant2(0) As Integer
160                             NPCcount2 = 1
162                             List2(0) = Npclist(i).name & ": (" & Npclist(i).Pos.x & "," & Npclist(i).Pos.Y & ")"
164                             NPCcant2(0) = 1
                            Else

166                             For j = 0 To NPCcount2 - 1

168                                 If Left$(List2(j), Len(Npclist(i).name)) = Npclist(i).name Then
170                                     List2(j) = List2(j) & ", (" & Npclist(i).Pos.x & "," & Npclist(i).Pos.Y & ")"
172                                     NPCcant2(j) = NPCcant2(j) + 1
                                        Exit For

                                    End If

174                             Next j

176                             If j = NPCcount2 Then
178                                 ReDim Preserve List2(0 To NPCcount2) As String
180                                 ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
182                                 NPCcount2 = NPCcount2 + 1
184                                 List2(j) = Npclist(i).name & ": (" & Npclist(i).Pos.x & "," & Npclist(i).Pos.Y & ")"
186                                 NPCcant2(j) = 1

                                End If

                            End If

                        End If

                    End If

188             Next i
            
190             Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)

192             If NPCcount1 = 0 Then
194                 Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles", FontTypeNames.FONTTYPE_INFO)
                Else

196                 For j = 0 To NPCcount1 - 1
198                     Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
200                 Next j

                End If

202             Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)

204             If NPCcount2 = 0 Then
206                 Call WriteConsoleMsg(UserIndex, "No hay mís NPCS", FontTypeNames.FONTTYPE_INFO)
                Else

208                 For j = 0 To NPCcount2 - 1
210                     Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
212                 Next j

                End If

214             Call LogGM(.name, "Numero enemigos en mapa " & Map)

            End If

        End With

        
        Exit Sub

HandleCreaturesInMap_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCreaturesInMap", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And PlayerType.user Then Exit Sub
        
106         Call WarpUserChar(UserIndex, .flags.TargetMap, .flags.TargetX, .flags.TargetY, True)
108         Call LogGM(.name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.Map)

        End With

        
        Exit Sub

HandleWarpMeToTarget_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleWarpMeToTarget", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 7 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim Map      As Integer

        Dim x        As Byte

        Dim Y        As Byte

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        Map = buffer.ReadInteger()
        x = buffer.ReadByte()
        Y = buffer.ReadByte()
        
        If Not .flags.Privilegios And PlayerType.user Then
            If MapaValido(Map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)

                    End If

                Else
                    tUser = UserIndex

                End If
            
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                ElseIf InMapBounds(Map, x, Y) Then
                    Call FindLegalPos(tUser, Map, x, Y)
                    Call WarpUserChar(tUser, Map, x, Y, True)
                    Call WriteConsoleMsg(UserIndex, UserList(tUser).name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                    If tUser <> UserIndex Then Call LogGM(.name, "Transportí a " & UserList(tUser).name & " hacia " & "Mapa" & Map & " X:" & x & " Y:" & Y)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.user Then
            tUser = NameIndex(UserName)
        
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(UserIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "ESTIMADO USUARIO, ud ha sido silenciado por los administradores. Sus denuncias serín ignoradas por el servidor de aquí en mís. Utilice /GM para contactar un administrador.")
                    Call LogGM(.name, "/silenciar " & UserList(tUser).name)
                
                    'Flush the other user's buffer
                    
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(UserIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "/DESsilenciar " & UserList(tUser).name)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And PlayerType.user Then Exit Sub
106         Call WriteShowSOSForm(UserIndex)

        End With

        
        Exit Sub

HandleSOSShowList_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleSOSShowList", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.user Then Call Ayuda.Quitar(UserName)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim x        As Byte

        Dim Y        As Byte
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then

            'Si es dios o Admins no podemos salvo que nosotros tambiín lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                
                    x = UserList(tUser).Pos.x
                    Y = UserList(tUser).Pos.Y + 1
                    Call FindLegalPos(UserIndex, UserList(tUser).Pos.Map, x, Y)
                
                    Call WarpUserChar(UserIndex, UserList(tUser).Pos.Map, x, Y, True)
                    
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                        

                    End If
                    
                    Call LogGM(.name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.x & " Y:" & UserList(tUser).Pos.Y)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleDesbuggear(ByVal UserIndex As Integer)

    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String, tUser As Integer, i As Long, Count As Long
        
        UserName = buffer.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            If Len(UserName) > 0 Then
                tUser = NameIndex(UserName)
                
                If tUser > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario debe estar offline.", FontTypeNames.FONTTYPE_INFO)
                Else

                    Dim AccountID As Long, AccountOnline As Boolean
                    
                    AccountID = GetAccountIDDatabase(UserName)
                    
                    If AccountID >= 0 Then

                        For i = 1 To LastUser

                            If UserList(i).flags.UserLogged Then
                                If UserList(i).AccountID = AccountID Then
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
                            Call ResetLoggedDatabase(AccountID)
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleDarLlaveAUsuario(ByVal UserIndex As Integer)

    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String, tUser As Integer, Llave As Integer
        
        UserName = buffer.ReadASCIIString()
        Llave = buffer.ReadInteger()
        
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleSacarLlave(ByVal UserIndex As Integer)

    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(UserIndex)

        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Llave As Integer
        
        Llave = .incomingData.ReadInteger()
        
        ' Solo dios o admin
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
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

End Sub

Private Sub HandleVerLlaves(ByVal UserIndex As Integer)

    With UserList(UserIndex)
    
        Call .incomingData.ReadByte

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

End Sub

Private Sub HandleUseKey(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(UserIndex)
    
        Call .incomingData.ReadByte
        
        Dim slot As Byte
        slot = .incomingData.ReadByte

        Call UsarLlave(UserIndex, slot)
                
    End With

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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And PlayerType.user Then Exit Sub
        
106         Call DoAdminInvisible(UserIndex)
108         Call LogGM(.name, "/INVISIBLE")

        End With

        
        Exit Sub

HandleInvisible_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleInvisible", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And PlayerType.user Then Exit Sub
        
106         Call WriteShowGMPanelForm(UserIndex)

        End With

        
        Exit Sub

HandleGMPanel_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleGMPanel", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.RoleMaster) Then Exit Sub
        
106         ReDim names(1 To LastUser) As String
108         Count = 1
        
110         For i = 1 To LastUser

112             If (LenB(UserList(i).name) <> 0) Then
                
114                 names(Count) = UserList(i).name
116                 Count = Count + 1
 
                End If

118         Next i
        
120         If Count > 1 Then Call WriteUserNameList(UserIndex, names(), Count - 1)

        End With

        
        Exit Sub

HandleRequestUserList_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestUserList", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.RoleMaster) Then Exit Sub
        
106         For i = 1 To LastUser

108             If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
110                 Users = Users & ", " & UserList(i).name
                
                    ' Display the user being checked by the centinel
112                 If modCentinela.Centinela.RevisandoUserIndex = i Then Users = Users & " (*)"

                End If

114         Next i
        
116         If LenB(Users) <> 0 Then
118             Users = Right$(Users, Len(Users) - 2)
120             Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & Users, FontTypeNames.FONTTYPE_INFO)
            Else
122             Call WriteConsoleMsg(UserIndex, "No hay usuarios trabajando", FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        
        Exit Sub

HandleWorking_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleWorking", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.RoleMaster) Then Exit Sub
        
106         For i = 1 To LastUser

108             If (LenB(UserList(i).name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
110                 Users = Users & UserList(i).name & ", "

                End If

112         Next i
        
114         If LenB(Users) <> 0 Then
116             Users = Left$(Users, Len(Users) - 2)
118             Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & Users, FontTypeNames.FONTTYPE_INFO)
            Else
120             Call WriteConsoleMsg(UserIndex, "No hay usuarios ocultandose", FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        
        Exit Sub

HandleHiding_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleHiding", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim Reason   As String

        Dim jailTime As Byte

        Dim Count    As Byte

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        jailTime = buffer.ReadByte()
        
        If InStr(1, UserName, "+") Then
            UserName = Replace(UserName, "+", " ")

        End If
        
        '/carcel nick@motivo@<tiempo>
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.user) <> 0 Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "El usuario no estí online.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If Not UserList(tUser).flags.Privilegios And PlayerType.user Then
                        Call WriteConsoleMsg(UserIndex, "No podés encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > 60 Then
                        Call WriteConsoleMsg(UserIndex, "No podés encarcelar por mís de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
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
                        Call LogGM(.name, " encarcelo a " & UserName)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And PlayerType.user Then Exit Sub
        
            Dim tNPC   As Integer

            Dim auxNPC As npc
        
106         tNPC = .flags.TargetNPC
        
108         If tNPC > 0 Then
110             Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & Npclist(tNPC).name, FontTypeNames.FONTTYPE_INFO)
            
112             auxNPC = Npclist(tNPC)
114             Call QuitarNPC(tNPC)
116             Call ReSpawnNpc(auxNPC)
            Else
118             Call WriteConsoleMsg(UserIndex, "Debes hacer click sobre el NPC antes", FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        
        Exit Sub

HandleKillNPC_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleKillNPC", Erl)
        Resume Next
        
End Sub

''
' Handles the "WarnUser" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/26/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim Reason   As String

        Dim privs    As PlayerType

        Dim Count    As Byte
        
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.user) <> 0 Then
            If LenB(UserName) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                privs = UserDarPrivilegioLevel(UserName)
                
                If Not privs And PlayerType.user Then
                    Call WriteConsoleMsg(UserIndex, "No podés advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If (InStrB(UserName, "\") <> 0) Then
                        UserName = Replace(UserName, "\", "")

                    End If

                    If (InStrB(UserName, "/") <> 0) Then
                        UserName = Replace(UserName, "/", "")

                    End If
                    
                    If PersonajeExiste(UserName) Then
                        If Database_Enabled Then
                            Call SavePenaDatabase(UserName, .name & " - ADVERTENCIA: " & Reason & " " & Date & " " & Time)
                        Else
                            Count = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", Count + 1)
                            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & Count + 1, LCase$(.name) & ": ADVERTENCIA por: " & LCase$(Reason) & " " & Date & " " & Time)

                        End If
                        
                        Call WriteConsoleMsg(UserIndex, "Has advertido a " & UCase$(UserName), FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.name, " advirtio a " & UserName)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleMensajeUser(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Ladder
    'Last Modification: 04/jul/2014
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String

        Dim Mensaje  As String

        Dim privs    As PlayerType

        Dim Count    As Byte

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        Mensaje = buffer.ReadASCIIString()
        
        tUser = NameIndex(UserName)
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.user) <> 0 Then
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
                    
                ' If tUser <= 0 Then
          
                ' If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                '   Call WriteVar(CharPath & UserName & ".chr", "INIT", "MENSAJEINFORMACION", .name & " te ha dejado un mensaje: " & Mensaje)
                '   Call WriteConsoleMsg(UserIndex, "El usuario estaba offline. El mensaje fue grabado en el charfile.", FontTypeNames.FONTTYPE_INFO)
                '   Call LogGM(.name, " envio el siguiente mensaje ha " & UCase$(UserName) & ": " & LCase$(Mensaje))
                '  Else
                '  Call WriteConsoleMsg(UserIndex, "El usuario no existe.", FontTypeNames.FONTTYPE_INFO)
                ' End If
                ' Else
                ' Call WriteConsoleMsg(tUser, .name & " te ha dejado un mensaje: " & Mensaje, FontTypeNames.FONTTYPE_CENTINELA)
                ' Call WriteConsoleMsg(UserIndex, "El mensaje fue enviado.", FontTypeNames.FONTTYPE_INFO)
                ' End If
            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

''
' Handles the "EditChar" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleTraerBoveda(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Ladder
    'Last Modification: 04/jul/2014
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        Call UpdateUserHechizos(True, UserIndex, 0)
       
        Call UpdateUserInv(True, UserIndex, 0)
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleEditChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/28/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 8 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName      As String

        Dim tUser         As Integer

        Dim opcion        As Byte

        Dim Arg1          As String

        Dim Arg2          As String

        Dim valido        As Boolean

        Dim LoopC         As Byte

        Dim commandString As String

        Dim n             As Byte
        
        UserName = Replace(buffer.ReadASCIIString(), "+", " ")
        
        If UCase$(UserName) = "YO" Then
            tUser = UserIndex
        Else
            tUser = NameIndex(UserName)

        End If
        
        opcion = buffer.ReadByte()
        Arg1 = buffer.ReadASCIIString()
        Arg2 = buffer.ReadASCIIString()
        
        If .flags.Privilegios And PlayerType.RoleMaster Then

            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)

                Case PlayerType.Consejero
                    ' Los RMs consejeros sílo se pueden editar su head, body y level
                    valido = tUser = UserIndex And (opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_Level)
                
                Case PlayerType.SemiDios
                    ' Los RMs sílo se pueden editar su level y el head y body de cualquiera
                    valido = (opcion = eEditOptions.eo_Level And tUser = UserIndex) Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head
                
                Case PlayerType.Dios
                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    ' pero si quiere modificar el level sílo lo puede hacer sobre sí mismo
                    valido = (opcion = eEditOptions.eo_Level And tUser = UserIndex) Or opcion = eEditOptions.eo_Body Or opcion = eEditOptions.eo_Head Or opcion = eEditOptions.eo_CiticensKilled Or opcion = eEditOptions.eo_CriminalsKilled Or opcion = eEditOptions.eo_Class Or opcion = eEditOptions.eo_Skills

            End Select
            
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then   'Si no es RM debe ser dios para poder usar este comando
            valido = True

        End If
        
        If valido Then

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
                        Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

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
                        Call ChangeUserChar(tUser, UserList(tUser).Char.Body, val(Arg1), UserList(tUser).Char.heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

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
                            UserList(tUser).Faccion.CiudadanosMatados = MAXUSERMATADOS
                        Else
                            UserList(tUser).Faccion.CiudadanosMatados = val(Arg1)

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

                            If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then Exit For
                        Next LoopC
                        
                        If LoopC > NUMCLASES Then
                            Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            UserList(tUser).clase = LoopC

                        End If

                    End If
                
                Case eEditOptions.eo_Skills

                    For LoopC = 1 To NUMSKILLS

                        If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg1) Then Exit For
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
                        Call WriteVar(CharPath & UserName & ".chr", "STATS", "SkillPtsLibres", Arg1)
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
                
                Case Else
                    Call WriteConsoleMsg(UserIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)

            End Select

        End If
        
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
                
            Case Else
                commandString = commandString & "UNKOWN "

        End Select
        
        commandString = commandString & Arg1 & " " & Arg2
        
        If valido Then Call LogGM(.name, commandString & " " & UserName)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
                
        Dim targetName  As String

        Dim targetIndex As Integer
        
        targetName = Replace$(buffer.ReadASCIIString(), "+", " ")
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer

        UserName = buffer.ReadASCIIString()
        
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim LoopC    As Long

        Dim Message  As String
        
        UserName = buffer.ReadASCIIString()
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
                    Message = Message & "CHAR>" & SkillsNames(LoopC) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & LoopC) & vbCrLf
                Next LoopC
                
                Call WriteConsoleMsg(UserIndex, Message & "CHAR> Libres:" & GetVar(CharPath & UserName & ".chr", "STATS", "SKILLPTSLIBRES"), FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendUserSkillsTxt(UserIndex, tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim LoopC    As Byte
        
        UserName = buffer.ReadASCIIString()
        
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
                UserList(tUser).Char.speeding = VelocidadNormal
                'Call WriteVelocidadToggle(tUser)
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageSpeedingACT(UserList(tUser).Char.CharIndex, UserList(tUser).Char.speeding))
                
                
                
                Call LogGM(.name, "Resucito a " & UserName)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    
100     With UserList(UserIndex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub

106         priv = PlayerType.Consejero Or PlayerType.SemiDios

108         If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
        
110         For i = 1 To LastUser

112             If UserList(i).flags.UserLogged Then
114                 If UserList(i).flags.Privilegios And priv Then list = list & UserList(i).name & ", "

                End If

116         Next i
        
118         If LenB(list) <> 0 Then
120             list = Left$(list, Len(list) - 2)
122             Call WriteConsoleMsg(UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
            Else
124             Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        
        Exit Sub

HandleOnlineGM_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleOnlineGM", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub
        
            Dim LoopC As Long

            Dim list  As String

            Dim priv  As PlayerType
        
106         priv = PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios

108         If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)
        
110         For LoopC = 1 To LastUser

112             If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).Pos.Map = .Pos.Map Then
114                 If UserList(LoopC).flags.Privilegios And priv Then list = list & UserList(LoopC).name & ", "

                End If

116         Next LoopC
        
118         If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        
120         Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)

        End With

        
        Exit Sub

HandleOnlineMap_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleOnlineMap", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Se asegura que el target es un npc
104         If .flags.TargetNPC = 0 Then
106             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar al sacerdote.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate NPC and make sure player is dead
108         If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 1 Then Exit Sub
        
            'Make sure it's close enough
110         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
                'Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
112             Call WriteConsoleMsg(UserIndex, "El sacerdote no puede escuchar tus pecados debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         If UserList(UserIndex).Faccion.Status = 1 Or UserList(UserIndex).Faccion.ArmadaReal = 1 Then
                'Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
116             Call WriteChatOverHead(UserIndex, "Tu alma ya esta libre de pecados hijo mio.", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If
        
118         If UserList(UserIndex).Faccion.CiudadanosMatados > 0 Or UserList(UserIndex).Faccion.ArmadaReal > 0 Then
120             Call WriteChatOverHead(UserIndex, "Has matado gente inocente, lamentablemente no podre concebirte el perdon.", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If
        
            Dim Clanalineacion As Byte
                        
122         If .GuildIndex <> 0 Then
124             Clanalineacion = modGuilds.Alineacion(.GuildIndex)

126             If Clanalineacion = 1 Then
128                 Call WriteChatOverHead(UserIndex, "Te encuentras en un clan criminal... debes retirarte para que pueda perdonarte.", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

            End If
        
130         Call WriteChatOverHead(UserIndex, "Con estas palabras, te libero de todo tipo de pecados. íQue dios te acompaíe hijo mio!", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbYellow)

132         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, "80", 100, False))
134         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("100", UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
136         UserList(UserIndex).Faccion.Status = 1
138         Call RefreshCharStatus(UserIndex)

        End With

        
        Exit Sub

HandleForgive_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleForgive", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim rank     As Integer
        
        rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        UserName = buffer.ReadASCIIString()
        
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then
                'If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                'Call WriteConsoleMsg(UserIndex, "Estís loco?? como vas a piíatear un gm!!!! :@", FontTypeNames.FONTTYPE_INFO)
                'Else
                Call UserDie(tUser)
                Call SendData(SendTarget.ToSuperiores, 0, PrepareMessageConsoleMsg(.name & " ha ejecutado a " & UserName, FontTypeNames.FONTTYPE_EJECUCION))
                Call LogGM(.name, " ejecuto a " & UserName)
                'End If
            Else
                Call WriteConsoleMsg(UserIndex, "No estí online", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim Reason   As String
        
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(UserIndex, UserName, Reason)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleSilenciarUser(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim Time     As Byte
        
        UserName = buffer.ReadASCIIString()
        Time = buffer.ReadByte()
    
        Call SilenciarUserName(UserIndex, UserName, Time)
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub
        
106         If .flags.TargetNPC > 0 Then
108             Call DoFollow(.flags.TargetNPC, .name)
110             Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
112             Npclist(.flags.TargetNPC).flags.Paralizado = 0
114             Npclist(.flags.TargetNPC).Contadores.Paralisis = 0

            End If

        End With

        
        Exit Sub

HandleNPCFollow_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleNPCFollow", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UserName <> "" Then
                tUser = NameIndex(UserName)
            Else
                tUser = .flags.TargetUser

            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El jugador no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else

                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.user)) <> 0 Then
                    Call WriteConsoleMsg(tUser, .name & " te hí trasportado.", FontTypeNames.FONTTYPE_INFO)
                    Call WarpToLegalPos(tUser, .Pos.Map, .Pos.x, .Pos.Y + 1, True)
                    
                    If UserList(tUser).flags.BattleModo = 1 Then
                        Call WriteConsoleMsg(UserIndex, "¡¡¡ATENCIíN!!! [" & UCase(UserList(tUser).name) & "] SE ENCUENTRA EN MODO BATTLE.", FontTypeNames.FONTTYPE_WARNING)
                        Call LogGM(.name, "ATENCIíN /SUM EN MODO BATTLE " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.x & " Y:" & .Pos.Y)
                    Else
                        Call LogGM(.name, "/SUM " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.x & " Y:" & .Pos.Y)

                    End If
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No podés invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub
        
106         Call EnviarSpawnList(UserIndex)

        End With

        
        Exit Sub

HandleSpawnListRequest_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleSpawnListRequest", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 3 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim npc As Integer

108         npc = .incomingData.ReadInteger()
        
110         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
112             If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then Call SpawnNpc(Declaraciones.SpawnList(npc).NpcIndex, .Pos, True, False)
            
114             Call LogGM(.name, "Sumoneo " & Declaraciones.SpawnList(npc).NpcName)

            End If

        End With

        
        Exit Sub

HandleSpawnCreature_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleSpawnCreature", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
106         If .flags.TargetNPC = 0 Then Exit Sub
        
108         Call ResetNpcInv(.flags.TargetNPC)
110         Call LogGM(.name, "/RESETINV " & Npclist(.flags.TargetNPC).name)

        End With

        
        Exit Sub

HandleResetNPCInventory_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleResetNPCInventory", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte

104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

            Call Limpieza.LimpiezaForzada
            
        End With

        Exit Sub

HandleCleanWorld_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCleanWorld", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(Message) <> 0 Then
                Call LogGM(.name, "Mensaje Broadcast:" & Message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & "> " & Message, FontTypeNames.FONTTYPE_SERVER))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim priv     As PlayerType
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 5 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim ip    As String

            Dim LoopC As Long

            Dim lista As String

            Dim priv  As PlayerType
        
108         ip = .incomingData.ReadByte() & "."
110         ip = ip & .incomingData.ReadByte() & "."
112         ip = ip & .incomingData.ReadByte() & "."
114         ip = ip & .incomingData.ReadByte()
        
116         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
118         Call LogGM(.name, "IP2NICK Solicito los Nicks de IP " & ip)
        
120         If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
122             priv = PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
124             priv = PlayerType.user

            End If

126         For LoopC = 1 To LastUser

128             If UserList(LoopC).ip = ip Then
130                 If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
132                     If UserList(LoopC).flags.Privilegios And priv Then
134                         lista = lista & UserList(LoopC).name & ", "

                        End If

                    End If

                End If

136         Next LoopC
        
138         If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
140         Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

        End With

        
        Exit Sub

HandleIPToNick_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleIPToNick", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName As String

        Dim tGuild    As Integer
        
        GuildName = buffer.ReadASCIIString()
        
        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")

        End If
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tGuild = GuildIndex(GuildName)
            
            If tGuild > 0 Then
                Call WriteConsoleMsg(UserIndex, "Clan " & UCase(GuildName) & ": " & modGuilds.m_ListaDeMiembrosOnline(UserIndex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 5 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Mapa As Integer

            Dim x    As Byte

            Dim Y    As Byte
        
108         Mapa = .incomingData.ReadInteger()
110         x = .incomingData.ReadByte()
112         Y = .incomingData.ReadByte()
        
114         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
116         Call LogGM(.name, "/CT " & Mapa & "," & x & "," & Y)
        
118         If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, x, Y) Then Exit Sub
        
120         If MapData(.Pos.Map, .Pos.x, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then Exit Sub
        
122         If MapData(.Pos.Map, .Pos.x, .Pos.Y - 1).TileExit.Map > 0 Then Exit Sub
        
124         If MapData(Mapa, x, Y).ObjInfo.ObjIndex > 0 Then
126             Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
128         If MapData(Mapa, x, Y).TileExit.Map > 0 Then
130             Call WriteConsoleMsg(UserIndex, "No podés crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            Rem Call WriteParticleFloorCreate(UserIndex, 37, -1, .Pos.map, .Pos.X, .Pos.Y - 1)
        
            Dim Objeto As obj
        
132         Objeto.Amount = 1
134         Objeto.ObjIndex = 378
136         Call MakeObj(Objeto, .Pos.Map, .Pos.x, .Pos.Y - 1)
        
138         With MapData(.Pos.Map, .Pos.x, .Pos.Y - 1)
140             .TileExit.Map = Mapa
142             .TileExit.x = x
144             .TileExit.Y = Y

            End With

        End With

        
        Exit Sub

HandleTeleportCreate_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleTeleportCreate", Erl)
        Resume Next
        
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
        
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            '/dt
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
106         Mapa = .flags.TargetMap
108         x = .flags.TargetX
110         Y = .flags.TargetY
        
112         If Not InMapBounds(Mapa, x, Y) Then Exit Sub
        
114         With MapData(Mapa, x, Y)

116             If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
118             If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.Map > 0 Then
120                 Call LogGM(UserList(UserIndex).name, "/DT: " & Mapa & "," & x & "," & Y)
                
122                 Call EraseObj(.ObjInfo.Amount, Mapa, x, Y)
                
124                 If MapData(.TileExit.Map, .TileExit.x, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
126                     Call EraseObj(1, .TileExit.Map, .TileExit.x, .TileExit.Y)

                    End If
                
128                 .TileExit.Map = 0
130                 .TileExit.x = 0
132                 .TileExit.Y = 0

                End If

            End With

        End With

        
        Exit Sub

HandleTeleportDestroy_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleTeleportDestroy", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub
        
106         Call LogGM(.name, "/LLUVIA")
108         Lloviendo = Not Lloviendo
110         Nebando = Not Nebando
        
112         Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
114         Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())

116         If Lloviendo Then
118             Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(404, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
120             Call SendData(SendTarget.ToAll, 0, PrepareMessageEfectToScreen(&HF5D3F3, 250)) 'Rayo
122             Call ApagarFogatas

            End If

        End With

        
        Exit Sub

HandleRainToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRainToggle", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim tUser As Integer

        Dim Desc  As String
        
        Desc = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.TargetUser

            If tUser > 0 Then
                UserList(tUser).DescRM = Desc
            Else
                Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes!", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 4 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim midiID As Byte

            Dim Mapa   As Integer
        
108         midiID = .incomingData.ReadByte
110         Mapa = .incomingData.ReadInteger
        
            'Solo dioses, admins y RMS
112         If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

                'Si el mapa no fue enviado tomo el actual
114             If Not InMapBounds(Mapa, 50, 50) Then
116                 Mapa = .Pos.Map

                End If
        
118             If midiID = 0 Then
                    'Ponemos el default del mapa
120                 Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(MapInfo(.Pos.Map).music_numberLow))
                Else
                    'Ponemos el pedido por el GM
122                 Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(midiID))

                End If

            End If

        End With

        
        Exit Sub

HanldeForceMIDIToMap_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HanldeForceMIDIToMap", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 6 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim waveID As Byte

            Dim Mapa   As Integer

            Dim x      As Byte

            Dim Y      As Byte
        
108         waveID = .incomingData.ReadByte()
110         Mapa = .incomingData.ReadInteger()
112         x = .incomingData.ReadByte()
114         Y = .incomingData.ReadByte()
        
            'Solo dioses, admins y RMS
116         If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

                'Si el mapa no fue enviado tomo el actual
118             If Not InMapBounds(Mapa, x, Y) Then
120                 Mapa = .Pos.Map
122                 x = .Pos.x
124                 Y = .Pos.Y

                End If
            
                'Ponemos el pedido por el GM
126             Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayWave(waveID, x, Y))

            End If

        End With

        
        Exit Sub

HandleForceWAVEToMap_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleForceWAVEToMap", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("ARMADA REAL> " & Message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & Message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & Message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & Message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(Message, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
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
        
122         Call LogGM(UserList(UserIndex).name, "/MASSDEST")

        End With

        
        Exit Sub

HandleDestroyAllItemsInArea_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleDestroyAllItemsInArea", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim LoopC    As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))

                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.x, .Pos.Y, False)

                End With

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim LoopC    As Byte
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Consejo de la Legiín Oscura.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.Map, .Pos.x, .Pos.Y, False)

                End With

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
            Dim tObj  As Integer

            Dim lista As String

            Dim x     As Long

            Dim Y     As Long
        
106         For x = 5 To 95
108             For Y = 5 To 95
110                 tObj = MapData(.Pos.Map, x, Y).ObjInfo.ObjIndex

112                 If tObj > 0 Then
114                     If ObjData(tObj).OBJType <> eOBJType.otArboles Then
116                         Call WriteConsoleMsg(UserIndex, "(" & x & "," & Y & ") " & ObjData(tObj).name, FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

118             Next Y
120         Next x

        End With

        
        Exit Sub

HandleItemsInTheFloor_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleItemsInTheFloor", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)

            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(UserName)

            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
                

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     With UserList(UserIndex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
106         Call SecurityIp.DumpTables

        End With

        
        Exit Sub

HandleDumpIPTables_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleDumpIPTables", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
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
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.x, .Pos.Y)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill", FontTypeNames.FONTTYPE_CONSEJO))

                    End If
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de la Legiín Oscura", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.x, .Pos.Y)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de la Legiín Oscura", FontTypeNames.FONTTYPE_CONSEJO))

                    End If

                End With

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim tTrigger As Byte

            Dim tLog     As String
        
108         tTrigger = .incomingData.ReadByte()
        
110         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
112         If tTrigger >= 0 Then
114             MapData(.Pos.Map, .Pos.x, .Pos.Y).trigger = tTrigger
116             tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.x & "," & .Pos.Y
            
118             Call LogGM(.name, tLog)
120             Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        
        Exit Sub

HandleSetTrigger_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleSetTrigger", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         tTrigger = MapData(.Pos.Map, .Pos.x, .Pos.Y).trigger
        
108         Call LogGM(.name, "Miro el trigger en " & .Pos.Map & "," & .Pos.x & "," & .Pos.Y & ". Era " & tTrigger)
        
110         Call WriteConsoleMsg(UserIndex, "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.x & ", " & .Pos.Y, FontTypeNames.FONTTYPE_INFO)

        End With

        
        Exit Sub

HandleAskTrigger_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleAskTrigger", Erl)
        Resume Next
        
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
100     With UserList(UserIndex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
            Dim lista As String

            Dim LoopC As Long
        
106         Call LogGM(.name, "/BANIPLIST")
        
108         For LoopC = 1 To BanIps.Count
110             lista = lista & BanIps.Item(LoopC) & ", "
112         Next LoopC
        
114         If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
116         Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)

        End With

        
        Exit Sub

HandleBannedIPList_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleBannedIPList", Erl)
        Resume Next
        
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
100     With UserList(UserIndex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         Call BanIpGuardar
108         Call BanIpCargar

        End With

        
        Exit Sub

HandleBannedIPReload_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleBannedIPReload", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim GuildName   As String

        Dim cantMembers As Integer

        Dim LoopC       As Long

        Dim member      As String

        Dim Count       As Byte

        Dim tIndex      As Integer

        Dim tFile       As String
        
        GuildName = buffer.ReadASCIIString()
        
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
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    'member es la victima
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim bannedIP As String

        Dim tUser    As Integer

        Dim Reason   As String

        Dim i        As Long
        
        ' Is it by ip??
        If buffer.ReadBoolean() Then
            bannedIP = buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte() & "."
            bannedIP = bannedIP & buffer.ReadByte()
        Else
            tUser = NameIndex(buffer.ReadASCIIString())
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no estí online.", FontTypeNames.FONTTYPE_INFO)
            Else
                bannedIP = UserList(tUser).ip

            End If

        End If
        
        Reason = buffer.ReadASCIIString()
        
        If LenB(bannedIP) > 0 Then
            If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                Call LogGM(.name, "/BanIP " & bannedIP & " por " & Reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                Call BanIpAgrega(bannedIP)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " baneí la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))
                
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 5 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim bannedIP As String
        
108         bannedIP = .incomingData.ReadByte() & "."
110         bannedIP = bannedIP & .incomingData.ReadByte() & "."
112         bannedIP = bannedIP & .incomingData.ReadByte() & "."
114         bannedIP = bannedIP & .incomingData.ReadByte()
        
116         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
118         If BanIpQuita(bannedIP) Then
120             Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
            Else
122             Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        
        Exit Sub

HandleUnbanIP_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleUnbanIP", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 5 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte

            Dim tObj    As Integer

            Dim Cuantos As Integer
        
108         tObj = .incomingData.ReadInteger()
110         Cuantos = .incomingData.ReadInteger()
        
112         Call LogGM(.name, "/CI: " & tObj & " Cantidad : " & Cuantos)
        
114         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
116         If ObjData(tObj).donador = 1 Then
118             If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios) Then Exit Sub

            End If
         
            'If MapData(.Pos.Map, .Pos.X, .Pos.y - 1).ObjInfo.ObjIndex > 0 Then _
             Exit Sub

120         If Cuantos > 10000 Then Call WriteConsoleMsg(UserIndex, "Demasiados, míximo para crear : 10.000", FontTypeNames.FONTTYPE_TALK): Exit Sub

122         If MapData(.Pos.Map, .Pos.x, .Pos.Y - 1).TileExit.Map > 0 Then Exit Sub
        
124         If tObj < 1 Or tObj > NumObjDatas Then Exit Sub
        
            'Is the object not null?
126         If LenB(ObjData(tObj).name) = 0 Then Exit Sub
        
            Dim Objeto As obj
        
128         Objeto.Amount = Cuantos
130         Objeto.ObjIndex = tObj
132         Call MakeObj(Objeto, .Pos.Map, .Pos.x, .Pos.Y)

        End With

        
        Exit Sub

HandleCreateItem_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCreateItem", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
106         If MapData(.Pos.Map, .Pos.x, .Pos.Y).ObjInfo.ObjIndex = 0 Then Exit Sub
        
108         Call LogGM(.name, "/DEST")
        
            ' If ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then
            ''  Call WriteConsoleMsg(UserIndex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            '  Exit Sub
            ' End If
        
110         Call EraseObj(10000, .Pos.Map, .Pos.x, .Pos.Y)

        End With

        
        Exit Sub

HandleDestroyItems_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleDestroyItems", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte

            Dim midiID As Byte

108         midiID = .incomingData.ReadByte()
        
110         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
112         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " broadcast musica: " & midiID, FontTypeNames.FONTTYPE_SERVER))
        
114         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))

        End With

        
        Exit Sub

HandleForceMIDIAll_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleForceMIDIAll", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte

            Dim waveID As Byte

108         waveID = .incomingData.ReadByte()
        
110         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
112         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))

        End With

        
        Exit Sub

HandleForceWAVEAll_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleForceWAVEAll", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName   As String

        Dim punishment As Byte

        Dim NewText    As String
        
        UserName = buffer.ReadASCIIString()
        punishment = buffer.ReadByte
        NewText = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     With UserList(UserIndex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub

106         Call LogGM(.name, "/BLOQ")
        
108         If MapData(.Pos.Map, .Pos.x, .Pos.Y).Blocked = 0 Then
110             MapData(.Pos.Map, .Pos.x, .Pos.Y).Blocked = 2
            Else
112             MapData(.Pos.Map, .Pos.x, .Pos.Y).Blocked = 0

            End If
        
114         Call Bloquear(True, .Pos.Map, .Pos.x, .Pos.Y, MapData(.Pos.Map, .Pos.x, .Pos.Y).Blocked)

        End With

        
        Exit Sub

HandleTileBlockedToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleTileBlockedToggle", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub
        
106         If .flags.TargetNPC = 0 Then Exit Sub
        
108         Call QuitarNPC(.flags.TargetNPC)
110         Call LogGM(.name, "/MATA " & Npclist(.flags.TargetNPC).name)

        End With

        
        Exit Sub

HandleKillNPCNoRespawn_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleKillNPCNoRespawn", Erl)
        Resume Next
        
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
100     With UserList(UserIndex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
            Dim x As Long

            Dim Y As Long
        
106         For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
108             For x = .Pos.x - MinXBorder + 1 To .Pos.x + MinXBorder - 1

110                 If x > 0 And Y > 0 And x < 101 And Y < 101 Then
112                     If MapData(.Pos.Map, x, Y).NpcIndex > 0 Then Call QuitarNPC(MapData(.Pos.Map, x, Y).NpcIndex)

                    End If

114             Next x
116         Next Y

118         Call LogGM(.name, "/MASSKILL")

        End With

        
        Exit Sub

HandleKillAllNearbyNPCs_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleKillAllNearbyNPCs", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName   As String

        Dim lista      As String

        Dim LoopC      As Byte

        Dim priv       As Integer

        Dim validCheck As Boolean
        
        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then

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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 4 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Color As Long
        
108         Color = RGB(.incomingData.ReadByte(), .incomingData.ReadByte(), .incomingData.ReadByte())
        
110         If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
112             .flags.ChatColor = Color

            End If

        End With

        
        Exit Sub

HandleChatColor_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleChatColor", Erl)
        Resume Next
        
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
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
106             .flags.AdminPerseguible = Not .flags.AdminPerseguible

            End If

        End With

        
        Exit Sub

HandleIgnored_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleIgnored", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        'Reads the UserName and Slot Packets
        Dim UserName As String

        Dim slot     As Byte

        Dim tIndex   As Integer
        
        UserName = buffer.ReadASCIIString() 'Que UserName?
        slot = buffer.ReadByte() 'Que Slot?
        tIndex = NameIndex(UserName)  'Que user index?
        
        Call LogGM(.name, .name & " Checkeo el slot " & slot & " de " & UserName)
           
        If tIndex > 0 And UserList(UserIndex).flags.BattleModo = 0 Then
            If slot > 0 And slot <= UserList(UserIndex).CurrentInventorySlots Then
                If UserList(tIndex).Invent.Object(slot).ObjIndex > 0 Then
                    Call WriteConsoleMsg(UserIndex, " Objeto " & slot & ") " & ObjData(UserList(tIndex).Invent.Object(slot).ObjIndex).name & " Cantidad:" & UserList(tIndex).Invent.Object(slot).Amount, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay Objeto en slot seleccionado", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "Slot Invílido.", FontTypeNames.FONTTYPE_TALK)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     With UserList(UserIndex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
106         Call WriteConsoleMsg(UserIndex, "TID: " & CStr(ReiniciarAutoUpdate()), FontTypeNames.FONTTYPE_INFO)

        End With

        
        Exit Sub

HandleResetAutoUpdate_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleResetAutoUpdate", Erl)
        Resume Next
        
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
            'Remove Packet ID
102         Call .incomingData.ReadByte
    
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
            'time and Time BUG!
106         Call LogGM(.name, .name & " reinicio el mundo")
        
108         Call ReiniciarServidor(True)

        End With

        
        Exit Sub

HandleRestart_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRestart", Erl)
        Resume Next
        
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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         Call LogGM(.name, .name & " ha recargado a los objetos.")
        
108         Call LoadOBJData
110         Call LoadPesca
112         Call LoadRecursosEspeciales
114         Call WriteConsoleMsg(UserIndex, "Obj.dat recargado exitosamente.", FontTypeNames.FONTTYPE_SERVER)

        End With

        
        Exit Sub

HandleReloadObjects_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleReloadObjects", Erl)
        Resume Next
        
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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         Call LogGM(.name, .name & " ha recargado los hechizos.")
        
108         Call CargarHechizos

        End With

        
        Exit Sub

HandleReloadSpells_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleReloadSpells", Erl)
        Resume Next
        
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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         Call LogGM(.name, .name & " ha recargado los INITs.")
        
108         Call LoadSini

        End With

        
        Exit Sub

HandleReloadServerIni_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleReloadServerIni", Erl)
        Resume Next
        
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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
         
106         Call LogGM(.name, .name & " ha recargado los NPCs.")
    
108         Call CargaNpcsDat
    
110         Call WriteConsoleMsg(UserIndex, "Npcs.dat recargado exitosamente.", FontTypeNames.FONTTYPE_SERVER)

        End With

        
        Exit Sub

HandleReloadNPCs_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleReloadNPCs", Erl)
        Resume Next
        
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
100     With UserList(UserIndex)
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
                
            Dim list  As String

            Dim Count As Long

            Dim i     As Long
        
106         Call LogGM(.name, .name & " ha pedido las estadisticas del TCP.")
    
108         Call WriteConsoleMsg(UserIndex, "Los datos estín en BYTES.", FontTypeNames.FONTTYPE_INFO)
        
            'Send the stats
110         With TCPESStats
112             Call WriteConsoleMsg(UserIndex, "IN/s: " & .BytesRecibidosXSEG & " OUT/s: " & .BytesEnviadosXSEG, FontTypeNames.FONTTYPE_INFO)
114             Call WriteConsoleMsg(UserIndex, "IN/s MAX: " & .BytesRecibidosXSEGMax & " -> " & .BytesRecibidosXSEGCuando, FontTypeNames.FONTTYPE_INFO)
116             Call WriteConsoleMsg(UserIndex, "OUT/s MAX: " & .BytesEnviadosXSEGMax & " -> " & .BytesEnviadosXSEGCuando, FontTypeNames.FONTTYPE_INFO)

            End With
        
            'Search for users that are working
118         For i = 1 To LastUser

120             With UserList(i)

122                 If .flags.UserLogged And .ConnID >= 0 And .ConnIDValida Then
124                     If .outgoingData.length > 0 Then
126                         list = list & .name & " (" & CStr(.outgoingData.length) & "), "
128                         Count = Count + 1

                        End If

                    End If

                End With

130         Next i
        
132         Call WriteConsoleMsg(UserIndex, "Posibles pjs trabados: " & CStr(Count), FontTypeNames.FONTTYPE_INFO)
134         Call WriteConsoleMsg(UserIndex, list, FontTypeNames.FONTTYPE_INFO)

        End With

        
        Exit Sub

HandleRequestTCPStats_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestTCPStats", Erl)
        Resume Next
        
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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         Call LogGM(.name, .name & " ha echado a todos los personajes.")
        
108         Call EcharPjsNoPrivilegiados

        End With

        
        Exit Sub

HandleKickAllChars_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleKickAllChars", Erl)
        Resume Next
        
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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
106         HoraFanstasia = 1259
    
        End With

        
        Exit Sub

HandleNight_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleNight", Erl)
        Resume Next
        
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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         Call LogGM(.name, .name & " ha solicitado mostrar el formulario del servidor.")
108         Call frmMain.mnuMostrar_Click

        End With

        
        Exit Sub

HandleShowServerForm_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleShowServerForm", Erl)
        Resume Next
        
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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         Call LogGM(.name, .name & " ha borrado los SOS")
        
108         Call Ayuda.Reset

        End With

        
        Exit Sub

HandleCleanSOS_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCleanSOS", Erl)
        Resume Next
        
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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
106         Call LogGM(.name, .name & " ha guardado todos los chars")
        
108         Call GuardarUsuarios

        End With

        
        Exit Sub

HandleSaveChars_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleSaveChars", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove Packet ID
106         Call .incomingData.ReadByte
        
            Dim doTheBackUp As Boolean
        
108         doTheBackUp = .incomingData.ReadBoolean()
        
110         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
112         Call LogGM(.name, .name & " ha cambiado la informaciín sobre el BackUp")
        
            'Change the boolean to byte in a fast way
114         If doTheBackUp Then
116             MapInfo(.Pos.Map).backup_mode = 1
            Else
118             MapInfo(.Pos.Map).backup_mode = 0

            End If
        
            'Change the boolean to string in a fast way
120         Call WriteVar(MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapInfo(.Pos.Map).backup_mode)
        
122         Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).backup_mode, FontTypeNames.FONTTYPE_INFO)

        End With

        
        Exit Sub

HandleChangeMapInfoBackup_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleChangeMapInfoBackup", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove Packet ID
106         Call .incomingData.ReadByte
        
            Dim isMapPk As Boolean
        
108         isMapPk = .incomingData.ReadBoolean()
        
110         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
112         Call LogGM(.name, .name & " ha cambiado la informacion sobre si es seguro el mapa.")
        
114         MapInfo(.Pos.Map).Seguro = isMapPk
        
            'Change the boolean to string in a fast way
            Rem Call WriteVar(App.Path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "Pk", IIf(isMapPk, "1", "0"))

116         Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " PK: " & MapInfo(.Pos.Map).Seguro, FontTypeNames.FONTTYPE_INFO)

        End With

        
        Exit Sub

HandleChangeMapInfoPK_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleChangeMapInfoPK", Erl)
        Resume Next
        
End Sub

''
' Handle the "ChangeMapInfoRestricted" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
    '***************************************************
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    Dim tStr As String
    
    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.name, .name & " ha cambiado la informacion sobre si es Restringido el mapa.")
                MapInfo(UserList(UserIndex).Pos.Map).restrict_mode = tStr
                Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Restringir", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Restringido: " & MapInfo(.Pos.Map).restrict_mode, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
        Dim nomagic As Boolean
    
104     With UserList(UserIndex)
            'Remove Packet ID
106         Call .incomingData.ReadByte
        
108         nomagic = .incomingData.ReadBoolean
        
110         If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
112             Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar la Magia el mapa.")

                ' MapInfo(UserList(UserIndex).Pos.map).MagiaSinEfecto = nomagic
                'Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "MagiaSinEfecto", nomagic)
                '  Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " MagiaSinEfecto: " & MapInfo(.Pos.map).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        
        Exit Sub

HandleChangeMapInfoNoMagic_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleChangeMapInfoNoMagic", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
        Dim noinvi As Boolean
    
104     With UserList(UserIndex)
            'Remove Packet ID
106         Call .incomingData.ReadByte
        
108         noinvi = .incomingData.ReadBoolean()
        
110         If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
112             Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Invisibilidad el mapa.")

                ' MapInfo(UserList(UserIndex).Pos.map).InviSinEfecto = noinvi
                'Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "InviSinEfecto", noinvi)
                ' Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " InviSinEfecto: " & MapInfo(.Pos.map).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        
        Exit Sub

HandleChangeMapInfoNoInvi_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleChangeMapInfoNoInvi", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 2 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
        Dim noresu As Boolean
    
104     With UserList(UserIndex)
            'Remove Packet ID
106         Call .incomingData.ReadByte
        
108         noresu = .incomingData.ReadBoolean()
        
110         If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
112             Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Resucitar el mapa.")

                '  MapInfo(UserList(UserIndex).Pos.map).ResuSinEfecto = noresu
                ' Call WriteVar(App.Path & MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "ResuSinEfecto", noresu)
                ' Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " ResuSinEfecto: " & MapInfo(.Pos.map).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        
        Exit Sub

HandleChangeMapInfoNoResu_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleChangeMapInfoNoResu", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    Dim tStr As String
    
    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    Dim tStr As String
    
    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         Call LogGM(.name, .name & " ha guardado el mapa " & CStr(.Pos.Map))
        
            ' Call GrabarMapa(.Pos.Map, App.Path & "\WorldBackUp\Mapa" & .Pos.Map)
        
108         Call WriteConsoleMsg(UserIndex, "Mapa Guardado", FontTypeNames.FONTTYPE_INFO)

        End With

        
        Exit Sub

HandleSaveMap_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleSaveMap", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        
        guild = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(UserIndex, guild)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub

106         Call LogGM(.name, .name & " ha hecho un backup")
        
108         Call ES.DoBackUp 'Sino lo confunde con la id del paquete

        End With

        
        Exit Sub

HandleDoBackUp_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleDoBackUp", Erl)
        Resume Next
        
End Sub

''
' Handle the "ToggleCentinelActivated" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleToggleCentinelActivated(ByVal UserIndex As Integer)
        
        On Error GoTo HandleToggleCentinelActivated_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/26/06
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
        'Activate or desactivate the Centinel
        '***************************************************
100     With UserList(UserIndex)
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
106         centinelaActivado = Not centinelaActivado
        
108         With Centinela
110             .RevisandoUserIndex = 0
112             .clave = 0
114             .TiempoRestante = 0

            End With
    
116         If CentinelaNPCIndex Then
118             Call QuitarNPC(CentinelaNPCIndex)
120             CentinelaNPCIndex = 0

            End If
        
122         If centinelaActivado Then
124             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido activado.", FontTypeNames.FONTTYPE_SERVER))
            Else
126             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El centinela ha sido desactivado.", FontTypeNames.FONTTYPE_SERVER))

            End If

        End With

        
        Exit Sub

HandleToggleCentinelActivated_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleToggleCentinelActivated", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        'Reads the userName and newUser Packets
        Dim UserName     As String

        Dim newName      As String

        Dim changeNameUI As Integer

        Dim GuildIndex   As Integer
        
        UserName = buffer.ReadASCIIString()
        newName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim newMail  As String
        
        UserName = buffer.ReadASCIIString()
        newMail = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim copyFrom As String

        Dim Password As String
        
        UserName = Replace(buffer.ReadASCIIString(), "+", " ")
        copyFrom = Replace(buffer.ReadASCIIString(), "+", " ")
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
100     If UserList(UserIndex).incomingData.length < 3 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove Packet ID
106         Call .incomingData.ReadByte
        
            Dim NpcIndex As Integer
        
108         NpcIndex = .incomingData.ReadInteger()
        
110         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
112         NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        
114         If NpcIndex <> 0 Then
116             Call LogGM(.name, "Sumoneo a " & Npclist(NpcIndex).name & " en mapa " & .Pos.Map)

            End If

        End With

        
        Exit Sub

HandleCreateNPC_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCreateNPC", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 3 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove Packet ID
106         Call .incomingData.ReadByte
        
            Dim NpcIndex As Integer
        
108         NpcIndex = .incomingData.ReadInteger()
        
110         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
112         NpcIndex = SpawnNpc(NpcIndex, .Pos, True, True)
        
114         If NpcIndex <> 0 Then
116             Call LogGM(.name, "Sumoneo con respawn " & Npclist(NpcIndex).name & " en mapa " & .Pos.Map)

            End If

        End With

        
        Exit Sub

HandleCreateNPCWithRespawn_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCreateNPCWithRespawn", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 4 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove Packet ID
106         Call .incomingData.ReadByte
        
            Dim Index    As Byte

            Dim ObjIndex As Integer
        
108         Index = .incomingData.ReadByte()
110         ObjIndex = .incomingData.ReadInteger()
        
112         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
114         Select Case Index

                Case 1
                    ' ArmaduraImperial1 = objindex
            
116             Case 2
                    ' ArmaduraImperial2 = objindex
            
118             Case 3
                    ' ArmaduraImperial3 = objindex
            
120             Case 4

                    ' TunicaMagoImperial = objindex
            End Select

        End With

        
        Exit Sub

HandleImperialArmour_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleImperialArmour", Erl)
        Resume Next
        
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
100     If UserList(UserIndex).incomingData.length < 4 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove Packet ID
106         Call .incomingData.ReadByte
        
            Dim Index    As Byte

            Dim ObjIndex As Integer
        
108         Index = .incomingData.ReadByte()
110         ObjIndex = .incomingData.ReadInteger()
        
112         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
114         Select Case Index

                Case 1
                    '   ArmaduraCaos1 = objindex
            
116             Case 2
                    '   ArmaduraCaos2 = objindex
            
118             Case 3
                    '   ArmaduraCaos3 = objindex
            
120             Case 4

                    '  TunicaMagoCaos = objindex
            End Select

        End With

        
        Exit Sub

HandleChaosArmour_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleChaosArmour", Erl)
        Resume Next
        
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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub
        
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
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleNavigateToggle", Erl)
        Resume Next
        
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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         If ServerSoloGMs > 0 Then
108             Call WriteConsoleMsg(UserIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
110             ServerSoloGMs = 0
            Else
112             Call WriteConsoleMsg(UserIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
114             ServerSoloGMs = 1

            End If

        End With

        
        Exit Sub

HandleServerOpenToUsersToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleServerOpenToUsersToggle", Erl)
        Resume Next
        
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
    
100     With UserList(UserIndex)
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If Torneo.HayTorneoaActivo = False Then
106             Call WriteConsoleMsg(UserIndex, "No hay ningún evento disponible.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
                
108         If .flags.BattleModo = 1 Then
110             Call WriteConsoleMsg(UserIndex, "No podes participar desde aquí.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
    
112         If .flags.EnTorneo Then
114             Call WriteConsoleMsg(UserIndex, "Ya estás participando.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
    
116         If .Stats.ELV > Torneo.nivelmaximo Then
118             Call WriteConsoleMsg(UserIndex, "El nivel míximo para participar es " & Torneo.nivelmaximo & ".", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
    
120         If .Stats.ELV < Torneo.NivelMinimo Then
122             Call WriteConsoleMsg(UserIndex, "El nivel mínimo para participar es " & Torneo.NivelMinimo & ".", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
    
124         If .Stats.GLD < Torneo.costo Then
126             Call WriteConsoleMsg(UserIndex, "No tienes suficiente oro para ingresar.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
128         If .clase = Mage And Torneo.mago = 0 Then
130             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
132         If .clase = Cleric And Torneo.clerico = 0 Then
134             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
136         If .clase = Warrior And Torneo.guerrero = 0 Then
138             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
140         If .clase = Bard And Torneo.bardo = 0 Then
142             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
144         If .clase = Assasin And Torneo.asesino = 0 Then
146             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
   
148         If .clase = Druid And Torneo.druido = 0 Then
150             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
152         If .clase = Paladin And Torneo.Paladin = 0 Then
154             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
156         If .clase = Hunter And Torneo.cazador = 0 Then
158             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
160         If .clase = Trabajador And Torneo.cazador = 0 Then
162             Call WriteConsoleMsg(UserIndex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
   
164         If Torneo.Participantes = Torneo.cupos Then
166             Call WriteConsoleMsg(UserIndex, "Los cupos ya estan llenos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
  
168         Call ParticiparTorneo(UserIndex)

        End With

        
        Exit Sub

HandleParticipar_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleParticipar", Erl)
        Resume Next
        
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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "/CONDEN " & UserName)
            
            tUser = NameIndex(UserName)

            If tUser > 0 Then Call VolverCriminal(tUser)

        End If
                
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "/RAJAR " & UserName)
            
            tUser = NameIndex(UserName)
            
            If tUser > 0 Then Call ResetFacciones(tUser)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName   As String

        Dim GuildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "/RAJARCLAN " & UserName)
            
            GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
            
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "No pertenece a ningún clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim mail     As String
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If FileExist(CharPath & UserName & ".chr") Then
                mail = GetVar(CharPath & UserName & ".chr", "CONTACTO", "email")
                
                Call WriteConsoleMsg(UserIndex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim Message As String

        Message = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "Mensaje de sistema:" & Message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(Message))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim newMOTD           As String

        Dim auxiliaryString() As String

        Dim LoopC             As Long
        
        newMOTD = buffer.ReadASCIIString()
        auxiliaryString = Split(newMOTD, vbCrLf)
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

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
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) Then
                Exit Sub

            End If
        
            Dim auxiliaryString As String

            Dim LoopC           As Long
        
106         For LoopC = LBound(MOTD()) To UBound(MOTD())
108             auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
110         Next LoopC
        
112         If Len(auxiliaryString) >= 2 Then
114             If Right$(auxiliaryString, 2) = vbCrLf Then
116                 auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)

                End If

            End If
        
118         Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)

        End With

        
        Exit Sub

HandleChangeMOTD_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleChangeMOTD", Erl)
        Resume Next
        
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
100     With UserList(UserIndex)
            'Remove Packet ID
102         Call .incomingData.ReadByte

            Dim Time As Long
        
104         Time = .incomingData.ReadLong()
        
106         Call WritePong(UserIndex, Time)

        End With

        
        Exit Sub

HandlePing_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandlePing", Erl)
        Resume Next
        
End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)

    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.logged)
    Exit Sub
Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteHora(ByVal UserIndex As Integer)

    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageHora())
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteNadarToggle(ByVal UserIndex As Integer, ByVal Puede As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.NadarToggle)
        Call .WriteBoolean(Puede)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If
    
End Sub

Public Sub WriteEquiteToggle(ByVal UserIndex As Integer)
        
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.EquiteToggle)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteVelocidadToggle(ByVal UserIndex As Integer)
        
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.VelocidadToggle)
        Call .WriteSingle(UserList(UserIndex).Char.speeding)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteMacroTrabajoToggle(ByVal UserIndex As Integer, ByVal Activar As Boolean)

    If Not Activar Then
        UserList(UserIndex).flags.TargetObj = 0 ' Sacamos el targer del objeto
        UserList(UserIndex).flags.UltimoMensaje = 0
        UserList(UserIndex).Counters.Trabajando = 0
        UserList(UserIndex).flags.UsandoMacro = False
       
    Else
        UserList(UserIndex).flags.UsandoMacro = True

    End If

    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MacroTrabajoToggle)
        Call .WriteBoolean(Activar)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Disconnect" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    'If UserList(UserIndex).flags.BattleModo = 0 Then
    '    Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
    'Else
    '    Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "Battle", "Puntos", UserList(UserIndex).flags.BattlePuntos)
    'End If
    
    'Call WriteVar(CuentasPath & UCase$(UserList(UserIndex).cuenta) & ".act", "INIT", "LOGEADA", 0)
    Call WritePersonajesDeCuenta(UserIndex)

    Call WriteMostrarCuenta(UserIndex)
    
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Disconnect)

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankEnd)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
    Call UserList(UserIndex).outgoingData.WriteASCIIString(Npclist(UserList(UserIndex).flags.TargetNPC).name)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankInit)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowBlacksmithForm)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowCarpenterForm)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteShowAlquimiaForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowAlquimiaForm)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteShowSastreForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowSastreForm)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCKillUser(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NPCKillUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NPCKillUser)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldUser(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldUser)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldOther(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldOther)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CharSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharSwing(ByVal UserIndex As Integer, ByVal CharIndex As Integer, Optional ByVal FX As Boolean = True, Optional ByVal ShowText As Boolean = True)

    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharSwing(CharIndex, FX, ShowText))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Function PrepareMessageCharSwing(ByVal CharIndex As Integer, Optional ByVal FX As Boolean = True, Optional ByVal ShowText As Boolean = True) As String
        
        On Error GoTo PrepareMessageCharSwing_Err
        

        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.CharSwing)
104         Call .WriteInteger(CharIndex)
106         Call .WriteBoolean(FX)
108         Call .WriteBoolean(ShowText)
        
110         PrepareMessageCharSwing = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageCharSwing_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCharSwing", Erl)
        Resume Next
        
End Function

''
' Writes the "SafeModeOn" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOn(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeModeOn" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.SafeModeOn)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "SafeModeOff" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOff(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeModeOff" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.SafeModeOff)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "PartySafeOn" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySafeOn(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Rapsodius
    'Last Modification: 10/10/07
    'Writes the "PartySafeOn" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.PartySafeOn)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "PartySafeOff" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySafeOff(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Rapsodius
    'Last Modification: 10/10/07
    'Writes the "PartySafeOff" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.PartySafeOff)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteClanSeguro(ByVal UserIndex As Integer, ByVal estado As Boolean)

    '***************************************************
    'Author: Rapsodius
    'Last Modification: 10/10/07
    'Writes the "PartySafeOff" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ClanSeguro)
    Call UserList(UserIndex).outgoingData.WriteBoolean(estado)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCantUseWhileMeditating(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CantUseWhileMeditating)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)

    'Call SendData(SendTarget.ToDiosesYclan, UserIndex, PrepareMessageCharUpdateHP(UserIndex))
    Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, PrepareMessageCharUpdateHP(UserIndex))

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateGold" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateExp" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, ByVal Map As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMap" message to the given user's outgoing data buffer
    '***************************************************

    Dim Version As Integer

    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(Map)
        Call .WriteInteger(Version)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PosUpdate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(UserIndex).Pos.x)
        Call .WriteByte(UserList(UserIndex).Pos.Y)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "NPCHitUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCHitUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal damage As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NPCHitUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.NPCHitUser)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserHitNPC" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    damage The number of HP lost by the target creature.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHitNPC(ByVal UserIndex As Integer, ByVal damage As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserHitNPC" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserHitNPC)
        
        'It is a long to allow the "drake slayer" (matadracos) to kill the great red dragon of one blow.
        Call .WriteLong(damage)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserAttackedSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex       User to which the message is intended.
' @param    attackerIndex   The user index of the user that attacked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserAttackedSwing(ByVal UserIndex As Integer, ByVal attackerIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserAttackedSwing" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserAttackedSwing)
        Call .WriteInteger(UserList(attackerIndex).Char.CharIndex)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserHittedByUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackerChar Char index of the user hitted.
' @param    damage The number of HP lost by the hit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHittedByUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal attackerChar As Integer, ByVal damage As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserHittedByUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserHittedByUser)
        Call .WriteInteger(attackerChar)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserHittedUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    target Part of the body where the user was hitted.
' @param    attackedChar Char index of the user hitted.
' @param    damage The number of HP lost by the oponent hitted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHittedUser(ByVal UserIndex As Integer, ByVal Target As PartesCuerpo, ByVal attackedChar As Integer, ByVal damage As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserHittedUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserHittedUser)
        Call .WriteInteger(attackedChar)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, ByVal chat As String, ByVal CharIndex As Integer, ByVal Color As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatOverHead" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(chat, CharIndex, Color))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteEfectOverHead(ByVal UserIndex As Integer, ByVal chat As String, ByVal CharIndex As Integer, Optional ByVal Color As Long = &HFF0000)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatOverHead" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageEfectOverHead(chat, CharIndex, Color))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteExpOverHead(ByVal UserIndex As Integer, ByVal chat As String, ByVal CharIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatOverHead" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageExpOverHead(chat, CharIndex))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteOroOverHead(ByVal UserIndex As Integer, ByVal chat As String, ByVal CharIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatOverHead" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageOroOverHead(chat, CharIndex))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(chat, FontIndex))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteRenderValueMsg(ByVal UserIndex As Integer, ByVal x As Byte, ByVal Y As Byte, ByVal rValue As Double, ByVal rType As Byte)

    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateRenderValue(x, Y, rValue, rType))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteLocaleMsg(ByVal UserIndex As Integer, ByVal Id As Integer, ByVal FontIndex As FontTypeNames, Optional ByVal strExtra As String = vbNullString)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageLocaleMsg(Id, strExtra, FontIndex))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteListaCorreo(ByVal UserIndex As Integer, ByVal actualizar As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageListaCorreo(UserIndex, actualizar))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal chat As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildChat" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(chat))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(Message)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteMostrarCuenta(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MostrarCuenta)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(UserIndex)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal heading As eHeading, ByVal CharIndex As Integer, ByVal x As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal name As String, ByVal Status As Byte, ByVal privileges As Byte, ByVal ParticulaFx As Byte, ByVal Head_Aura As String, ByVal Arma_Aura As String, ByVal Body_Aura As String, ByVal Otra_Aura As String, ByVal Escudo_Aura As String, ByVal speeding As Single, ByVal EsNPC As Boolean, ByVal donador As Byte, ByVal appear As Byte, ByVal group_index As Integer, ByVal clan_index As Integer, ByVal clan_nivel As Byte, ByVal UserMinHp As Long, ByVal UserMaxHp As Long, ByVal Simbolo As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterCreate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(Body, Head, heading, CharIndex, x, Y, weapon, shield, FX, FXLoops, helmet, name, Status, privileges, ParticulaFx, Head_Aura, Arma_Aura, Body_Aura, Otra_Aura, Escudo_Aura, speeding, EsNPC, donador, appear, group_index, clan_index, clan_nivel, UserMinHp, UserMaxHp, Simbolo))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal Desvanecido As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterRemove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex, Desvanecido))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal x As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterMove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, x, Y))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal heading As eHeading, ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterChange" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(Body, Head, heading, CharIndex, weapon, shield, FX, FXLoops, helmet))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal x As Byte, ByVal Y As Byte)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectCreate" message to the given user's outgoing data buffer
    '***************************************************

    'If ObjIndex = 251 Then
    ' Debug.Print "Crear la puerta"
    'End If
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(ObjIndex, x, Y))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteParticleFloorCreate(ByVal UserIndex As Integer, ByVal Particula As Integer, ByVal ParticulaTime As Integer, ByVal Map As Integer, ByVal x As Byte, ByVal Y As Byte)

    On Error GoTo Errhandler
  
    If Particula = 0 Then Exit Sub
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageParticleFXToFloor(x, Y, Particula, ParticulaTime))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteLightFloorCreate(ByVal UserIndex As Integer, ByVal LuzColor As Long, ByVal Rango As Byte, ByVal Map As Integer, ByVal x As Byte, ByVal Y As Byte)

    On Error GoTo Errhandler
     
    MapData(Map, x, Y).Luz.Color = LuzColor
    MapData(Map, x, Y).Luz.Rango = Rango

    If Rango = 0 Then Exit Sub
    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageLightFXToFloor(x, Y, LuzColor, Rango))
    Exit Sub
    
Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteFxPiso(ByVal UserIndex As Integer, ByVal GrhIndex As Integer, ByVal x As Byte, ByVal Y As Byte)

    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageFxPiso(GrhIndex, x, Y))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal x As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectDelete" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(x, Y))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, ByVal x As Byte, ByVal Y As Byte, ByVal Blocked As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockPosition" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(x)
        Call .WriteByte(Y)
        Call .WriteBoolean(Blocked)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMidi(ByVal UserIndex As Integer, ByVal midi As Byte, Optional ByVal loops As Integer = -1)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PlayMidi" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMidi(midi, loops))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal UserIndex As Integer, ByVal wave As Integer, ByVal x As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, x, Y))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim Tmp As String

    Dim i   As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AreaChanged" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(UserIndex).Pos.x)
        Call .WriteByte(UserList(UserIndex).Pos.Y)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PauseToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRainToggle())
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteNubesToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageNieblandoToggle(IntensidadDeNubes))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteTrofeoToggleOn(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageTrofeoToggleOn())
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteTrofeoToggleOff(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageTrofeoToggleOff())
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateFX" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
    '***************************************************
    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharUpdateHP(UserIndex))
    Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, PrepareMessageCharUpdateHP(UserIndex))

    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
        Call .WriteLong(UserList(UserIndex).Stats.GLD)
        Call .WriteByte(UserList(UserIndex).Stats.ELV)
        Call .WriteLong(UserList(UserIndex).Stats.ELU)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteUpdateUserKey(ByVal UserIndex As Integer, ByVal slot As Integer, ByVal Llave As Integer)
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserKey)
        Call .WriteInteger(slot)
        Call .WriteInteger(Llave)
    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal UserIndex As Integer, ByVal Skill As eSkill)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.WorkRequestTarget)
        Call .WriteByte(Skill)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

' Writes the "InventoryUnlockSlots" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInventoryUnlockSlots(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Ruthnar
    'Last Modification: 30/09/20
    'Writes the "WriteInventoryUnlockSlots" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.InventoryUnlockSlots)
        Call .WriteByte(UserList(UserIndex).Stats.InventLevel)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteIntervals(ByVal UserIndex As Integer)

    On Error GoTo Errhandler

    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.Intervals)
        Call .outgoingData.WriteLong(.Intervals.Arco)
        Call .outgoingData.WriteLong(.Intervals.Caminar)
        Call .outgoingData.WriteLong(.Intervals.Golpe)
        Call .outgoingData.WriteLong(.Intervals.GolpeMagia)
        Call .outgoingData.WriteLong(.Intervals.magia)
        Call .outgoingData.WriteLong(.Intervals.MagiaGolpe)
        Call .outgoingData.WriteLong(.Intervals.GolpeUsar)
        Call .outgoingData.WriteLong(.Intervals.Trabajar)
        Call .outgoingData.WriteLong(.Intervals.UsarU)
        Call .outgoingData.WriteLong(.Intervals.UsarClic)
        Call .outgoingData.WriteLong(IntervaloTirar)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal slot As Byte)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 3/12/09
    'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
    '3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
    '***************************************************

    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(slot)
        
        Dim ObjIndex As Integer
        
        ObjIndex = UserList(UserIndex).Invent.Object(slot).ObjIndex
        
        Dim PodraUsarlo As Byte
    
        'Ladder
        If ObjIndex > 0 Then
            PodraUsarlo = PuedeUsarObjeto(UserIndex, ObjIndex)
            'PodraUsarlo = IIf(SexoPuedeUsarItem(UserIndex, OBJIndex) = True And UserList(UserIndex).Stats.ELV >= ObjData(OBJIndex).MinELV And ClasePuedeUsarItem(UserIndex, OBJIndex) = True And CheckRazaUsaRopa(UserIndex, OBJIndex) = True, 1, 0)
            'Ladder
    
        End If
    
        Call .WriteInteger(ObjIndex)
        Call .WriteInteger(UserList(UserIndex).Invent.Object(slot).Amount)
        Call .WriteBoolean(UserList(UserIndex).Invent.Object(slot).Equipped)
        Call .WriteSingle(SalePrice(ObjIndex))
        Call .WriteByte(PodraUsarlo)
        Call .WriteByte(UserList(UserIndex).flags.ResistenciaMagica)
        Call .WriteByte(UserList(UserIndex).flags.DañoMagico)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(slot)
        
        Dim ObjIndex As Integer

        Dim obData   As ObjData
        
        ObjIndex = UserList(UserIndex).BancoInvent.Object(slot).ObjIndex
        
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)

        End If
        
        Dim PodraUsarlo As Byte
    
        'Ladder
        If ObjIndex > 0 Then
            PodraUsarlo = PuedeUsarObjeto(UserIndex, ObjIndex)

            'PodraUsarlo = IIf(SexoPuedeUsarItem(UserIndex, OBJIndex) = True And UserList(UserIndex).Stats.ELV >= ObjData(OBJIndex).MinELV = True And ClasePuedeUsarItem(UserIndex, OBJIndex) = True And CheckRazaUsaRopa(UserIndex, OBJIndex) = True, 1, 0)
            'Ladder
        End If

        Call .WriteInteger(UserList(UserIndex).BancoInvent.Object(slot).Amount)
        Call .WriteLong(obData.Valor)
        Call .WriteByte(PodraUsarlo)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal slot As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(slot)
        Call .WriteInteger(UserList(UserIndex).Stats.UserHechizos(slot))
        
        If UserList(UserIndex).Stats.UserHechizos(slot) > 0 Then
            Call .WriteByte(UserList(UserIndex).Stats.UserHechizos(slot))
        Else
            Call .WriteByte("255")

        End If

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Atributes" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.atributes)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
    'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)
        
        For i = 1 To UBound(ArmasHerrero())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) Then
                Count = Count + 1
                validIndexes(Count) = i

            End If

        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            obj = ObjData(ArmasHerrero(validIndexes(i)))
            'Call .WriteASCIIString(obj.Index)
            Call .WriteInteger(ArmasHerrero(validIndexes(i)))
            Call .WriteInteger(obj.LingH)
            Call .WriteInteger(obj.LingP)
            Call .WriteInteger(obj.LingO)
        Next i

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithArmors(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
    'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)
        
        For i = 1 To UBound(ArmadurasHerrero())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).clase), 0) Then
                Count = Count + 1
                validIndexes(Count) = i

            End If

        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            obj = ObjData(ArmadurasHerrero(validIndexes(i)))
            Call .WriteASCIIString(obj.name)
            Call .WriteInteger(obj.LingH)
            Call .WriteInteger(obj.LingP)
            Call .WriteInteger(obj.LingO)
            Call .WriteInteger(ArmadurasHerrero(validIndexes(i)))
        Next i

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i              As Long

    'Dim obj As ObjData
    Dim validIndexes() As Integer

    Dim Count          As Byte
    
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CarpenterObjects)
        
        For i = 1 To UBound(ObjCarpintero())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) Then
                If i = 1 Then Debug.Print UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(UserIndex).clase)
                Count = Count + 1
                validIndexes(Count) = i

            End If

        Next i
        
        ' Write the number of objects in the list
        Call .WriteByte(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Call .WriteInteger(ObjCarpintero(validIndexes(i)))
            'Call .WriteInteger(obj.Madera)
            'Call .WriteLong(obj.GrhIndex)
            ' Ladder 07/07/2014   Ahora se envia el grafico de los objetos
        Next i

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteAlquimistaObjects(ByVal UserIndex As Integer)

    '***************************************************
    On Error GoTo Errhandler

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ObjAlquimista()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlquimistaObj)
        
        For i = 1 To UBound(ObjAlquimista())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjAlquimista(i)).SkPociones <= UserList(UserIndex).Stats.UserSkills(eSkill.Alquimia) \ ModAlquimia(UserList(UserIndex).clase) Then
                'If i = 1 Then Debug.Print UserList(UserIndex).Stats.UserSkills(eSkill.alquimia) \ ModAlquimia(UserList(UserIndex).clase)
                Count = Count + 1
                validIndexes(Count) = i

            End If

        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            'obj = ObjData(ObjAlquimista(validIndexes(i)))
            ' Call .WriteASCIIString(obj.name)
            Call .WriteInteger(ObjAlquimista(validIndexes(i)))
            
        Next i

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteSastreObjects(ByVal UserIndex As Integer)

    '***************************************************
    On Error GoTo Errhandler

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ObjSastre()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SastreObj)
        
        For i = 1 To UBound(ObjSastre())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjSastre(i)).SkMAGOria <= UserList(UserIndex).Stats.UserSkills(eSkill.Sastreria) Then

                ' Round(UserList(UserIndex).Stats.UserSkills(eSkill.Sastreria) / ModSastre(UserList(UserIndex).clase), 0)
                Count = Count + 1
                validIndexes(Count) = i

            End If

        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            ' obj = ObjData(ObjSastre(validIndexes(i)))
            
            'Call .WriteASCIIString(obj.name)
            Call .WriteInteger(ObjSastre(validIndexes(i)))
            ' Call .WriteInteger(obj.PielLobo)
            'Call .WriteInteger(obj.PielOsoPardo)
            'Call .WriteInteger(obj.PielOsoPolaR)
            
            'Call .WriteInteger(ObjSastre(validIndexes(i)))
        Next i

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RestOK" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RestOK)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ErrorMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(Message))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Blind" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Blind)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Dumb" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Dumb)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
'Optimizacion de protocolo por Ladder

Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowSignal" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteInteger(ObjIndex)
        Call .WriteInteger(ObjData(ObjIndex).GrhSecundario)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, ByVal slot As Byte, ByRef obj As obj, ByVal price As Single)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Last Modified by: Nicolas Ezequiel Bouhid (NicoNZ)
    'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim ObjInfo As ObjData
    
    If obj.ObjIndex >= LBound(ObjData()) And obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(obj.ObjIndex)

    End If
    
    Dim PodraUsarlo As Byte
    
    'Ladder
    If obj.ObjIndex > 0 Then
        PodraUsarlo = PuedeUsarObjeto(UserIndex, obj.ObjIndex)

        'PodraUsarlo = IIf(SexoPuedeUsarItem(UserIndex, obj.OBJIndex) = True And UserList(UserIndex).Stats.ELV >= ObjData(obj.OBJIndex).MinELV And ClasePuedeUsarItem(UserIndex, obj.OBJIndex) = True And CheckRazaUsaRopa(UserIndex, obj.OBJIndex) = True, 1, 0)
        'Ladder
    End If
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(slot)
        Call .WriteInteger(obj.ObjIndex)
        Call .WriteInteger(obj.Amount)
        Call .WriteSingle(price)
        Call .WriteByte(PodraUsarlo)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(UserIndex).Stats.MaxAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MinAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MaxHam)
        Call .WriteByte(UserList(UserIndex).Stats.MinHam)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteLight(ByVal UserIndex As Integer, ByVal Map As Integer)

    On Error GoTo Errhandler

    Dim light As String
 
    light = MapInfo(Map).base_light

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.light)
        Call .WriteASCIIString(light)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteEfectToScreen(ByVal UserIndex As Integer, ByVal Color As Long, ByVal Time As Long, Optional ByVal Ignorar As Boolean = False)

    On Error GoTo Errhandler
 
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.EfectToScreen)
        Call .WriteLong(Color)
        Call .WriteLong(Time)
        Call .WriteBoolean(Ignorar)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteFYA(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.FYA)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(1))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(2))
        Call .WriteInteger(UserList(UserIndex).flags.DuracionEfecto)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteCerrarleCliente(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CerrarleCliente)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteOxigeno(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Oxigeno)
        Call .WriteInteger(UserList(UserIndex).Counters.Oxigeno)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteContadores(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Contadores)
        Call .WriteInteger(UserList(UserIndex).Counters.Invisibilidad)
        Call .WriteInteger(UserList(UserIndex).Counters.ScrollExperiencia)
        Call .WriteInteger(UserList(UserIndex).Counters.ScrollOro)

        If UserList(UserIndex).flags.NecesitaOxigeno Then
            Call .WriteInteger(UserList(UserIndex).Counters.Oxigeno)
        Else
            Call .WriteInteger(0)

        End If

        Call .WriteInteger(UserList(UserIndex).flags.DuracionEfecto)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteBindKeys(ByVal UserIndex As Integer)

    '***************************************************
    'Envia los macros al cliente!
    'Por Ladder
    '23/09/2014
    'Flor te amo!
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BindKeys)
        Call .WriteByte(UserList(UserIndex).ChatCombate)
        Call .WriteByte(UserList(UserIndex).ChatGlobal)
        
    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MiniStats" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        
        Call .WriteLong(UserList(UserIndex).Faccion.CiudadanosMatados)
        Call .WriteLong(UserList(UserIndex).Faccion.CriminalesMatados)
        Call .WriteByte(UserList(UserIndex).Faccion.Status)
        
        'TODO : Este valor es calculable, no debería NI EXISTIR, ya sea en el servidor ni en el cliente!!!
        'Call .WriteLong(UserList(UserIndex).Stats.UsuariosMatados)
        
        Call .WriteInteger(UserList(UserIndex).Stats.NPCsMuertos)
        
        Call .WriteByte(UserList(UserIndex).clase)
        Call .WriteLong(UserList(UserIndex).Counters.Pena)
        
        'Ladder 31/07/08  Envio mas estadisticas :P
        Call .WriteLong(UserList(UserIndex).flags.VecesQueMoriste)
        Call .WriteByte(UserList(UserIndex).genero)
        Call .WriteByte(UserList(UserIndex).raza)
        
        Call .WriteByte(UserList(UserIndex).donador.activo)
        Call .WriteLong(CreditosDonadorCheck(UserList(UserIndex).Cuenta))
        'ARREGLANDO
        
        Call .WriteInteger(DiasDonadorCheck(UserList(UserIndex).Cuenta))
        
        Call .WriteLong(UserList(UserIndex).flags.BattlePuntos)
                
    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal UserIndex As Integer, ByVal skillPoints As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LevelUp" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.LevelUp)
        Call .WriteInteger(skillPoints)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data buffer.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, ByVal title As String, ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AddForumMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddForumMsg)
        Call .WriteASCIIString(title)
        Call .WriteASCIIString(Message)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowForumForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowForumForm)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetInvisible" message to the given user's outgoing data buffer
    '***************************************************

    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.
''
' Writes the "DiceRoll" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDiceRoll(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DiceRoll" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.DiceRoll)
        ' TODO: SACAR ESTE PAQUETE USAR EL DE ATRIBUTOS
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MeditateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlindNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DumbNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SendSkills" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SendSkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(UserList(UserIndex).Stats.UserSkills(i))
        Next i

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i   As Long

    Dim str As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To Npclist(NpcIndex).NroCriaturas
            str = str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then str = Left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNews(ByVal UserIndex As Integer, ByVal guildNews As String, ByRef guildList() As String, ByRef MemberList() As String, ByVal ClanNivel As Byte, ByVal ExpAcu As Integer, ByVal ExpNe As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildNews" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildNews)
        Call .WriteASCIIString(guildNews)

        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString

        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
          
        Call .WriteASCIIString(Tmp)
        Call .WriteByte(ClanNivel)
        Call .WriteInteger(ExpAcu)
        Call .WriteInteger(ExpNe)
        
    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OfferDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)
        
        Call .WriteASCIIString(details)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)
        
        ' Prepare guild's list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)
                
        ' Prepare guilds' list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CharacterInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    charName The requested char's name.
' @param    race The requested char's race.
' @param    class The requested char's class.
' @param    gender The requested char's gender.
' @param    level The requested char's level.
' @param    gold The requested char's gold.
' @param    reputation The requested char's reputation.
' @param    previousPetitions The requested char's previous petitions to enter guilds.
' @param    currentGuild The requested char's current guild.
' @param    previousGuilds The requested char's previous guilds.
' @param    RoyalArmy True if tha char belongs to the Royal Army.
' @param    CaosLegion True if tha char belongs to the Caos Legion.
' @param    citicensKilled The number of citicens killed by the requested char.
' @param    criminalsKilled The number of criminals killed by the requested char.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal CharName As String, ByVal race As eRaza, ByVal Class As eClass, ByVal gender As eGenero, ByVal level As Byte, ByVal gold As Long, ByVal bank As Long, ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharacterInfo)
        
        Call .WriteASCIIString(CharName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        Call .WriteByte(gender)
        
        Call .WriteByte(level)
        Call .WriteLong(gold)
        Call .WriteLong(bank)
        
        Call .WriteASCIIString(previousPetitions)
        Call .WriteASCIIString(currentGuild)
        Call .WriteASCIIString(previousGuilds)
        
        Call .WriteBoolean(RoyalArmy)
        Call .WriteBoolean(CaosLegion)
        
        Call .WriteLong(citicensKilled)
        Call .WriteLong(criminalsKilled)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @param    guildNews The guild's news.
' @param    joinRequests The list of chars which requested to join the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeaderInfo(ByVal UserIndex As Integer, ByRef guildList() As String, ByRef MemberList() As String, ByVal guildNews As String, ByRef joinRequests() As String, ByVal NivelDeClan As Byte, ByVal ExpActual As Integer, ByVal ExpNecesaria As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString

        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Store guild news
        Call .WriteASCIIString(guildNews)
        
        ' Prepare the join request's list
        Tmp = vbNullString

        For i = LBound(joinRequests()) To UBound(joinRequests())
            Tmp = Tmp & joinRequests(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Call .WriteByte(NivelDeClan)
        
        Call .WriteInteger(ExpActual)
        Call .WriteInteger(ExpNecesaria)
        
    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "GuildDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildName The requested guild's name.
' @param    founder The requested guild's founder.
' @param    foundationDate The requested guild's foundation date.
' @param    leader The requested guild's current leader.
' @param    URL The requested guild's website.
' @param    memberCount The requested guild's member count.
' @param    electionsOpen True if the clan is electing it's new leader.
' @param    alignment The requested guild's alignment.
' @param    enemiesCount The requested guild's enemy count.
' @param    alliesCount The requested guild's ally count.
' @param    antifactionPoints The requested guild's number of antifaction acts commited.
' @param    codex The requested guild's codex.
' @param    guildDesc The requested guild's description.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDetails(ByVal UserIndex As Integer, ByVal GuildName As String, ByVal founder As String, ByVal foundationDate As String, ByVal leader As String, ByVal memberCount As Integer, ByVal alignment As String, ByVal guildDesc As String, ByVal NivelDeClan As Byte, ByVal ExpActual As Integer, ByVal ExpNecesaria As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i    As Long

    Dim temp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildDetails)
        
        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(founder)
        Call .WriteASCIIString(foundationDate)
        Call .WriteASCIIString(leader)
        
        Call .WriteInteger(memberCount)
        
        Call .WriteASCIIString(alignment)
        
        Call .WriteASCIIString(guildDesc)
        
        Call .WriteByte(NivelDeClan)

        ' Call .WriteInteger(ExpActual)
        ' Call .WriteInteger(ExpNecesaria)
    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/12/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    'Writes the "ParalizeOK" message to the given user's outgoing data buffer
    'And updates user position
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ParalizeOK)

    
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteInmovilizaOK(ByVal UserIndex As Integer)

    '***************************************************
    'Inmovilizar
    'Por Ladder
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.InmovilizaOK)
    '  Call WritePosUpdate(UserIndex)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        
        Call .WriteASCIIString(details)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    Amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)
        
        Call .WriteInteger(ObjIndex)
        Call .WriteASCIIString(ObjData(ObjIndex).name)
        Call .WriteLong(Amount)
        Call .WriteInteger(ObjData(ObjIndex).GrhIndex)
        Call .WriteByte(ObjData(ObjIndex).OBJType)
        Call .WriteInteger(ObjData(ObjIndex).MaxHit)
        Call .WriteInteger(ObjData(ObjIndex).MinHIT)
        Call .WriteInteger(ObjData(ObjIndex).def)
        Call .WriteLong(SalePrice(ObjIndex))

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & i & SEPARATOR
            
        Next i
     
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, ByVal currentMOTD As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMOTDEditionForm)
        
        Call .WriteASCIIString(currentMOTD)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteShowFundarClanForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowFundarClanForm)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06 NIGO:
    'Writes the "UserNameList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer, ByVal Time As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Pong" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Pong)
    Call UserList(UserIndex).outgoingData.WriteLong(Time)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

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
    
100     With UserList(UserIndex).outgoingData

102         If .length = 0 Then Exit Sub
        
            ' Tratamos de enviar los datos.
104         Dim ret As Long: ret = WsApiEnviar(UserIndex, .ReadASCIIStringFixed(.length))
    
            ' Si recibimos un error como respuesta de la API, cerramos el socket.
106         If ret <> 0 And ret <> WSAEWOULDBLOCK Then
                ' Close the socket avoiding any critical error
108             Call CloseSocketSL(UserIndex)
110             Call Cerrar_Usuario(UserIndex)
            End If

        End With

        
        Exit Sub

FlushBuffer_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.FlushBuffer", Erl)
        Resume Next
        
End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, ByVal invisible As Boolean) As String
        
        On Error GoTo PrepareMessageSetInvisible_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "SetInvisible" message and returns it.
        '***************************************************
        'Call WriteContadores(UserIndex)
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.SetInvisible)
        
104         Call .WriteInteger(CharIndex)
106         Call .WriteBoolean(invisible)
        
108         PrepareMessageSetInvisible = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageSetInvisible_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageSetInvisible", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageSetEscribiendo(ByVal CharIndex As Integer, ByVal Escribiendo As Boolean) As String
        
        On Error GoTo PrepareMessageSetEscribiendo_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "SetInvisible" message and returns it.
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.SetEscribiendo)
        
104         Call .WriteInteger(CharIndex)
106         Call .WriteBoolean(Escribiendo)
        
108         PrepareMessageSetEscribiendo = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageSetEscribiendo_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageSetEscribiendo", Erl)
        Resume Next
        
End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal chat As String, ByVal CharIndex As Integer, ByVal Color As Long, Optional ByVal name As String = "") As String
        
        On Error GoTo PrepareMessageChatOverHead_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ChatOverHead" message and returns it.
        '***************************************************
        Dim R, g, b As Byte

100     b = (Color And 16711680) / 65536
102     g = (Color And 65280) / 256
104     R = Color And 255

        'b = color \ 65536
        'g = (color - b * 65536) \ 256
        ' r = color - b * 65536 - g * 256
106     With auxiliarBuffer
108         Call .WriteByte(ServerPacketID.ChatOverHead)
110         Call .WriteASCIIString(chat)
112         Call .WriteInteger(CharIndex)
        
            ' Write rgb channels and save one byte from long :D
114         Call .WriteByte(R)
116         Call .WriteByte(g)
118         Call .WriteByte(b)
120         Call .WriteLong(Color)
        
            'Call .WriteASCIIString(name) Anulado gracias a Optimizacion ^^
        
122         PrepareMessageChatOverHead = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageChatOverHead_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageChatOverHead", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageEfectOverHead(ByVal chat As String, ByVal CharIndex As Integer, Optional ByVal Color As Long = &HFF0000) As String
        
        On Error GoTo PrepareMessageEfectOverHead_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ChatOverHead" message and returns it.
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.EfectOverHead)
104         Call .WriteASCIIString(chat)
106         Call .WriteInteger(CharIndex)
108         Call .WriteLong(Color)
110         PrepareMessageEfectOverHead = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageEfectOverHead_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageEfectOverHead", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageExpOverHead(ByVal chat As String, ByVal CharIndex As Integer, Optional ByVal name As String = "") As String
        
        On Error GoTo PrepareMessageExpOverHead_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ExpOverHEad)
104         Call .WriteASCIIString(chat)
106         Call .WriteInteger(CharIndex)
108         PrepareMessageExpOverHead = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageExpOverHead_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageExpOverHead", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageOroOverHead(ByVal chat As String, ByVal CharIndex As Integer, Optional ByVal name As String = "") As String
        
        On Error GoTo PrepareMessageOroOverHead_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.OroOverHEad)
104         Call .WriteASCIIString(chat)
106         Call .WriteInteger(CharIndex)
108         PrepareMessageOroOverHead = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageOroOverHead_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageOroOverHead", Erl)
        Resume Next
        
End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal chat As String, ByVal FontIndex As FontTypeNames) As String
        
        On Error GoTo PrepareMessageConsoleMsg_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ConsoleMsg" message and returns it.
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ConsoleMsg)
104         Call .WriteASCIIString(chat)
106         Call .WriteByte(FontIndex)
        
108         PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageConsoleMsg_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageConsoleMsg", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageLocaleMsg(ByVal Id As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames) As String
        
        On Error GoTo PrepareMessageLocaleMsg_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ConsoleMsg" message and returns it.
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.LocaleMsg)
104         Call .WriteInteger(Id)
106         Call .WriteASCIIString(chat)
108         Call .WriteByte(FontIndex)
        
110         PrepareMessageLocaleMsg = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageLocaleMsg_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageLocaleMsg", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageListaCorreo(ByVal UserIndex As Integer, ByVal actualizar As Boolean) As String
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ConsoleMsg" message and returns it.
        '***************************************************
        
        On Error GoTo PrepareMessageListaCorreo_Err
        

        Dim cant As Byte

        Dim i    As Byte

100     cant = UserList(UserIndex).Correo.CantCorreo
102     UserList(UserIndex).Correo.NoLeidos = 0

104     With auxiliarBuffer
106         Call .WriteByte(ServerPacketID.ListaCorreo)
108         Call .WriteByte(cant)

110         If cant > 0 Then

112             For i = 1 To cant
114                 Call .WriteASCIIString(UserList(UserIndex).Correo.Mensaje(i).Remitente)
116                 Call .WriteASCIIString(UserList(UserIndex).Correo.Mensaje(i).Mensaje)
118                 Call .WriteByte(UserList(UserIndex).Correo.Mensaje(i).ItemCount)
120                 Call .WriteASCIIString(UserList(UserIndex).Correo.Mensaje(i).Item)

122                 Call .WriteByte(UserList(UserIndex).Correo.Mensaje(i).Leido)
124                 Call .WriteASCIIString(UserList(UserIndex).Correo.Mensaje(i).Fecha)
                    'Call ReadMessageCorreo(UserIndex, i)
126             Next i

            End If

128         Call .WriteBoolean(actualizar)
        
130         PrepareMessageListaCorreo = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageListaCorreo_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageListaCorreo", Erl)
        Resume Next
        
End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer) As String
        
        On Error GoTo PrepareMessageCreateFX_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CreateFX" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.CreateFX)
104         Call .WriteInteger(CharIndex)
106         Call .WriteInteger(FX)
108         Call .WriteInteger(FXLoops)
        
110         PrepareMessageCreateFX = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageCreateFX_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCreateFX", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageMeditateToggle(ByVal CharIndex As Integer, ByVal FX As Integer) As String
        '***************************************************
        
        On Error GoTo PrepareMessageMeditateToggle_Err
        
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.MeditateToggle)
104         Call .WriteInteger(CharIndex)
106         Call .WriteInteger(FX)
        
108         PrepareMessageMeditateToggle = .ReadASCIIStringFixed(.length)
        End With

        
        Exit Function

PrepareMessageMeditateToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageMeditateToggle", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageParticleFX(ByVal CharIndex As Integer, ByVal Particula As Integer, ByVal Time As Long, ByVal Remove As Boolean) As String
        
        On Error GoTo PrepareMessageParticleFX_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CreateFX" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ParticleFX)
104         Call .WriteInteger(CharIndex)
106         Call .WriteInteger(Particula)
108         Call .WriteLong(Time)
110         Call .WriteBoolean(Remove)
        
112         PrepareMessageParticleFX = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageParticleFX_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageParticleFX", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageParticleFXWithDestino(ByVal Emisor As Integer, ByVal Receptor As Integer, ByVal ParticulaViaje As Integer, ByVal ParticulaFinal As Integer, ByVal Time As Long, ByVal wav As Integer, ByVal FX As Integer) As String
        
        On Error GoTo PrepareMessageParticleFXWithDestino_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CreateFX" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ParticleFXWithDestino)
104         Call .WriteInteger(Emisor)
106         Call .WriteInteger(Receptor)
108         Call .WriteInteger(ParticulaViaje)
110         Call .WriteInteger(ParticulaFinal)
112         Call .WriteLong(Time)
114         Call .WriteInteger(wav)
116         Call .WriteInteger(FX)
        
118         PrepareMessageParticleFXWithDestino = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageParticleFXWithDestino_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageParticleFXWithDestino", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageParticleFXWithDestinoXY(ByVal Emisor As Integer, ByVal ParticulaViaje As Integer, ByVal ParticulaFinal As Integer, ByVal Time As Long, ByVal wav As Integer, ByVal FX As Integer, ByVal x As Byte, ByVal Y As Byte) As String
        
        On Error GoTo PrepareMessageParticleFXWithDestinoXY_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CreateFX" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ParticleFXWithDestinoXY)
104         Call .WriteInteger(Emisor)
106         Call .WriteInteger(ParticulaViaje)
108         Call .WriteInteger(ParticulaFinal)
110         Call .WriteLong(Time)
112         Call .WriteInteger(wav)
114         Call .WriteInteger(FX)
116         Call .WriteByte(x)
118         Call .WriteByte(Y)
        
120         PrepareMessageParticleFXWithDestinoXY = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageParticleFXWithDestinoXY_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageParticleFXWithDestinoXY", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageAuraToChar(ByVal CharIndex As Integer, ByVal Aura As String, ByVal Remove As Boolean, ByVal Tipo As Byte) As String
        
        On Error GoTo PrepareMessageAuraToChar_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CreateFX" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.AuraToChar)
104         Call .WriteInteger(CharIndex)
106         Call .WriteASCIIString(Aura)
108         Call .WriteBoolean(Remove)
110         Call .WriteByte(Tipo)
112         PrepareMessageAuraToChar = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageAuraToChar_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageAuraToChar", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageSpeedingACT(ByVal CharIndex As Integer, ByVal speeding As Single) As String
        
        On Error GoTo PrepareMessageSpeedingACT_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CreateFX" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.SpeedTOChar)
104         Call .WriteInteger(CharIndex)
106         Call .WriteSingle(speeding)
108         PrepareMessageSpeedingACT = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageSpeedingACT_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageSpeedingACT", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageParticleFXToFloor(ByVal x As Byte, ByVal Y As Byte, ByVal Particula As Integer, ByVal Time As Long) As String
        
        On Error GoTo PrepareMessageParticleFXToFloor_Err
        

        '***************************************************
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ParticleFXToFloor)
104         Call .WriteByte(x)
106         Call .WriteByte(Y)
108         Call .WriteInteger(Particula)
110         Call .WriteLong(Time)
112         PrepareMessageParticleFXToFloor = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageParticleFXToFloor_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageParticleFXToFloor", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageLightFXToFloor(ByVal x As Byte, ByVal Y As Byte, ByVal LuzColor As Long, ByVal Rango As Byte) As String
        
        On Error GoTo PrepareMessageLightFXToFloor_Err
        

        '***************************************************
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.LightToFloor)
104         Call .WriteByte(x)
106         Call .WriteByte(Y)
108         Call .WriteLong(LuzColor)
110         Call .WriteByte(Rango)
112         PrepareMessageLightFXToFloor = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageLightFXToFloor_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageLightFXToFloor", Erl)
        Resume Next
        
End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Integer, ByVal x As Byte, ByVal Y As Byte) As String
        
        On Error GoTo PrepareMessagePlayWave_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 08/08/07
        'Last Modified by: Rapsodius
        'Added X and Y positions for 3D Sounds
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.PlayWave)
104         Call .WriteInteger(wave)
106         Call .WriteByte(x)
108         Call .WriteByte(Y)
        
110         PrepareMessagePlayWave = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessagePlayWave_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessagePlayWave", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageUbicacionLlamada(ByVal Mapa As Integer, ByVal x As Byte, ByVal Y As Byte) As String
        
        On Error GoTo PrepareMessageUbicacionLlamada_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 08/08/07
        'Last Modified by: Rapsodius
        'Added X and Y positions for 3D Sounds
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.PosLLamadaDeClan)
104         Call .WriteInteger(Mapa)
106         Call .WriteByte(x)
108         Call .WriteByte(Y)
        
110         PrepareMessageUbicacionLlamada = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageUbicacionLlamada_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageUbicacionLlamada", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageCharUpdateHP(ByVal UserIndex As Integer) As String
        
        On Error GoTo PrepareMessageCharUpdateHP_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 08/08/07
        'Last Modified by: Rapsodius
        'Added X and Y positions for 3D Sounds
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.CharUpdateHP)
104         Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
106         Call .WriteInteger(UserList(UserIndex).Stats.MinHp)
108         Call .WriteInteger(UserList(UserIndex).Stats.MaxHp)
        
110         PrepareMessageCharUpdateHP = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageCharUpdateHP_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCharUpdateHP", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageArmaMov(ByVal CharIndex As Integer) As String
        
        On Error GoTo PrepareMessageArmaMov_Err
        

        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ArmaMov)
104         Call .WriteInteger(CharIndex)
        
106         PrepareMessageArmaMov = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageArmaMov_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageArmaMov", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageEscudoMov(ByVal CharIndex As Integer) As String
        
        On Error GoTo PrepareMessageEscudoMov_Err
        

        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.EscudoMov)
104         Call .WriteInteger(CharIndex)
        
106         PrepareMessageEscudoMov = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageEscudoMov_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageEscudoMov", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageEfectToScreen(ByVal Color As Long, ByVal duracion As Long, Optional ByVal Ignorar As Boolean = False) As String
        
        On Error GoTo PrepareMessageEfectToScreen_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 08/08/07
        'Last Modified by: Rapsodius
        'Added X and Y positions for 3D Sounds
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.EfectToScreen)
104         Call .WriteLong(Color)
106         Call .WriteLong(duracion)
108         Call .WriteBoolean(Ignorar)
110         PrepareMessageEfectToScreen = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageEfectToScreen_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageEfectToScreen", Erl)
        Resume Next
        
End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal chat As String) As String
        
        On Error GoTo PrepareMessageGuildChat_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "GuildChat" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.GuildChat)
104         Call .WriteASCIIString(chat)
        
106         PrepareMessageGuildChat = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageGuildChat_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageGuildChat", Erl)
        Resume Next
        
End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal chat As String) As String
        
        On Error GoTo PrepareMessageShowMessageBox_Err
        

        '***************************************************
        'Author: Fredy Horacio Treboux (liquid)
        'Last Modification: 01/08/07
        'Prepares the "ShowMessageBox" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ShowMessageBox)
104         Call .WriteASCIIString(chat)
        
106         PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageShowMessageBox_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageShowMessageBox", Erl)
        Resume Next
        
End Function

''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMidi(ByVal midi As Byte, Optional ByVal loops As Integer = -1) As String
        
        On Error GoTo PrepareMessagePlayMidi_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "GuildChat" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.PlayMidi)
104         Call .WriteByte(midi)
106         Call .WriteInteger(loops)
        
108         PrepareMessagePlayMidi = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessagePlayMidi_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessagePlayMidi", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageOnlineUser(ByVal UserOnline As Integer) As String
        '***************************************************
        
        On Error GoTo PrepareMessageOnlineUser_Err
        

        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.UserOnline)
104         Call .WriteInteger(UserOnline)
        
106         PrepareMessageOnlineUser = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageOnlineUser_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageOnlineUser", Erl)
        Resume Next
        
End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String
        
        On Error GoTo PrepareMessagePauseToggle_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "PauseToggle" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.PauseToggle)
104         PrepareMessagePauseToggle = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessagePauseToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessagePauseToggle", Erl)
        Resume Next
        
End Function

''
' Prepares the "RainToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRainToggle() As String
        
        On Error GoTo PrepareMessageRainToggle_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "RainToggle" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.RainToggle)
        
104         PrepareMessageRainToggle = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageRainToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageRainToggle", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageTrofeoToggleOn() As String
        
        On Error GoTo PrepareMessageTrofeoToggleOn_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "TrofeoToggle" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.TrofeoToggleOn)
        
104         PrepareMessageTrofeoToggleOn = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageTrofeoToggleOn_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageTrofeoToggleOn", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageTrofeoToggleOff() As String
        
        On Error GoTo PrepareMessageTrofeoToggleOff_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "TrofeoToggle" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.TrofeoToggleoff)
        
104         PrepareMessageTrofeoToggleOff = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageTrofeoToggleOff_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageTrofeoToggleOff", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageHora() As String
        
        On Error GoTo PrepareMessageHora_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "RainToggle" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.Hora)
104         Call .WriteInteger(HoraFanstasia)
106         Call .WriteInteger(TimerHoraFantasia)
        
108         PrepareMessageHora = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageHora_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageHora", Erl)
        Resume Next
        
End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal x As Byte, ByVal Y As Byte) As String
        
        On Error GoTo PrepareMessageObjectDelete_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ObjectDelete" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ObjectDelete)
104         Call .WriteByte(x)
106         Call .WriteByte(Y)
        
108         PrepareMessageObjectDelete = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageObjectDelete_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageObjectDelete", Erl)
        Resume Next
        
End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal x As Byte, ByVal Y As Byte, ByVal Blocked As Boolean) As String
        
        On Error GoTo PrepareMessageBlockPosition_Err
        

        '***************************************************
        'Author: Fredy Horacio Treboux (liquid)
        'Last Modification: 01/08/07
        'Prepares the "BlockPosition" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.BlockPosition)
104         Call .WriteByte(x)
106         Call .WriteByte(Y)
108         Call .WriteBoolean(Blocked)
        
110         PrepareMessageBlockPosition = .ReadASCIIStringFixed(.length)

        End With
    
        
        Exit Function

PrepareMessageBlockPosition_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageBlockPosition", Erl)
        Resume Next
        
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.
'Optimizacion por Ladder
Public Function PrepareMessageObjectCreate(ByVal ObjIndex As Integer, ByVal x As Byte, ByVal Y As Byte) As String
        
        On Error GoTo PrepareMessageObjectCreate_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'prepares the "ObjectCreate" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ObjectCreate)
104         Call .WriteByte(x)
106         Call .WriteByte(Y)
108         Call .WriteInteger(ObjIndex)
        
110         PrepareMessageObjectCreate = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageObjectCreate_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageObjectCreate", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageFxPiso(ByVal GrhIndex As Integer, ByVal x As Byte, ByVal Y As Byte) As String
        
        On Error GoTo PrepareMessageFxPiso_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'prepares the "ObjectCreate" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.fxpiso)
104         Call .WriteByte(x)
106         Call .WriteByte(Y)
108         Call .WriteInteger(GrhIndex)
        
110         PrepareMessageFxPiso = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageFxPiso_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageFxPiso", Erl)
        Resume Next
        
End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer, ByVal Desvanecido As Boolean) As String
        
        On Error GoTo PrepareMessageCharacterRemove_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CharacterRemove" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.CharacterRemove)
104         Call .WriteInteger(CharIndex)
106         Call .WriteBoolean(Desvanecido)
        
108         PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageCharacterRemove_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCharacterRemove", Erl)
        Resume Next
        
End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String
        
        On Error GoTo PrepareMessageRemoveCharDialog_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.RemoveCharDialog)
104         Call .WriteInteger(CharIndex)
        
106         PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageRemoveCharDialog_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageRemoveCharDialog", Erl)
        Resume Next
        
End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal Body As Integer, ByVal Head As Integer, ByVal heading As eHeading, ByVal CharIndex As Integer, ByVal x As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal name As String, ByVal Status As Byte, ByVal privileges As Byte, ByVal ParticulaFx As Byte, ByVal Head_Aura As String, ByVal Arma_Aura As String, ByVal Body_Aura As String, ByVal Otra_Aura As String, ByVal Escudo_Aura As String, ByVal speeding As Single, ByVal EsNPC As Boolean, ByVal donador As Byte, ByVal appear As Byte, ByVal group_index As Integer, ByVal clan_index As Integer, ByVal clan_nivel As Byte, ByVal UserMinHp As Long, ByVal UserMaxHp As Long, ByVal Simbolo As Byte) As String
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CharacterCreate" message and returns it
        '***************************************************
        
        On Error GoTo PrepareMessageCharacterCreate_Err
        

100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.CharacterCreate)
        
104         Call .WriteInteger(CharIndex)
106         Call .WriteInteger(Body)
108         Call .WriteInteger(Head)
110         Call .WriteByte(heading)
112         Call .WriteByte(x)
114         Call .WriteByte(Y)
116         Call .WriteInteger(weapon)
118         Call .WriteInteger(shield)
120         Call .WriteInteger(helmet)
122         Call .WriteInteger(FX)
124         Call .WriteInteger(FXLoops)
126         Call .WriteASCIIString(name)
128         Call .WriteByte(Status)
130         Call .WriteByte(privileges)
132         Call .WriteByte(ParticulaFx)
134         Call .WriteASCIIString(Head_Aura)
136         Call .WriteASCIIString(Arma_Aura)
138         Call .WriteASCIIString(Body_Aura)
140         Call .WriteASCIIString(Otra_Aura)
142         Call .WriteASCIIString(Escudo_Aura)
144         Call .WriteSingle(speeding)
146         Call .WriteBoolean(EsNPC)
148         Call .WriteByte(donador)
150         Call .WriteByte(appear)
152         Call .WriteInteger(group_index)
154         Call .WriteInteger(clan_index)
156         Call .WriteByte(clan_nivel)
158         Call .WriteLong(UserMinHp)
160         Call .WriteLong(UserMaxHp)
162         Call .WriteByte(Simbolo)

164         PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageCharacterCreate_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCharacterCreate", Erl)
        Resume Next
        
End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal Body As Integer, ByVal Head As Integer, ByVal heading As eHeading, ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer) As String
        
        On Error GoTo PrepareMessageCharacterChange_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CharacterChange" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.CharacterChange)
        
104         Call .WriteInteger(CharIndex)
106         Call .WriteInteger(Body)
108         Call .WriteInteger(Head)
110         Call .WriteByte(heading)
112         Call .WriteInteger(weapon)
114         Call .WriteInteger(shield)
116         Call .WriteInteger(helmet)
118         Call .WriteInteger(FX)
120         Call .WriteInteger(FXLoops)
        
122         PrepareMessageCharacterChange = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageCharacterChange_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCharacterChange", Erl)
        Resume Next
        
End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal x As Byte, ByVal Y As Byte) As String
        
        On Error GoTo PrepareMessageCharacterMove_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CharacterMove" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.CharacterMove)
104         Call .WriteInteger(CharIndex)
106         Call .WriteByte(x)
108         Call .WriteByte(Y)
        
110         PrepareMessageCharacterMove = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageCharacterMove_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCharacterMove", Erl)
        Resume Next
        
End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, Status As Byte, Tag As String) As String
        
        On Error GoTo PrepareMessageUpdateTagAndStatus_Err
        

        '***************************************************
        'Author: Alejandro Salvo (Salvito)
        'Last Modification: 04/07/07
        'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
        'Prepares the "UpdateTagAndStatus" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        
104         Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
106         Call .WriteByte(Status)
108         Call .WriteASCIIString(Tag)
110         Call .WriteInteger(UserList(UserIndex).Grupo.Lider)
        
112         PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageUpdateTagAndStatus_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageUpdateTagAndStatus", Erl)
        Resume Next
        
End Function

Public Sub WriteUpdateNPCSimbolo(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Simbolo As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateNPCSimbolo)
        Call .WriteInteger(Npclist(NpcIndex).Char.CharIndex)
        Call .WriteByte(Simbolo)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal Message As String) As String
        
        On Error GoTo PrepareMessageErrorMsg_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ErrorMsg" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ErrorMsg)
104         Call .WriteASCIIString(Message)
        
106         PrepareMessageErrorMsg = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageErrorMsg_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageErrorMsg", Erl)
        Resume Next
        
End Function

Private Sub HandleQuestionGM(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim Consulta       As String

        Dim TipoDeConsulta As String

        Consulta = buffer.ReadASCIIString()
        TipoDeConsulta = buffer.ReadASCIIString()

        If UserList(UserIndex).donador.activo = 1 Then
            Call Ayuda.Push(.name, Consulta, TipoDeConsulta & "-Prioritario")
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Se ha recibido un nuevo mensaje de soporte de " & UserList(UserIndex).name & "(Prioritario).", FontTypeNames.FONTTYPE_SERVER))
            
        Else
            Call Ayuda.Push(.name, Consulta, TipoDeConsulta)
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Se ha recibido un nuevo mensaje de soporte de " & UserList(UserIndex).name & ".", FontTypeNames.FONTTYPE_SERVER))

        End If

        Call WriteConsoleMsg(UserIndex, "Tu mensaje fue recibido por el equipo de soporte.", FontTypeNames.FONTTYPE_INFOIAO)
        'Call WriteConsoleMsg(UserIndex, "Tu mensaje fue recibido por el equipo de soporte.", FontTypeNames.FONTTYPE_INFOIAO)
        
        Call LogConsulta(.name & "(" & TipoDeConsulta & ") " & Consulta)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleOfertaInicial(ByVal UserIndex As Integer)
        
        On Error GoTo HandleOfertaInicial_Err
        

        'Author: Pablo Mercavides
100     If UserList(UserIndex).incomingData.length < 6 Then
102         Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(UserIndex)
            'Remove packet ID
106         Call .incomingData.ReadInteger

            Dim Oferta As Long

108         Oferta = .incomingData.ReadLong()
        
110         If UserList(UserIndex).flags.Muerto = 1 Then
112             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
                
                Exit Sub

            End If

114         If .flags.TargetNPC < 1 Then
116             Call WriteConsoleMsg(UserIndex, "Primero tenés que hacer click sobre el subastador.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

118         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Subastador Then
120             Call WriteConsoleMsg(UserIndex, "Primero tenés que hacer click sobre el subastador.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
122         If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 2 Then
124             Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del subastador.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
126         If .flags.Subastando = False Then
128             Call WriteChatOverHead(UserIndex, "Ollí amigo, tu no podés decirme cual es la oferta inicial.", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If
        
130         If Subasta.HaySubastaActiva = False And .flags.Subastando = False Then
132             Call WriteConsoleMsg(UserIndex, "No hay ninguna subasta en curso.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
134         If .flags.Subastando = True Then
136             UserList(UserIndex).Counters.TiempoParaSubastar = 0
138             Subasta.OfertaInicial = Oferta
140             Subasta.MejorOferta = 0
142             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " estí subastando: " & ObjData(Subasta.ObjSubastado).name & " (Cantidad: " & Subasta.ObjSubastadoCantidad & " ) - con un precio inicial de " & Subasta.OfertaInicial & " monedas. Escribe /OFERTAR (cantidad) para participar.", FontTypeNames.FONTTYPE_SUBASTA))
144             .flags.Subastando = False
146             Subasta.HaySubastaActiva = True
148             Subasta.Subastador = .name
150             Subasta.MinutosDeSubasta = 5
152             Subasta.TiempoRestanteSubasta = 300
154             Call LogearEventoDeSubasta("#################################################################################################################################################################################################")
156             Call LogearEventoDeSubasta("El dia: " & Date & " a las " & Time)
158             Call LogearEventoDeSubasta(.name & ": Esta subastando el item numero " & Subasta.ObjSubastado & " con una cantidad de " & Subasta.ObjSubastadoCantidad & " y con un precio inicial de " & Subasta.OfertaInicial & " monedas.")
160             frmMain.SubastaTimer.Enabled = True
162             Call WarpUserChar(UserIndex, 14, 27, 64, True)

                'lalala toda la bola de los timerrr
            End If

        End With

        
        Exit Sub

HandleOfertaInicial_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleOfertaInicial", Erl)
        Resume Next
        
End Sub

Private Sub HandleOfertaDeSubasta(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Dim Oferta   As Long

        Dim ExOferta As Long
        
        Oferta = buffer.ReadLong()
        
        If Subasta.HaySubastaActiva = False Then
            Call WriteConsoleMsg(UserIndex, "No hay ninguna subasta en curso.", FontTypeNames.FONTTYPE_INFOIAO)
            Call .incomingData.CopyBuffer(buffer)
            Exit Sub

        End If
        
        If UserList(UserIndex).flags.BattleModo = 1 Then
            Call WriteConsoleMsg(UserIndex, "Subastador > íComo vas a ofertar con dinero que no es tuyo? Bríbon.", FontTypeNames.FONTTYPE_INFOIAO)
            Call .incomingData.CopyBuffer(buffer)
            Exit Sub

        End If
        
        If Oferta < Subasta.MejorOferta + 100 Then
            Call WriteConsoleMsg(UserIndex, "Debe haber almenos una diferencia de 100 monedas a la ultima oferta!", FontTypeNames.FONTTYPE_INFOIAO)
            Call .incomingData.CopyBuffer(buffer)
            Exit Sub

        End If
        
        If .name = Subasta.Subastador Then
            Call WriteConsoleMsg(UserIndex, "No podés auto ofertar en tus subastas. La proxima vez iras a la carcel...", FontTypeNames.FONTTYPE_INFOIAO)
            Call .incomingData.CopyBuffer(buffer)
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
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Oferta mejorada por: " & .name & " (Ofrece " & Oferta & " Monedas de oro) - Tiempo Extendido. Escribe /SUBASTA para mas informaciín.", FontTypeNames.FONTTYPE_SUBASTA))
                Call LogearEventoDeSubasta(.name & ": Mejoro la oferta en el ultimo minuto ofreciendo " & Oferta & " monedas.")
                Subasta.TiempoRestanteSubasta = Subasta.TiempoRestanteSubasta + 30
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Oferta mejorada por: " & .name & " (Ofrece " & Oferta & " Monedas de oro). Escribe /SUBASTA para mas informaciín.", FontTypeNames.FONTTYPE_SUBASTA))
                Call LogearEventoDeSubasta(.name & ": Mejoro la oferta ofreciendo " & Oferta & " monedas.")
                Subasta.HuboOferta = True
                Subasta.PosibleCancelo = False

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "No posees esa cantidad de oro.", FontTypeNames.FONTTYPE_INFOIAO)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleGlobalMessage(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String

        chat = buffer.ReadASCIIString()

        If .flags.Silenciado = 1 Then
            Call WriteLocaleMsg(UserIndex, "110", FontTypeNames.FONTTYPE_VENENO, .flags.MinutosRestantes)
            'Call WriteConsoleMsg(UserIndex, "Los administradores te han impedido hablar durante los proximos " & .flags.MinutosRestantes & " minutos debido a tu comportamiento.", FontTypeNames.FONTTYPE_VENENO)
        Else

            If EstadoGlobal Then
                If LenB(chat) <> 0 Then
                    'Analize chat...
                    Call Statistics.ParseChat(chat)
                    Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[" & .name & "] " & chat, FontTypeNames.FONTTYPE_GLOBAL))

                    'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
                    'Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.map, "||" & vbBlue & "í< " & rData & " >í" & CStr(UserList(UserIndex).Char.CharIndex))
                End If

            Else
                Call WriteConsoleMsg(UserIndex, "El global se encuentra Desactivado.", FontTypeNames.FONTTYPE_GLOBAL)

            End If

        End If
    
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub HandleGlobalOnOff(ByVal UserIndex As Integer)
        
        On Error GoTo HandleGlobalOnOff_Err
        

        'Author: Pablo Mercavides
100     With UserList(UserIndex)
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
106         Call LogGM(.name, "/GLOBAL")
        
108         If EstadoGlobal = False Then
110             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Chat general habilitado. Escribe" & Chr(34) & "/CONSOLA" & Chr(34) & " o " & Chr(34) & ";" & Chr(34) & " y su mensaje para utilizarlo.", FontTypeNames.FONTTYPE_SERVER))
112             EstadoGlobal = True
            Else
114             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Chat General deshabilitado.", FontTypeNames.FONTTYPE_SERVER))
116             EstadoGlobal = False

            End If
        
        End With

        
        Exit Sub

HandleGlobalOnOff_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleGlobalOnOff", Erl)
        Resume Next
        
End Sub

Public Sub SilenciarUserName(ByVal SilencioUserIndex As Integer, ByVal UserName As String, ByVal Time As Byte)
        
        On Error GoTo SilenciarUserName_Err
        

        'Author: Pablo Mercavides
        Dim tUser     As Integer

        Dim userPriv  As Byte

        Dim cantPenas As Byte

        Dim rank      As Integer
    
100     If InStrB(UserName, "+") Then
102         UserName = Replace(UserName, "+", " ")

        End If
    
104     tUser = NameIndex(UserName)
    
106     rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
108     With UserList(SilencioUserIndex)

110         If tUser <= 0 Then
112             Call WriteConsoleMsg(SilencioUserIndex, "El usuario no esta online, pena grabada en el charfile.", FontTypeNames.FONTTYPE_TALK)
            
114             If FileExist(CharPath & UserName & ".chr", vbNormal) Then
116                 userPriv = UserDarPrivilegioLevel(UserName)
                
118                 If (userPriv And rank) > (.flags.Privilegios And rank) Then
120                     Call WriteConsoleMsg(SilencioUserIndex, "No podes silenciar a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        
                        'ponemos el flag de silencio a 1 y los minutos
122                     Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Silenciado", "1")
124                     Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "MinutosRestantes", Time)
126                     Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "SegundosPasados", "0")
                        
                        'ponemos la pena
128                     cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
130                     Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
132                     Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.name) & ": Silenciado durante " & Time & " minutos. " & Date & " ")

                    End If

                Else
134                 Call WriteConsoleMsg(SilencioUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

136             If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
138                 Call WriteConsoleMsg(SilencioUserIndex, "No podes silenciar a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)

                End If
            
140             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .name & " ha silenciado a " & UserList(tUser).name & ", por " & Time & " minutos.", FontTypeNames.FONTTYPE_SERVER))
            
                'Ponemos el flag de ban a 1
142             UserList(tUser).flags.Silenciado = 1
144             UserList(tUser).flags.MinutosRestantes = Time
146             UserList(tUser).flags.SegundosPasados = 0
148             Call LogGM(.name, "Silencio a " & UserName)
            
                'ponemos el flag de silencio
150             Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Silenciado", "1")
152             Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "MinutosRestantes", Time)
154             Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "SegundosPasados", "0")
                'ponemos la pena
156             cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
158             Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
160             Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.name) & ": silencio por " & Time & " minutos. " & Date & " " & Time)
                'Call WriteConsoleMsg(tUser, "Has sido silenciado durante " & Time & " minutos.", FontTypeNames.FONTTYPE_INFO)
162             Call WriteLocaleMsg(tUser, "11", FontTypeNames.FONTTYPE_VENENO)

            End If

        End With

        
        Exit Sub

SilenciarUserName_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.SilenciarUserName", Erl)
        Resume Next
        
End Sub

Private Sub HandleCrearCuenta(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 18 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo Errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim CuentaEmail    As String

    Dim CuentaPassword As String
    
    CuentaEmail = buffer.ReadASCIIString()
    CuentaPassword = buffer.ReadASCIIString()
  
    If Not CheckMailString(CuentaEmail) Then
        Call WriteErrorMsg(UserIndex, "Email inválido.")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If

    If Not CuentaExiste(CuentaEmail) Then

        Call SaveNewAccount(UserIndex, CuentaEmail, SDesencriptar(CuentaPassword))
    
        Call EnviarCorreo(CuentaEmail)
        Call WriteShowFrmLogear(UserIndex)
        Call WriteShowMessageBox(UserIndex, "Cuenta creada. Se ha enviado un código de validación a su email, debe activar la cuenta antes de poder usarla. Recuerde revisar SPAM en caso de no encontrar el mail.")
        
        Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    Else
        Call WriteShowMessageBox(UserIndex, "El email ya está en uso.")
        
        Call CloseSocket(UserIndex)

    End If
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleValidarCuenta(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 7 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo Errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim CuentaEmail    As String

    Dim ValidacionCode As String
    
    CuentaEmail = buffer.ReadASCIIString()
    ValidacionCode = buffer.ReadASCIIString()

    If Not CheckMailString(CuentaEmail) Then
        Call WriteShowFrmLogear(UserIndex)
        Call WriteShowMessageBox(UserIndex, "Email inválido.")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    If CuentaExiste(CuentaEmail) Then
        If Not ObtenerValidacion(CuentaEmail) Then
            If UCase$(ObtenerCodigo(CuentaEmail)) = UCase$(ValidacionCode) Then
                If Database_Enabled Then
                    Call ValidarCuentaDatabase(CuentaEmail)
                Else
                    Call WriteVar(CuentasPath & LCase$(CuentaEmail) & ".act", "INIT", "Activada", "1")

                End If

                Call WriteShowFrmLogear(UserIndex)
                Call WriteShowMessageBox(UserIndex, "Cuenta activada con éxito, ya puede ingresar.")
            Else
                Call WriteShowFrmLogear(UserIndex)
                Call WriteShowMessageBox(UserIndex, "¡Código de activación inválido!")

            End If

        Else
            Call WriteShowFrmLogear(UserIndex)
            Call WriteShowMessageBox(UserIndex, "La cuenta ya ha sido validada anteriormente.")

        End If

    Else
        Call WriteShowFrmLogear(UserIndex)
        Call WriteShowMessageBox(UserIndex, "La cuenta no existe.")

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleReValidarCuenta(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo Errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserCuenta As String

    Dim Useremail  As String
    
    UserCuenta = buffer.ReadASCIIString()
    
    Useremail = buffer.ReadASCIIString()
    
    'WyroX: TODO:
    'Saco este paquete, por el momento
    Exit Sub
    
    If Not AsciiValidos(UserCuenta) Then
        Call WriteShowFrmLogear(UserIndex)
        Call WriteShowMessageBox(UserIndex, "Nombre invalido.")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    'If Useremail <> ObtenerEmail(UserCuenta) Then
    Call WriteShowFrmLogear(UserIndex)
    Call WriteShowMessageBox(UserIndex, "El email introducido no coincide con el email registrador.")
    
    Call CloseSocket(UserIndex)
    Exit Sub
    'End If
    
    If CuentaExiste(UserCuenta) Then
        If ObtenerValidacion(UserCuenta) = 0 Then
            'Call EnviarCorreo(UserCuenta, ObtenerEmail(UserCuenta))
            Call WriteShowFrmLogear(UserIndex)
            Call WriteShowMessageBox(UserIndex, "Se ha enviado el mail de validación a la dirección designada cuando se creo la cuenta.")
                
        Else
            Call WriteShowFrmLogear(UserIndex)
            Call WriteShowMessageBox(UserIndex, "La cuenta ya ha sido validada anteriormente.")

        End If

    Else
        Call WriteShowFrmLogear(UserIndex)
        Call WriteShowMessageBox(UserIndex, "La cuenta no existe.")

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleIngresarConCuenta(ByVal UserIndex As Integer)

    Dim Version As String

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 14 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo Errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim CuentaEmail    As String

    Dim CuentaPassword As String

    Dim MacAddress     As String

    Dim HDserial       As Long
    
    CuentaEmail = buffer.ReadASCIIString()
    CuentaPassword = buffer.ReadASCIIString()
    Version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())

    If Not VersionOK(Version) Then
        Call WriteShowMessageBox(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    'If ServerSoloGMs > 0 Then
    '    If Condicion Then
    '        Call WriteShowMessageBox(UserIndex, "El servidor sera habilitado a las 18 horas. Por el momento solo podes crear cuentas. Te esperamos!")
    '
    '        Call CloseSocket(UserIndex)
    '        Exit Sub
    '    End If
    'End If
    
    'Seguridad Ladder
    MacAddress = buffer.ReadASCIIString()
    HDserial = buffer.ReadLong()
    'Seguridad Ladder
    
  

    
    If ServerSoloGMs > 0 Then
    
     
        Select Case LCase$(CuentaEmail)
        
           Case "djpablo@djpablo.com.ar" ' Ladder
           
           Case "reformasapef@gmail.com" 'Sensui
            
           Case "gulfas@gmail.com" 'Morgolock
            
           Case "alexiscaraballo96@gmail.com" 'Wyrox
            
           Case "jopiodz00@gmail.com" 'jopi
    
           Case "juanitonelli@gmail.com" 'Danv
           
           
           Case "lucas.recoaro@gmail.com" 'Recox
           
           
           Case "reyarb@fibertel.com.ar" 'Recox
           
           
           Case "hgarofalo79@gmail.com" 'Haracin
        
            Case Else
                    Call WriteShowMessageBox(UserIndex, "El servidor se encuentra habilitado solo para administradores. ¡Te esperamos pronto!")
                    Call FlushBuffer(UserIndex)
                    Call CloseSocket(UserIndex)
                    Exit Sub
            End Select
    End If
    
    
    
    If EntrarCuenta(UserIndex, CuentaEmail, CuentaPassword, MacAddress, HDserial) Then
        Call WritePersonajesDeCuenta(UserIndex)
        Call WriteMostrarCuenta(UserIndex)
    Else
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBorrarPJ(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 15 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo Errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserDelete     As String

    Dim CuentaEmail    As String

    Dim CuentaPassword As String

    Dim MacAddress     As String

    Dim HDserial       As Long

    Dim Version        As String
    
    UserDelete = buffer.ReadASCIIString()
    CuentaEmail = buffer.ReadASCIIString()
    CuentaPassword = buffer.ReadASCIIString()
    Version = CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte()) & "." & CStr(buffer.ReadByte())
    
    If Not VersionOK(Version) Then
        Call WriteShowMessageBox(UserIndex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If

    MacAddress = buffer.ReadASCIIString()
    HDserial = buffer.ReadLong()
    
    If Not EntrarCuenta(UserIndex, CuentaEmail, CuentaPassword, MacAddress, HDserial) Then
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    If Not AsciiValidos(UserDelete) Then
        Call WriteShowMessageBox(UserIndex, "Nombre inválido.")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    If Database_Enabled Then
        Call BorrarUsuarioDatabase(UserDelete)
    Else

        If PersonajeExiste(UserDelete) Then
            Call FileCopy(CharPath & UserDelete & ".chr", DeletePath & UCase$(UserDelete) & ".chr")
         
            Call BorrarPJdeCuenta(UserDelete)
        
            'Call WriteShowMessageBox(UserIndex, "El personaje " & UserDelete & " a sido borrado de la cuenta.")
            Call Kill(CharPath & UserDelete & ".chr")

        End If

    End If
    
    Call WritePersonajesDeCuenta(UserIndex)
  
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBorrandoCuenta(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 7 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo Errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim AccountDelete As String

    Dim UserMail      As String

    Dim Password      As String
    
    AccountDelete = buffer.ReadASCIIString()
    UserMail = buffer.ReadASCIIString()
    Password = buffer.ReadASCIIString()
    
    If CuentaExiste(AccountDelete) Then
    
        If Not AsciiValidos(AccountDelete) Then
            Call WriteShowFrmLogear(UserIndex)
            Call WriteShowMessageBox(UserIndex, "Cuenta invalida.")
            
            Call CloseSocket(UserIndex)
            Exit Sub

        End If
        
        If UserMail <> ObtenerEmail(AccountDelete) Then
            Call WriteShowFrmLogear(UserIndex)
            Call WriteShowMessageBox(UserIndex, "El email introducido no coincide con el email registrador.")
            
            Call CloseSocket(UserIndex)
            Exit Sub

        End If
        
        If True Then ' Desactivado
            Call WriteShowFrmLogear(UserIndex)
            Call WriteShowMessageBox(UserIndex, "La contraseña introducida no es correcta.")
            
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        Call BorrarCuenta(AccountDelete)
        Call WriteShowFrmLogear(UserIndex)
        Call WriteShowMessageBox(UserIndex, "La cuenta ha sido borrada.")
        
    Else
        Call WriteShowFrmLogear(UserIndex)
        Call WriteShowMessageBox(UserIndex, "La cuenta ingresada no existe.")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleRecuperandoContraseña(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo Errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim AcountDelete As String

    Dim UserMail     As String
    
    AcountDelete = buffer.ReadASCIIString()
    UserMail = buffer.ReadASCIIString()
    
    If FileExist(CuentasPath & UCase$(AcountDelete) & ".act", vbNormal) Then
    
        If Not AsciiValidos(AcountDelete) Then
            Call WriteShowFrmLogear(UserIndex)
            Call WriteShowMessageBox(UserIndex, "Cuenta invalida.")
            
            
            Call CloseSocket(UserIndex)
            Exit Sub

        End If

        If UserMail <> ObtenerEmail(AcountDelete) Then
            Call WriteShowFrmLogear(UserIndex)
            Call WriteShowMessageBox(UserIndex, "El email introducido no coincide con el email registrador.")
            
            Call CloseSocket(UserIndex)
            Exit Sub

        End If
        
        If EnviarCorreoRecuperacion(AcountDelete, ObtenerEmail(AcountDelete)) Then
            Call WriteShowFrmLogear(UserIndex)
            Call WriteShowMessageBox(UserIndex, "La contraseña de la cuenta a sido enviada por email a la direccion registrada.")
        Else
            Call WriteShowFrmLogear(UserIndex)
            Call WriteShowMessageBox(UserIndex, "Se ha provocado un error al recuperar la clave, reintente mas tarde.")

        End If

    Else
        Call WriteShowFrmLogear(UserIndex)
        Call WriteShowMessageBox(UserIndex, "La cuenta ingresada no existe.")
        
        Call CloseSocket(UserIndex)
        Exit Sub

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(buffer)
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub WritePersonajesDeCuenta(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    
    Dim UserCuenta                     As String

    Dim CantPersonajes                 As Byte

    Dim Personaje(1 To MAX_PERSONAJES) As PersonajeCuenta

    Dim donador                        As Boolean

    Dim i                              As Byte
    
    UserCuenta = UserList(UserIndex).Cuenta
    
    donador = DonadorCheck(UserCuenta)

    If Database_Enabled Then
        CantPersonajes = GetPersonajesCuentaDatabase(UserList(UserIndex).AccountID, Personaje)
    Else
        CantPersonajes = ObtenerCantidadDePersonajes(UserCuenta)
        
        For i = 1 To CantPersonajes
            Personaje(i).nombre = ObtenerNombrePJ(UserCuenta, i)
            Personaje(i).Cabeza = ObtenerCabeza(Personaje(i).nombre)
            Personaje(i).clase = ObtenerClase(Personaje(i).nombre)
            Personaje(i).cuerpo = ObtenerCuerpo(Personaje(i).nombre)
            Personaje(i).Mapa = ReadField(1, ObtenerMapa(Personaje(i).nombre), Asc("-"))
            Personaje(i).nivel = ObtenerNivel(Personaje(i).nombre)
            Personaje(i).Status = ObtenerCriminal(Personaje(i).nombre)
            Personaje(i).Casco = ObtenerCasco(Personaje(i).nombre)
            Personaje(i).Escudo = ObtenerEscudo(Personaje(i).nombre)
            Personaje(i).Arma = ObtenerArma(Personaje(i).nombre)
            Personaje(i).ClanIndex = GetUserGuildIndexCharfile(Personaje(i).nombre)
        Next i

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PersonajesDeCuenta)
        Call .WriteByte(CantPersonajes)
            
        For i = 1 To CantPersonajes
            Call .WriteASCIIString(Personaje(i).nombre)
            Call .WriteByte(Personaje(i).nivel)
            Call .WriteInteger(Personaje(i).Mapa)
            Call .WriteInteger(Personaje(i).cuerpo)
            Call .WriteInteger(Personaje(i).Cabeza)
            Call .WriteByte(Personaje(i).Status)
            Call .WriteByte(Personaje(i).clase)
            Call .WriteInteger(Personaje(i).Casco)
            Call .WriteInteger(Personaje(i).Escudo)
            Call .WriteInteger(Personaje(i).Arma)
            Call .WriteASCIIString(modGuilds.GuildName(Personaje(i).ClanIndex))
        Next i
            
        Call .WriteByte(IIf(donador, 1, 0))

    End With
    
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleCuentaRegresiva(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim Seconds As Byte
        
        Seconds = buffer.ReadByte()

        If Not .flags.Privilegios And PlayerType.user Then
            CuentaRegresivaTimer = Seconds
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("íEmpezando cuenta regresiva desde: " & Seconds & " segundos...!", FontTypeNames.FONTTYPE_GUILD))
        
            'If we got here then packet is complete, copy data back to original queue
        End If

        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandlePossUser(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String
        
        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.user Then
        
            If Database_Enabled Then
                Call SetPositionDatabase(UserName, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y)
            Else
                Call WriteVar(CharPath & UCase$(UserName) & ".chr", "INIT", "Position", UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.x & "-" & UserList(UserIndex).Pos.Y)
            End If

            Call WriteConsoleMsg(UserIndex, "Servidor> Acción realizada con exito! La nueva posicion de " & UserName & "es: " & UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.x & "-" & UserList(UserIndex).Pos.Y & "...", FontTypeNames.FONTTYPE_INFO)

            ' Call SendData(UserIndex, UserIndex, PrepareMessageConsoleMsg("Acciín realizada con exito! La nueva posicion de " & UserName & "es: " & UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.y & "...", FontTypeNames.FONTTYPE_SERVER))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleDuelo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As clsByteQueue
        Set buffer = New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
        MapaOcupado = False

        Dim UserRetado As Integer: UserRetado = .flags.TargetUser

        If .flags.SolicitudPendienteDe = 0 Then
            
            Select Case UserRetado
            
                Case 0
                    Call WriteConsoleMsg(UserIndex, "Duelos> Primero haz click sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                
                Case Is < 0
                    Call WriteConsoleMsg(UserIndex, "Duelos> ¡El persona se encuentra offline!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                
            End Select

            If MapaOcupado Then
                Call WriteConsoleMsg(UserIndex, "Duelos> El mapa de duelos esta ocupado, intentalo mas tarde.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            .flags.RetoA = UserList(UserRetado).name
            UserList(UserRetado).flags.SolicitudPendienteDe = .name
        
            Call WriteConsoleMsg(UserRetado, "Duelos> Has sido retado a duelo por " & .name & " si quieres aceptar el duelo escribe /DUELO.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Duelos> La solicitud a sido enviada al usuario, ahora debes esperar la respuesta de " & UserList(UserRetado).name & ".", FontTypeNames.FONTTYPE_INFO)
               
        Else

           Exit Sub

        End If

        Call SendData(UserIndex, 0, PrepareMessageConsoleMsg("Duelo comenzado!", FontTypeNames.FONTTYPE_SERVER))

    End With
    
    Exit Sub
    
Errhandler:

    Dim Error As Long: Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub WriteGoliathInit(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Goliath)
        Call .WriteLong(UserList(UserIndex).Stats.Banco)
        Call .WriteByte(UserList(UserIndex).BancoInvent.NroItems)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteShowFrmLogear(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowFrmLogear)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteShowFrmMapa(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowFrmMapa)
        
        If UserList(UserIndex).donador.activo = 1 Then
            Call .WriteInteger(ExpMult * UserList(UserIndex).flags.ScrollExp * 1.1)
        Else
            Call .WriteInteger(ExpMult * UserList(UserIndex).flags.ScrollExp)

        End If

        Call .WriteInteger(OroMult * UserList(UserIndex).flags.ScrollOro)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleNieveToggle(ByVal UserIndex As Integer)
        
        On Error GoTo HandleNieveToggle_Err
        

        'Author: Pablo Mercavides
100     With UserList(UserIndex)
102         Call .incomingData.ReadInteger 'Remove packet ID
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub
        
106         Call LogGM(.name, "/NIEVE")
108         Nebando = Not Nebando
        
110         Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())

        End With

        
        Exit Sub

HandleNieveToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleNieveToggle", Erl)
        Resume Next
        
End Sub

Private Sub HandleNieblaToggle(ByVal UserIndex As Integer)
        
        On Error GoTo HandleNieblaToggle_Err
        

        'Author: Pablo Mercavides
100     With UserList(UserIndex)
102         Call .incomingData.ReadInteger 'Remove packet ID

104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub
        
106         Call LogGM(.name, "/NIEBLA")
108         Call ResetMeteo
            'Nieblando = Not Nieblando
        
            ' Call SendData(SendTarget.ToAll, 0, PrepareMessageNieblandoToggle())
        End With

        
        Exit Sub

HandleNieblaToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleNieblaToggle", Erl)
        Resume Next
        
End Sub

Private Sub HandleTransFerGold(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 8 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String

        Dim Cantidad As Long

        Dim tUser    As Integer
        
        Cantidad = buffer.ReadLong()
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
            If tUser <= 0 Then

                If Database_Enabled Then
                    Call AddOroBancoDatabase(UserName, Cantidad)
                Else
                    Dim FileUser  As String
                    Dim OroenBove As Long

                    FileUser = CharPath & UCase$(UserName) & ".chr"
                    OroenBove = val(GetVar(FileUser, "STATS", "BANCO"))
                    OroenBove = OroenBove + val(Cantidad)

                    Call WriteVar(FileUser, "STATS", "BANCO", CLng(OroenBove)) 'Guardamos en bove
                End If
                UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(Cantidad) 'Quitamos el oro al usuario
            Else
                UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(Cantidad) 'Quitamos el oro al usuario
                UserList(tUser).Stats.Banco = UserList(tUser).Stats.Banco + val(Cantidad) 'Se lo damos al otro.

            End If

            Call WriteChatOverHead(UserIndex, "¡El envio se ha realizado con exito! Gracias por utilizar los servicios de Finanzas Goliath", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("173", UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
            
            Call WriteUpdateGold(UserIndex)
            Call WriteGoliathInit(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "El usuario es inexistente.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleMoveItem(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim SlotViejo As Byte

        Dim SlotNuevo As Byte
        
        SlotViejo = buffer.ReadByte()
        SlotNuevo = buffer.ReadByte()
        
        Dim Objeto    As obj
        
        Dim Equipado  As Boolean

        Dim Equipado2 As Boolean

        Dim Equipado3 As Boolean
        
        If (SlotViejo > .CurrentInventorySlots) Or (SlotNuevo > .CurrentInventorySlots) Then
            Call WriteConsoleMsg(UserIndex, "Slot bloqueado.", FontTypeNames.FONTTYPE_INFOIAO)
        Else
    
            If .Invent.Object(SlotNuevo).ObjIndex <> 0 Then
                Objeto.Amount = .Invent.Object(SlotViejo).Amount
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
                .Invent.Object(SlotViejo).Amount = .Invent.Object(SlotNuevo).Amount
                
                .Invent.Object(SlotNuevo).ObjIndex = Objeto.ObjIndex
                .Invent.Object(SlotNuevo).Amount = Objeto.Amount
                
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
            If .Invent.AnilloEqpSlot = SlotViejo Then
                .Invent.AnilloEqpSlot = SlotNuevo
            ElseIf .Invent.AnilloEqpSlot = SlotNuevo Then
                .Invent.AnilloEqpSlot = SlotViejo

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
                .Invent.Object(SlotNuevo).Amount = .Invent.Object(SlotViejo).Amount
                .Invent.Object(SlotNuevo).Equipped = .Invent.Object(SlotViejo).Equipped
                        
                .Invent.Object(SlotViejo).ObjIndex = 0
                .Invent.Object(SlotViejo).Amount = 0
                .Invent.Object(SlotViejo).Equipped = 0

            End If
            
            Call UpdateUserInv(False, UserIndex, SlotViejo)
            Call UpdateUserInv(False, UserIndex, SlotNuevo)

        End If
    
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBovedaMoveItem(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim SlotViejo As Byte
        Dim SlotNuevo As Byte
        
        SlotViejo = buffer.ReadByte()
        SlotNuevo = buffer.ReadByte()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
        Dim Objeto    As obj
        Dim Equipado  As Boolean
        Dim Equipado2 As Boolean
        Dim Equipado3 As Boolean
        
        Objeto.ObjIndex = UserList(UserIndex).BancoInvent.Object(SlotViejo).ObjIndex
        Objeto.Amount = UserList(UserIndex).BancoInvent.Object(SlotViejo).Amount
        
        UserList(UserIndex).BancoInvent.Object(SlotViejo).ObjIndex = UserList(UserIndex).BancoInvent.Object(SlotNuevo).ObjIndex
        UserList(UserIndex).BancoInvent.Object(SlotViejo).Amount = UserList(UserIndex).BancoInvent.Object(SlotNuevo).Amount
         
        UserList(UserIndex).BancoInvent.Object(SlotNuevo).ObjIndex = Objeto.ObjIndex
        UserList(UserIndex).BancoInvent.Object(SlotNuevo).Amount = Objeto.Amount
    
        'Actualizamos el banco
        Call UpdateBanUserInv(False, UserIndex, SlotViejo)
        Call UpdateBanUserInv(False, UserIndex, SlotNuevo)
        

    End With
    
    Exit Sub
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleQuieroFundarClan(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim refError As String
        
        If UserList(UserIndex).GuildIndex > 0 Then
            refError = "Ya perteneces a un clan, no podés fundar otro."
        Else

            If UserList(UserIndex).Stats.ELV < 45 Or UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) < 80 Then
                refError = "Para fundar un clan debes ser nivel 45, tener 80 en liderazgo y tener en tu inventario las 4 gemas: Gema Azul(1), Gema Naranja(1), Gema Gris(1), Gema Roja(1)."
            Else

                If Not TieneObjetos(407, 1, UserIndex) Then
                    refError = "Para fundar un clan debes ser nivel 45, tener 80 en liderazgo y tener en tu inventario las 4 gemas: Gema Azul(1), Gema Naranja(1), Gema Gris(1), Gema Roja(1)."
                Else

                    If Not TieneObjetos(408, 1, UserIndex) Then
                        refError = "Para fundar un clan debes ser nivel 45, tener 80 en liderazgo y tener en tu inventario las 4 gemas: Gema Azul(1), Gema Naranja(1), Gema Gris(1), Gema Roja(1)."
                    Else

                        If Not TieneObjetos(409, 1, UserIndex) Then
                            refError = "Para fundar un clan debes ser nivel 45, tener 80 en liderazgo y tener en tu inventario las 4 gemas: Gema Azul(1), Gema Naranja(1), Gema Gris(1), Gema Roja(1)."
                        Else

                            If Not TieneObjetos(411, 1, UserIndex) Then
                                refError = "Para fundar un clan debes ser nivel 45, tener 80 en liderazgo y tener en tu inventario las 4 gemas: Gema Azul(1), Gema Naranja(1), Gema Gris(1), Gema Roja(1)."
                            Else

                                If UserList(UserIndex).flags.BattleModo = 1 Then
                                    refError = "No podés fundar un clan ací."
                                Else
                                    refError = "Servidor> íComenzamos a fundar el clan! Ingresa todos los datos solicitados."
                                    Call WriteShowFundarClanForm(UserIndex)
                                    
                                End If

                            End If

                        End If

                    End If

                End If

            End If

        End If
                    
        Call WriteConsoleMsg(UserIndex, refError, FontTypeNames.FONTTYPE_INFOIAO)
    
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleLlamadadeClan(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim refError   As String
        
        Dim clan_nivel As Byte
                        
        If .GuildIndex <> 0 Then
            clan_nivel = modGuilds.NivelDeClan(.GuildIndex)

            If clan_nivel > 1 Then
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Clan> [" & .name & "] solicita apoyo de su clan en " & DarNameMapa(.Pos.Map) & " (" & .Pos.Map & "-" & .Pos.x & "-" & .Pos.Y & "). Puedes ver su ubicaciín en el mapa del mundo.", FontTypeNames.FONTTYPE_GUILD))
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave("43", NO_3D_SOUND, NO_3D_SOUND))
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageUbicacionLlamada(.Pos.Map, .Pos.x, .Pos.Y))
            Else
                Call WriteConsoleMsg(UserIndex, "Servidor> El nivel de tu clan debe ser 2 para utilizar esta opciín.", FontTypeNames.FONTTYPE_INFOIAO)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Servidor> No Perteneces a ningun clan.", FontTypeNames.FONTTYPE_INFOIAO)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Function PrepareMessageNieblandoToggle(ByVal IntensidadMax As Byte) As String
        
        On Error GoTo PrepareMessageNieblandoToggle_Err
        

        '***************************************************
        'Author: Pablo Mercavides
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.NieblaToggle)
104         Call .WriteByte(IntensidadMax)
        
106         PrepareMessageNieblandoToggle = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageNieblandoToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageNieblandoToggle", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageNevarToggle() As String
        
        On Error GoTo PrepareMessageNevarToggle_Err
        

        '***************************************************
        'Author: Pablo Mercavides
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.NieveToggle)
        
104         PrepareMessageNevarToggle = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageNevarToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageNevarToggle", Erl)
        Resume Next
        
End Function

Private Sub HandleGenio(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
        
            Dim i As Byte
        
            For i = 1 To NUMSKILLS
                UserList(UserIndex).Stats.UserSkills(i) = 100
            Next i
        
            Call WriteConsoleMsg(UserIndex, "Tus skills fueron editados.", FontTypeNames.FONTTYPE_INFOIAO)

        End If

        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCasamiento(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String

        Dim tUser    As Integer

        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If .flags.TargetNPC > 0 Then
            If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor Then
                Call WriteConsoleMsg(UserIndex, "Primero haz click sobre un sacerdote.", FontTypeNames.FONTTYPE_INFO)
            Else

                If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
                    Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede casarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Else
            
                    If tUser = UserIndex Then
                        Call WriteConsoleMsg(UserIndex, "No podés casarte contigo mismo.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If tUser <= 0 Then
                            Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
                        Else

                            If UserList(tUser).flags.Candidato = UserIndex Then
                                UserList(tUser).flags.Casado = 1
                                UserList(tUser).flags.Pareja = UserList(UserIndex).name
                                UserList(UserIndex).flags.Casado = 1
                                UserList(UserIndex).flags.Pareja = UserList(tUser).name
                                Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(FXSound.Casamiento_sound, NO_3D_SOUND, NO_3D_SOUND))
                                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El sacerdote de " & DarNameMapa(.Pos.Map) & " celebra el casamiento entre " & UserList(UserIndex).name & " y " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_WARNING))
                                Call WriteChatOverHead(UserIndex, "Los declaro unidos en legal matrimonio íFelicidades!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                                Call WriteChatOverHead(tUser, "Los declaro unidos en legal matrimonio íFelicidades!", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                                
                            Else
                                Call WriteChatOverHead(UserIndex, "La solicitud de casamiento a sido enviada a " & UserName & ".", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                                Call WriteConsoleMsg(tUser, .name & " desea casarse contigo, para permitirlo haz click en el sacerdote y escribe /PROPONER " & .name & ".", FontTypeNames.FONTTYPE_TALK)
                                UserList(UserIndex).flags.Candidato = tUser

                            End If

                        End If

                    End If

                End If

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click sobre el sacerdote.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleEnviarCodigo(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim Codigo As String

        Codigo = buffer.ReadASCIIString()

        Call CheckearCodigo(UserIndex, Codigo)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCrearTorneo(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 26 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
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

        Dim x           As Byte

        Dim Y           As Byte

        Dim nombre      As String

        Dim reglas      As String

        NivelMinimo = buffer.ReadByte
        nivelmaximo = buffer.ReadByte
        cupos = buffer.ReadByte
        costo = buffer.ReadLong
        mago = buffer.ReadByte
        clerico = buffer.ReadByte
        guerrero = buffer.ReadByte
        asesino = buffer.ReadByte
        bardo = buffer.ReadByte
        druido = buffer.ReadByte
        Paladin = buffer.ReadByte
        cazador = buffer.ReadByte
 
        Trabajador = buffer.ReadByte

        Mapa = buffer.ReadInteger
        x = buffer.ReadByte
        Y = buffer.ReadByte
        nombre = buffer.ReadASCIIString
        reglas = buffer.ReadASCIIString
  
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
            Torneo.x = x
            Torneo.Y = Y
            Torneo.nombre = nombre
            Torneo.reglas = reglas

            Call IniciarTorneo

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleComenzarTorneo(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        'NivelMinimo = buffer.ReadByte
  
        If EsGM(UserIndex) Then

            Call ComenzarTorneoOk

        End If
            
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCancelarTorneo(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
  
        If EsGM(UserIndex) Then
            Call ResetearTorneo

        End If
            
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBusquedaTesoro(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim Tipo As Byte

        Tipo = buffer.ReadByte()
        
        Dim Mapa As Byte
  
        If EsGM(UserIndex) Then
    
            If BusquedaTesoroActiva = False And BusquedaRegaloActiva = False Then

                Select Case Tipo

                    Case 0
                        Call PerderTesoro

                    Case 1
                        Call PerderRegalo

                    Case 2
                
                        Dim Pos As WorldPos

                        Mapa = RandomNumber(1, 8)
                
                        Select Case Mapa

                            Case 1
                                Pos.Map = 187

                            Case 2
                                Pos.Map = 188

                            Case 3
                                Pos.Map = 190

                            Case 4
                                Pos.Map = 191

                            Case 5
                                Pos.Map = 234

                            Case 6
                                Pos.Map = 235

                            Case 7
                                Pos.Map = 30

                            Case 8
                                Pos.Map = 98

                            Case 8
                                Pos.Map = 75

                        End Select

                        Pos.Y = 50
                        Pos.x = 50
                        Call SpawnNpc(RandomNumber(592, 593), Pos, True, False, True)
                
                End Select
            
            Else
            
                If BusquedaTesoroActiva = True Then
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Todavia nadie fue capaz de encontar el tesoro, recorda que se encuentra en " & DarNameMapa(TesoroNumMapa) & "(" & TesoroNumMapa & "). íQuien sera el valiente que lo encuentre?", FontTypeNames.FONTTYPE_TALK))
                    Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: " & TesoroNumMapa & "-" & TesoroX & "-" & TesoroY, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Ningun valiente fue capaz de encontrar el item misterioso, recorda que se encuentra en " & DarNameMapa(RegaloNumMapa) & "(" & RegaloNumMapa & "). íTení cuidado!", FontTypeNames.FONTTYPE_TALK))
                    Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: " & RegaloNumMapa & "-" & RegaloX & "-" & RegaloY, FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
            
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleDropItem(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim Item         As Byte

        Dim x            As Byte

        Dim Y            As Byte

        Dim Depositado   As Byte

        Dim DropCantidad As Integer

        Item = buffer.ReadByte()
        x = buffer.ReadByte()
        Y = buffer.ReadByte()
        DropCantidad = buffer.ReadInteger()
        Depositado = 0

        If UserList(UserIndex).flags.Muerto = 1 Then
            ' Call WriteConsoleMsg(UserIndex, "Estas muerto!.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
        Else
    
            If MapData(UserList(UserIndex).Pos.Map, x, Y).Blocked = 1 Or MapData(UserList(UserIndex).Pos.Map, x, Y).TileExit.Map > 0 Or MapData(UserList(UserIndex).Pos.Map, x, Y).NpcIndex > 0 Or HayAgua(UserList(UserIndex).Pos.Map, x, Y) Then
            
                'Call WriteConsoleMsg(UserIndex, "Area invalida para tirar el item.", FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)
            Else
            
                If UserList(UserIndex).flags.BattleModo = 1 Then
                    Call WriteConsoleMsg(UserIndex, "No podes tirar items en este mapa.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If ObjData(.Invent.Object(Item).ObjIndex).Destruye = 1 Then
                        Call WriteConsoleMsg(UserIndex, "Acciín no disponible.", FontTypeNames.FONTTYPE_INFO)
                    Else
                
                        If ObjData(.Invent.Object(Item).ObjIndex).Instransferible = 1 Then
                            Call WriteConsoleMsg(UserIndex, "Acciín no disponible.", FontTypeNames.FONTTYPE_INFO)
                        Else
            
                            If ObjData(.Invent.Object(Item).ObjIndex).Newbie = 1 Then
                                Call WriteConsoleMsg(UserIndex, "No se pueden tirar los objetos Newbies.", FontTypeNames.FONTTYPE_INFO)
                            Else

                                If ObjData(.Invent.Object(Item).ObjIndex).Intirable = 1 Then
                                    Call WriteConsoleMsg(UserIndex, "Este objeto es imposible de tirar.", FontTypeNames.FONTTYPE_INFO)
                                Else
                    
                                    If ObjData(.Invent.Object(Item).ObjIndex).OBJType = eOBJType.otBarcos And UserList(UserIndex).flags.Navegando Then
                                        Call WriteConsoleMsg(UserIndex, "Para tirar la barca deberias estar en tierra firme.", FontTypeNames.FONTTYPE_INFO)
        
                                    Else
                                            
                                        If ObjData(.Invent.Object(Item).ObjIndex).OBJType = eOBJType.otMonturas And UserList(UserIndex).flags.Montado Then
                                            Call WriteConsoleMsg(UserIndex, "Para tirar tu montura deberias descender de ella.", FontTypeNames.FONTTYPE_INFO)
        
                                        Else
                
                                            Call DropObj(UserIndex, Item, DropCantidad, UserList(UserIndex).Pos.Map, x, Y)

                                        End If

                                        'End If
                                    End If

                                End If

                            End If

                        End If

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleFlagTrabajar(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        UserList(UserIndex).Counters.Trabajando = 0
        UserList(UserIndex).flags.UsandoMacro = False
        UserList(UserIndex).flags.TargetObj = 0 ' Sacamos el targer del objeto
        UserList(UserIndex).flags.UltimoMensaje = 0
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleEscribiendo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        If .flags.Escribiendo = False Then
            .flags.Escribiendo = True
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetEscribiendo(.Char.CharIndex, True))
        Else
            .flags.Escribiendo = False
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetEscribiendo(.Char.CharIndex, False))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleRequestFamiliar(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        'Remove packet ID
        
        On Error GoTo HandleRequestFamiliar_Err
        
100     Call UserList(UserIndex).incomingData.ReadInteger

102     Call WriteFamiliar(UserIndex)

        
        Exit Sub

HandleRequestFamiliar_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestFamiliar", Erl)
        Resume Next
        
End Sub

Public Sub WriteFamiliar(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Familiar)
        Call .WriteByte(UserList(UserIndex).Familiar.Existe)
        Call .WriteByte(UserList(UserIndex).Familiar.Muerto)
        Call .WriteASCIIString(UserList(UserIndex).Familiar.nombre)
        Call .WriteLong(UserList(UserIndex).Familiar.Exp)
        Call .WriteLong(UserList(UserIndex).Familiar.ELU)
        Call .WriteByte(UserList(UserIndex).Familiar.nivel)
        Call .WriteInteger(UserList(UserIndex).Familiar.MinHp)
        Call .WriteInteger(UserList(UserIndex).Familiar.MaxHp)
        Call .WriteInteger(UserList(UserIndex).Familiar.MinHIT)
        Call .WriteInteger(UserList(UserIndex).Familiar.MaxHit)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Function PrepareMessageBarFx(ByVal CharIndex As Integer, ByVal BarTime As Integer, ByVal BarAccion As Byte) As String
        
        On Error GoTo PrepareMessageBarFx_Err
        

        '***************************************************
        'Author: Pablo Mercavides
        'Last Modification: 20/10/2014
        'Envia el Efecto de Barra
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.BarFx)
104         Call .WriteInteger(CharIndex)
106         Call .WriteInteger(BarTime)
108         Call .WriteByte(BarAccion)
        
110         PrepareMessageBarFx = .ReadASCIIStringFixed(.length)

        End With

        
        Exit Function

PrepareMessageBarFx_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageBarFx", Erl)
        Resume Next
        
End Function

Private Sub HandleCompletarAccion(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim accion As Byte

        accion = buffer.ReadByte()
        
        If .accion.AccionPendiente = True Then
            If .accion.TipoAccion = accion Then
                Call CompletarAccionFin(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "Servidor> La acciín que solicitas no se corresponde.", FontTypeNames.FONTTYPE_SERVER)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Servidor> Tu no tenias ninguna acciín pendiente. ", FontTypeNames.FONTTYPE_SERVER)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleReclamarRecompensa(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Dim Index  As Byte

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Index = buffer.ReadByte()
        
        Call EntregarRecompensas(UserIndex, Index)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleTraerRecompensas(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Call EnviarRecompensaStat(UserIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub WriteRecompensas(ByVal UserIndex As Integer)
        
        On Error GoTo WriteRecompensas_Err
        

        '***************************************************
        'Envia las recompensas al cliente!
        'Por Ladder
        '22/04/2015
        'Flor te amo!
        '***************************************************

100     With UserList(UserIndex).outgoingData
    
            Dim a, b, c As Byte
 
102         b = UserList(UserIndex).UserLogros + 1
104         a = UserList(UserIndex).NPcLogros + 1
106         c = UserList(UserIndex).LevelLogros + 1
        
108         Call .WriteByte(ServerPacketID.Logros)
            'Logros NPC
110         Call .WriteASCIIString(NPcLogros(a).nombre)
112         Call .WriteASCIIString(NPcLogros(a).Desc)
114         Call .WriteInteger(NPcLogros(a).cant)
116         Call .WriteByte(NPcLogros(a).TipoRecompensa)
        
118         If NPcLogros(a).TipoRecompensa = 1 Then
120             Call .WriteASCIIString(NPcLogros(a).ObjRecompensa)

            End If

122         If NPcLogros(a).TipoRecompensa = 2 Then
124             Call .WriteLong(NPcLogros(a).OroRecompensa)

            End If

126         If NPcLogros(a).TipoRecompensa = 3 Then
128             Call .WriteLong(NPcLogros(a).ExpRecompensa)

            End If
        
130         If NPcLogros(a).TipoRecompensa = 4 Then
132             Call .WriteByte(NPcLogros(a).HechizoRecompensa)

            End If
        
134         Call .WriteInteger(UserList(UserIndex).Stats.NPCsMuertos)
        
136         If UserList(UserIndex).Stats.NPCsMuertos >= NPcLogros(a).cant Then
138             Call .WriteBoolean(True)
            Else
140             Call .WriteBoolean(False)

            End If
        
            'Logros User
142         Call .WriteASCIIString(UserLogros(b).nombre)
144         Call .WriteASCIIString(UserLogros(b).Desc)
146         Call .WriteInteger(UserLogros(b).cant)
148         Call .WriteInteger(UserLogros(b).TipoRecompensa)
150         Call .WriteInteger(UserList(UserIndex).Stats.UsuariosMatados)

152         If UserLogros(a).TipoRecompensa = 1 Then
154             Call .WriteASCIIString(UserLogros(b).ObjRecompensa)

            End If
        
156         If UserLogros(a).TipoRecompensa = 2 Then
158             Call .WriteLong(UserLogros(b).OroRecompensa)

            End If

160         If UserLogros(a).TipoRecompensa = 3 Then
162             Call .WriteLong(UserLogros(b).ExpRecompensa)

            End If
        
164         If UserLogros(a).TipoRecompensa = 4 Then
166             Call .WriteByte(UserLogros(b).HechizoRecompensa)

            End If

168         If UserList(UserIndex).Stats.UsuariosMatados >= UserLogros(b).cant Then
170             Call .WriteBoolean(True)
            Else
172             Call .WriteBoolean(False)

            End If

            'Nivel User
174         Call .WriteASCIIString(LevelLogros(c).nombre)
176         Call .WriteASCIIString(LevelLogros(c).Desc)
178         Call .WriteInteger(LevelLogros(c).cant)
180         Call .WriteInteger(LevelLogros(c).TipoRecompensa)
182         Call .WriteByte(UserList(UserIndex).Stats.ELV)

184         If LevelLogros(c).TipoRecompensa = 1 Then
186             Call .WriteASCIIString(LevelLogros(c).ObjRecompensa)

            End If
        
188         If LevelLogros(c).TipoRecompensa = 2 Then
190             Call .WriteLong(LevelLogros(c).OroRecompensa)

            End If

192         If LevelLogros(c).TipoRecompensa = 3 Then
194             Call .WriteLong(LevelLogros(c).ExpRecompensa)

            End If
        
196         If LevelLogros(c).TipoRecompensa = 4 Then
198             Call .WriteByte(LevelLogros(c).HechizoRecompensa)

            End If

200         If UserList(UserIndex).Stats.ELV >= LevelLogros(c).cant Then
202             Call .WriteBoolean(True)
            Else
204             Call .WriteBoolean(False)

            End If

        End With

        Exit Sub

        
        Exit Sub

WriteRecompensas_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.WriteRecompensas", Erl)
        Resume Next
        
End Sub

Private Sub HandleDecimeLaHora(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Call WriteHora(UserIndex)
        
        Dim msg As String

        msg = "Servidor> Son las " & GetTimeFormated(HoraFanstasia) & "... "
        
        If HoraFanstasia > 0 And HoraFanstasia < 300 Then
            msg = msg & "ídespierto a estas horas? íno olvides visitar El Mesín Hostigado!"

        End If

        If HoraFanstasia > 300 And HoraFanstasia < 720 Then
            msg = msg & "el sol se asoma lentamente en el horizonte."

        End If
        
        If HoraFanstasia > 720 And HoraFanstasia < 1080 Then
            ' msg = msg & "lentamente el dia termina"
            msg = msg & "íno pierdas el tiempo!"

        End If
                
        If HoraFanstasia > 1080 And HoraFanstasia < 1260 Then
            msg = msg & "lentamente el dia termina."

        End If
        
        If HoraFanstasia > 1260 And HoraFanstasia < 1440 Then
            msg = msg & "ídespierto a estas horas? íno olvides visitar El Mesín Hostigado!"

        End If

        '  Call EnviarRecompensaStat(UserIndex)
        Call WriteConsoleMsg(UserIndex, msg, FontTypeNames.FONTTYPE_SUBASTA)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCorreo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Call WriteListaCorreo(UserIndex, False)
        '    Call EnviarRecompensaStat(UserIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleSendCorreo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 7 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Dim Nick               As String

        Dim msg                As String

        Dim ItemCount          As Byte

        Dim cant               As Integer

        Dim IndexReceptor      As Integer

        Dim Itemlista(1 To 10) As obj

        Nick = buffer.ReadASCIIString()
        msg = buffer.ReadASCIIString()
        ItemCount = buffer.ReadByte()
        
        Dim ObjIndex   As Integer

        Dim FinalCount As Byte

        Dim HuboError  As Boolean
                
        If ItemCount > 0 Then 'Si el correo tiene item

            Dim i As Byte

            For i = 1 To ItemCount
                Itemlista(i).ObjIndex = buffer.ReadByte
                Itemlista(i).Amount = buffer.ReadInteger
            Next i

        Else 'Si es solo texto
            'IndexReceptor = NameIndex(Nick)
            FinalCount = 0
            AddCorreo UserIndex, Nick, msg, 0, FinalCount

        End If
        
        Dim ObjArray As String
        
        If UserList(UserIndex).flags.BattleModo = 0 Then

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
                                
                                    Call QuitarUserInvItem(UserIndex, Itemlista(i).ObjIndex, Itemlista(i).Amount)
                                    Call UpdateUserInv(False, UserIndex, Itemlista(i).ObjIndex)
                                    FinalCount = FinalCount + 1
                                    ObjArray = ObjArray & ObjIndex & "-" & Itemlista(i).Amount & "@"

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
            Call WriteConsoleMsg(UserIndex, "No podes usar el correo desde el battle.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
    Exit Sub
Errhandler:
    LogError "Error HandleSendCorreo"

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleRetirarItemCorreo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Dim MsgIndex As Integer

        MsgIndex = buffer.ReadInteger()
        
        Call ExtractItemCorreo(UserIndex, MsgIndex)
        
        Call .incomingData.CopyBuffer(buffer)

    End With
    
    Exit Sub
    
Errhandler:
    LogError "Error handleRetirarItem"

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBorrarCorreo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Dim MsgIndex As Integer

        MsgIndex = buffer.ReadInteger()
        
        Call BorrarCorreoMail(UserIndex, MsgIndex)
        
        Call .incomingData.CopyBuffer(buffer)

    End With
    
    Exit Sub
    
Errhandler:

    LogError "Error BorrarCorreo"

    Dim Error As Long

    Error = Err.Number
    
    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleInvitarGrupo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteWorkRequestTarget(UserIndex, eSkill.Grupo)

        End If
            
        Call .incomingData.CopyBuffer(buffer)

    End With

    Exit Sub
    
Errhandler:
    LogError "Error InvitarGrupo"

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleMarcaDeClan(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
       
        Call WriteWorkRequestTarget(UserIndex, eSkill.MarcaDeClan)
                
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleMarcaDeGM(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
          
        Call WriteWorkRequestTarget(UserIndex, eSkill.MarcaDeGM)
                
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub WritePreguntaBox(ByVal UserIndex As Integer, ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowPregunta)
        Call .WriteASCIIString(Message)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleResponderPregunta(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Dim respuesta As Boolean

        Dim DeDonde   As String

        respuesta = buffer.ReadBoolean()
        
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
                                
                                Dim Index As Byte
                                
                                Log = "Repuesta Afirmativa 1-5 "
                                
                                For Index = 2 To UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.CantidadMiembros - 1
                                    Call WriteLocaleMsg(UserList(UserList(UserIndex).Grupo.PropuestaDe).Grupo.Miembros(Index), "40", FontTypeNames.FONTTYPE_INFOIAO, UserList(UserIndex).name)
                                
                                Next Index
                                
                                Log = "Repuesta Afirmativa 1-6 "
                                'Call WriteConsoleMsg(UserList(UserIndex).Grupo.PropuestaDe, "í" & UserList(UserIndex).name & " a sido aíadido al grupo!", FontTypeNames.FONTTYPE_INFOIAO)
                                Call WriteLocaleMsg(UserList(UserIndex).Grupo.PropuestaDe, "40", FontTypeNames.FONTTYPE_INFOIAO, UserList(UserIndex).name)
                                
                                Call WriteConsoleMsg(UserIndex, "¡Has sido aíadido al grupo!", FontTypeNames.FONTTYPE_INFOIAO)
                                
                                Log = "Repuesta Afirmativa 1-7 "
                                
                                Call RefreshCharStatus(UserList(UserIndex).Grupo.PropuestaDe)
                                Call RefreshCharStatus(UserIndex)
                                 
                                Log = "Repuesta Afirmativa 1-8"

                            End If

                        End If

                    Else
                    
                        Call WriteConsoleMsg(UserIndex, "Servidor> Solicitud de grupo invalida, reintente...", FontTypeNames.FONTTYPE_SERVER)
                    
                    End If

                    'unirlo
                Case 2
                    Log = "Repuesta Afirmativa 2"
                    UserList(UserIndex).Faccion.Status = 1
                    Call WriteConsoleMsg(UserIndex, "íAhora sos un ciudadano!", FontTypeNames.FONTTYPE_INFOIAO)
                    Call RefreshCharStatus(UserIndex)
                    
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
                            
                        Case eCiudad.CHillidan
                            DeDonde = " Hillidan"
                            
                        Case Else
                            DeDonde = "Ullathorpe"

                    End Select
                    
                    If UserList(UserIndex).flags.TargetNPC <> 0 Then
                    
                        Call WriteChatOverHead(UserIndex, "íGracias " & UserList(UserIndex).name & "! Ahora perteneces a la ciudad de " & DeDonde & ".", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
                    Else
                        Call WriteConsoleMsg(UserIndex, "íGracias " & UserList(UserIndex).name & "! Ahora perteneces a la ciudad de " & DeDonde & ".", FontTypeNames.FONTTYPE_INFOIAO)

                    End If
                    
                Case 4
                    Log = "Repuesta Afirmativa 4"
                
                    If UserList(UserIndex).flags.TargetUser <> 0 Then
                
                        If UserList(UserList(UserIndex).flags.TargetUser).flags.BattleModo = 1 Then
                            Call WriteConsoleMsg(UserIndex, "No podes usar el sistema de comercio cuando el otro personaje esta en el battle.", FontTypeNames.FONTTYPE_EXP)
                        
                        Else
                    
                            UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                            UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).name
                            UserList(UserIndex).ComUsu.cant = 0
                            UserList(UserIndex).ComUsu.Objeto = 0
                            UserList(UserIndex).ComUsu.Acepto = False
                    
                            'Rutina para comerciar con otro usuario
                            Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "Servidor> Solicitud de comercio invalida, reintente...", FontTypeNames.FONTTYPE_SERVER)
                
                    End If
                
                Case 5
                    Log = "Repuesta Afirmativa 5"
                
                    If UCase$(MapInfo(UserList(UserIndex).Pos.Map).restrict_mode) = "NEWBIE" Then
                        Call WarpToLegalPos(UserIndex, 140, 53, 58)
                    
                        If UserList(UserIndex).donador.activo = 0 Then ' Donador no espera tiempo
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, 400, False))
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 400, Accion_Barra.Resucitar))
                        Else
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Resucitar, 10, False))
                            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 10, Accion_Barra.Resucitar))

                        End If
                    
                        UserList(UserIndex).accion.AccionPendiente = True
                        UserList(UserIndex).accion.Particula = ParticulasIndex.Resucitar
                        UserList(UserIndex).accion.TipoAccion = Accion_Barra.Resucitar
    
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("104", UserList(UserIndex).Pos.x, UserList(UserIndex).Pos.Y))
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
                    UserList(UserIndex).Faccion.Status = 0
                    Call WriteConsoleMsg(UserIndex, "¡Continuas siendo neutral!", FontTypeNames.FONTTYPE_INFOIAO)
                    Call RefreshCharStatus(UserIndex)

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
                            
                        Case eCiudad.CHillidan
                            DeDonde = " Hillidan"
                            
                        Case Else
                            DeDonde = "Ullathorpe"

                    End Select
                    
                    If UserList(UserIndex).flags.TargetNPC <> 0 Then
                        Call WriteChatOverHead(UserIndex, "¡No hay problema " & UserList(UserIndex).name & "! Sos bienvenido en " & DeDonde & " cuando gustes.", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)

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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
    Exit Sub
    
Errhandler:

    LogError "Error ResponderPregunta " & Log

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleRequestGrupo(ByVal UserIndex As Integer)

    On Error GoTo hErr

    'Author: Pablo Mercavides
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadInteger

    Call WriteDatosGrupo(UserIndex)
    
    Exit Sub
    
hErr:
    LogError "Error HandleRequestGrupo"

End Sub

Public Sub WriteDatosGrupo(ByVal UserIndex As Integer)

    Dim i As Byte

    'Author: Pablo Mercavides
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.DatosGrupo)
        Call .WriteBoolean(UserList(UserIndex).Grupo.EnGrupo)
        
        If UserList(UserIndex).Grupo.EnGrupo = True Then
            Call .WriteByte(UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros)
            'Call .WriteByte(UserList(UserList(UserIndex).Grupo.Lider).name)
   
            If UserList(UserIndex).Grupo.Lider = UserIndex Then
             
                For i = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros

                    If i = 1 Then
                        Call .WriteASCIIString(UserList(UserList(UserIndex).Grupo.Miembros(i)).name & "(Lider)")
                    Else
                        Call .WriteASCIIString(UserList(UserList(UserIndex).Grupo.Miembros(i)).name)

                    End If

                Next i

            Else
          
                For i = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
                
                    If i = 1 Then
                        Call .WriteASCIIString(UserList(UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)).name & "(Lider)")
                    Else
                        Call .WriteASCIIString(UserList(UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)).name)

                    End If

                Next i

            End If

        End If
   
    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleAbandonarGrupo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
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
                
        Call .incomingData.CopyBuffer(buffer)

    End With
    
    Exit Sub
Errhandler:

    LogError "Error HandleAbandonarGrupo"

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub WriteUbicacion(ByVal UserIndex As Integer, ByVal Miembro As Byte, ByVal GPS As Integer)

    Dim i   As Byte

    Dim x   As Byte

    Dim Y   As Byte

    Dim Map As Integer

    'Author: Pablo Mercavides
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.Ubicacion)
        Call .WriteByte(Miembro)

        If GPS > 0 Then
        
            Call .WriteByte(UserList(GPS).Pos.x)
            Call .WriteByte(UserList(GPS).Pos.Y)
            Call .WriteInteger(UserList(GPS).Pos.Map)
        Else
            Call .WriteByte(0)
            Call .WriteByte(0)
            Call .WriteInteger(0)

        End If
   
    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleHecharDeGrupo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim indice As Byte

        indice = buffer.ReadByte()
        
        Call HecharMiembro(UserIndex, indice)
                
        Call .incomingData.CopyBuffer(buffer)

    End With
    
    Exit Sub
Errhandler:
    LogError "Error HandleHecharDeGrupo"

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleMacroPos(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        UserList(UserIndex).ChatCombate = buffer.ReadByte()
        UserList(UserIndex).ChatGlobal = buffer.ReadByte()
                
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub WriteCorreoPicOn(ByVal UserIndex As Integer)

    '***************************************************
    '***************************************************
    On Error GoTo Errhandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CorreoPicOn)
    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleSubastaInfo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        If Subasta.HaySubastaActiva Then

            Call WriteConsoleMsg(UserIndex, "Subastador: " & Subasta.Subastador, FontTypeNames.FONTTYPE_SUBASTA)
            Call WriteConsoleMsg(UserIndex, "Objeto: " & ObjData(Subasta.ObjSubastado).name & " (" & Subasta.ObjSubastadoCantidad & ")", FontTypeNames.FONTTYPE_SUBASTA)

            If Subasta.HuboOferta Then
                Call WriteConsoleMsg(UserIndex, "Mejor oferta: " & Subasta.MejorOferta & " monedas de oro por " & Subasta.Comprador & ".", FontTypeNames.FONTTYPE_SUBASTA)
                Call WriteConsoleMsg(UserIndex, "Podes realizar una oferta escribiendo /OFERTAR " & Subasta.MejorOferta + 100, FontTypeNames.FONTTYPE_SUBASTA)
            Else
                Call WriteConsoleMsg(UserIndex, "Oferta inicial: " & Subasta.OfertaInicial & " monedas de oro.", FontTypeNames.FONTTYPE_SUBASTA)
                Call WriteConsoleMsg(UserIndex, "Podes realizar una oferta escribiendo /OFERTAR " & Subasta.OfertaInicial + 100, FontTypeNames.FONTTYPE_SUBASTA)

            End If

            Call WriteConsoleMsg(UserIndex, "Tiempo Restante de subasta:  " & SumarTiempo(Subasta.TiempoRestanteSubasta), FontTypeNames.FONTTYPE_SUBASTA)
            
        Else
            Call WriteConsoleMsg(UserIndex, "No hay ninguna subasta activa en este momento.", FontTypeNames.FONTTYPE_SUBASTA)

        End If
                
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleScrollInfo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
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
                
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCancelarExit(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleCancelarExit_Err
        

100     With UserList(UserIndex)
            'Remove Packet ID
102         Call .incomingData.ReadInteger
    
            'If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

104         Call CancelExit(UserIndex)

        End With
        
        
        Exit Sub

HandleCancelarExit_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCancelarExit", Erl)
        Resume Next
        
End Sub

Private Sub HandleBanCuenta(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 6 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String

        Dim Reason   As String
        
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanAccount(UserIndex, UserName, Reason)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleUnBanCuenta(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call UnBanAccount(UserIndex, UserName)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBanSerial(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String
         
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanSerialOK(UserIndex, UserName)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleUnBanSerial(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String
         
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call UnBanSerialOK(UserIndex, UserName)
            
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCerrarCliente(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 4 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String

        Dim tUser    As Integer
         
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then

            tUser = NameIndex(UserName)
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " cerro el cliente de " & UserName & ".", FontTypeNames.FONTTYPE_INFO))
                    
                Call WriteCerrarleCliente(tUser)
                'Call CloseSocket(tUser)
                Call LogGM(.name, "Cerro el cliene de:" & UserName)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleEventoInfo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

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
 
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCrearEvento(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo Mercavides
    '***************************************************
    If UserList(UserIndex).incomingData.length < 5 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim Tipo           As Byte

        Dim duracion       As Byte

        Dim multiplicacion As Byte
        
        Tipo = buffer.ReadByte()
        duracion = buffer.ReadByte()
        multiplicacion = buffer.ReadByte()
        
        '/
        If .flags.Privilegios >= PlayerType.Admin Then
            If EventoActivo = False Then
                If LenB(Tipo) = 0 Or LenB(duracion) = 0 Or LenB(multiplicacion) = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Utilice /CREAREVENTO TIPO@DURACION@MULTIPLICACION.", FontTypeNames.FONTTYPE_New_Eventos)
                Else
                
                    Call ForzarEvento(Tipo, duracion, multiplicacion, UserList(UserIndex).name)
                  
                End If

            Else
                Call WriteConsoleMsg(UserIndex, "Ya hay un evento en curso. Finalicelo con /FINEVENTO primero.", FontTypeNames.FONTTYPE_New_Eventos)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBanTemporal(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.length < 7 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String

        Dim Reason   As String

        Dim dias     As Byte
        
        UserName = buffer.ReadASCIIString()
        Reason = buffer.ReadASCIIString()
        dias = buffer.ReadByte()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call Admin.BanTemporal(UserName, dias, Reason, UserList(UserIndex).name)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleTraerShop(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        If UserList(UserIndex).flags.BattleModo = 1 Then
            Call WriteConsoleMsg(UserIndex, "No disponible aquí.", FontTypeNames.FONTTYPE_INFOIAO)
        Else
            Call WriteShop(UserIndex)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleTraerRanking(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Call WriteRanking(UserIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandlePareja(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 2 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim parejaindex As Integer

        If Not UserList(UserIndex).flags.BattleModo Then
                
            If UserList(UserIndex).donador.activo = 1 Then
                If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
                    If UserList(UserIndex).flags.Casado = 1 Then
                        parejaindex = NameIndex(UserList(UserIndex).flags.Pareja)
                        
                        If parejaindex > 0 Then
                            If Not UserList(parejaindex).flags.BattleModo Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, 600, False))
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 600, Accion_Barra.GoToPareja))
                                UserList(UserIndex).accion.AccionPendiente = True
                                UserList(UserIndex).accion.Particula = ParticulasIndex.Runa
                                UserList(UserIndex).accion.TipoAccion = Accion_Barra.GoToPareja
                            Else
                                Call WriteConsoleMsg(UserIndex, "Tu pareja esta en modo battle. No podés teletransportarte hacia ella.", FontTypeNames.FONTTYPE_INFOIAO)

                            End If
                                
                        Else
                            Call WriteConsoleMsg(UserIndex, "Tu pareja no esta online.", FontTypeNames.FONTTYPE_INFOIAO)

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "No estas casado con nadie.", FontTypeNames.FONTTYPE_INFOIAO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Solo disponible en zona segura.", FontTypeNames.FONTTYPE_INFOIAO)

                End If
                
            Else
                Call WriteConsoleMsg(UserIndex, "Opcion disponible unicamente para usuarios donadores.", FontTypeNames.FONTTYPE_INFOIAO)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "No podés usar esta opciín en el battle.", FontTypeNames.FONTTYPE_INFOIAO)
        
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub WriteShop(ByVal UserIndex As Integer)

    '***************************************************
    On Error GoTo Errhandler

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ObjDonador()))
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.DonadorObj)
        
        For i = 1 To UBound(ObjDonador())
            Count = Count + 1
            validIndexes(Count) = i
        Next i
        
        ' Write the number of objects in the list
        Call .WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Call .WriteInteger(ObjDonador(validIndexes(i)).ObjIndex)
            Call .WriteInteger(ObjDonador(validIndexes(i)).Valor)
        Next i
        
        Call .WriteLong(CreditosDonadorCheck(UserList(UserIndex).Cuenta))
        Call .WriteInteger(DiasDonadorCheck(UserList(UserIndex).Cuenta))

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteRanking(ByVal UserIndex As Integer)

    '***************************************************
    On Error GoTo Errhandler

    Dim i As Byte
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Ranking)

        For i = 1 To 10
            Call .WriteASCIIString(Rankings(1).user(i).Nick)
            Call .WriteInteger(Rankings(1).user(i).Value)
        Next i
        
    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleComprarItem(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 3 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Dim ItemIndex    As Byte
        
        Dim ObjComprado  As obj

        Dim LogeoDonador As String

        ItemIndex = buffer.ReadByte()
        
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
                ObjComprado.Amount = ObjDonador(ItemIndex).Cantidad
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
    
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCompletarViaje(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides

    If UserList(UserIndex).incomingData.length < 7 Then
        Err.raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo Errhandler

    With UserList(UserIndex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Dim Destino As Byte

        Dim costo   As Long

        Destino = buffer.ReadByte()
        costo = buffer.ReadLong()
        
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
                        
                Case eCiudad.CHillidan
                    DeDonde = CityHillidan
                        
                Case Else
                    DeDonde = CityUllathorpe

            End Select
        
            If DeDonde.NecesitaNave > 0 Then
                If UserList(UserIndex).Stats.UserSkills(eSkill.Navegacion) < 80 Then
                    Rem Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_WARNING)
                Else

                    If UserList(UserIndex).flags.TargetNPC <> 0 Then
                        If Npclist(UserList(UserIndex).flags.TargetNPC).SoundClose <> 0 Then
                            Call WritePlayWave(UserIndex, Npclist(UserList(UserIndex).flags.TargetNPC).SoundClose, NO_3D_SOUND, NO_3D_SOUND)

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

                Dim x   As Byte

                Dim Y   As Byte
            
                Map = DeDonde.MapaViaje
                x = DeDonde.ViajeX
                Y = DeDonde.ViajeY

                If UserList(UserIndex).flags.TargetNPC <> 0 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).SoundClose <> 0 Then
                        Call WritePlayWave(UserIndex, Npclist(UserList(UserIndex).flags.TargetNPC).SoundClose, NO_3D_SOUND, NO_3D_SOUND)

                    End If

                End If
                
                Call WarpUserChar(UserIndex, Map, x, Y, True)
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
    
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
Errhandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Function PrepareMessageCreateRenderValue(ByVal x As Byte, ByVal Y As Byte, ByVal rValue As Double, ByVal rType As Byte)
        '***************************************************
        'Author: maTih.-
        'Last Modification: 09/06/2012 - ^[GS]^
        '***************************************************
        
        On Error GoTo PrepareMessageCreateRenderValue_Err
        

        ' @ Envia el paquete para crear un valor en el render
     
100     With auxiliarBuffer
102         .WriteByte ServerPacketID.CreateRenderText
104         .WriteByte x
106         .WriteByte Y
108         .WriteDouble rValue
110         .WriteByte rType
         
112         PrepareMessageCreateRenderValue = .ReadASCIIStringFixed(.length)
         
        End With
     
        
        Exit Function

PrepareMessageCreateRenderValue_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCreateRenderValue", Erl)
        Resume Next
        
End Function

Public Sub WriteActShop(ByVal UserIndex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.ACTSHOP)
        Call .WriteLong(CreditosDonadorCheck(UserList(UserIndex).Cuenta))
        
        Call .WriteInteger(DiasDonadorCheck(UserList(UserIndex).Cuenta))

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteViajarForm(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.ViajarForm)
        
        Dim destinos As Byte

        Dim i        As Byte

        destinos = Npclist(NpcIndex).NumDestinos
        
        Call .WriteByte(destinos)
        
        For i = 1 To destinos
            Call .WriteASCIIString(Npclist(NpcIndex).Dest(i))
        Next i
        
        Call .WriteByte(Npclist(NpcIndex).Interface)

    End With

    Exit Sub

Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub HandleQuest(ByVal UserIndex As Integer)
        
        On Error GoTo HandleQuest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete Quest.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim NpcIndex As Integer

        Dim tmpByte  As Byte
 
        'Leemos el paquete
    
100     Call UserList(UserIndex).incomingData.ReadInteger
 
102     NpcIndex = UserList(UserIndex).flags.TargetNPC
    
104     If NpcIndex = 0 Then Exit Sub
    
        'Esta el personaje en la distancia correcta?
106     If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > 5 Then
108         Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        'El NPC hace quests?
110     If Npclist(NpcIndex).QuestNumber = 0 Then
112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ninguna mision para ti.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
            Exit Sub

        End If
    
        'El personaje ya hizo la quest?
114     If UserDoneQuest(UserIndex, Npclist(NpcIndex).QuestNumber) Then
116         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Ya has hecho una mision para mi.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
            Exit Sub

        End If
 
        'El personaje tiene suficiente nivel?
118     If UserList(UserIndex).Stats.ELV < QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel Then
120         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Debes ser por lo menos nivel " & QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel & " para emprender esta mision.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
            Exit Sub

        End If
    
        'A esta altura ya analizo todas las restricciones y esta preparado para el handle propiamente dicho
 
122     tmpByte = TieneQuest(UserIndex, Npclist(NpcIndex).QuestNumber)
    
124     If tmpByte Then
            'El usuario esta haciendo la quest, entonces va a hablar con el NPC para recibir la recompensa.
126         Call FinishQuest(UserIndex, Npclist(NpcIndex).QuestNumber, tmpByte)
        Else
            'El usuario no esta haciendo la quest, entonces primero recibe un informe con los detalles de la mision.
128         tmpByte = FreeQuestSlot(UserIndex)
        
            'El personaje tiene algun slot de quest para la nueva quest?
130         If tmpByte = 0 Then
132             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                Exit Sub

            End If
        
            'Enviamos los detalles de la quest
134         Call WriteQuestDetails(UserIndex, Npclist(NpcIndex).QuestNumber)

        End If

        
        Exit Sub

HandleQuest_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleQuest", Erl)
        Resume Next
        
End Sub

Public Sub HandleQuestAccept(ByVal UserIndex As Integer)
        
        On Error GoTo HandleQuestAccept_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el evento de aceptar una quest.
        'Last modified: 31/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim NpcIndex  As Integer

        Dim QuestSlot As Byte
 
100     Call UserList(UserIndex).incomingData.ReadInteger
 
102     NpcIndex = UserList(UserIndex).flags.TargetNPC
    
104     If NpcIndex = 0 Then Exit Sub
    
        'Esta el personaje en la distancia correcta?
106     If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > 5 Then
108         Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
110     QuestSlot = FreeQuestSlot(UserIndex)
    
        'Agregamos la quest.
112     With UserList(UserIndex).QuestStats.Quests(QuestSlot)
114         .QuestIndex = Npclist(NpcIndex).QuestNumber
        
116         If QuestList(.QuestIndex).RequiredNPCs Then ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
118         Call WriteConsoleMsg(UserIndex, "Has aceptado la mision " & Chr(34) & QuestList(.QuestIndex).nombre & Chr(34) & ".", FontTypeNames.FONTTYPE_INFOIAO)
        
        End With

        
        Exit Sub

HandleQuestAccept_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleQuestAccept", Erl)
        Resume Next
        
End Sub

Public Sub HandleQuestDetailsRequest(ByVal UserIndex As Integer)
        
        On Error GoTo HandleQuestDetailsRequest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestInfoRequest.
        'Last modified: 30/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim QuestSlot As Byte
 
        'Leemos el paquete
100     Call UserList(UserIndex).incomingData.ReadInteger
    
102     QuestSlot = UserList(UserIndex).incomingData.ReadByte
    
104     Call WriteQuestDetails(UserIndex, UserList(UserIndex).QuestStats.Quests(QuestSlot).QuestIndex, QuestSlot)

        
        Exit Sub

HandleQuestDetailsRequest_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleQuestDetailsRequest", Erl)
        Resume Next
        
End Sub
 
Public Sub HandleQuestAbandon(ByVal UserIndex As Integer)
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestAbandon.
        'Last modified: 31/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Leemos el paquete.
        
        On Error GoTo HandleQuestAbandon_Err
        
100     Call UserList(UserIndex).incomingData.ReadInteger
    
        'Borramos la quest.
102     Call CleanQuestSlot(UserIndex, UserList(UserIndex).incomingData.ReadByte)
    
        'Ordenamos la lista de quests del usuario.
104     Call ArrangeUserQuests(UserIndex)
    
        'Enviamos la lista de quests actualizada.
106     Call WriteQuestListSend(UserIndex)

        
        Exit Sub

HandleQuestAbandon_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleQuestAbandon", Erl)
        Resume Next
        
End Sub

Public Sub HandleQuestListRequest(ByVal UserIndex As Integer)
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestListRequest.
        'Last modified: 30/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        
        On Error GoTo HandleQuestListRequest_Err
        
 
        'Leemos el paquete
100     Call UserList(UserIndex).incomingData.ReadInteger
    
102     If UserList(UserIndex).flags.BattleModo = 0 Then
104         Call WriteQuestListSend(UserIndex)
        Else
106         Call WriteConsoleMsg(UserIndex, "No disponible aquí.", FontTypeNames.FONTTYPE_INFOIAO)

        End If

        
        Exit Sub

HandleQuestListRequest_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleQuestListRequest", Erl)
        Resume Next
        
End Sub

Public Sub WriteQuestDetails(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, Optional QuestSlot As Byte = 0)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Envía el paquete QuestDetails y la informaciín correspondiente.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
 
    On Error GoTo Errhandler

    With UserList(UserIndex).outgoingData
        'ID del paquete
        Call .WriteByte(ServerPacketID.QuestDetails)
        
        'Se usa la variable QuestSlot para saber si enviamos la info de una quest ya empezada o la info de una quest que no se aceptí todavía (1 para el primer caso y 0 para el segundo)
        Call .WriteByte(IIf(QuestSlot, 1, 0))
        
        'Enviamos nombre, descripciín y nivel requerido de la quest
        'Call .WriteASCIIString(QuestList(QuestIndex).Nombre)
        'Call .WriteASCIIString(QuestList(QuestIndex).Desc)
        Call .WriteInteger(QuestIndex)
        Call .WriteByte(QuestList(QuestIndex).RequiredLevel)
        
        'Enviamos la cantidad de npcs requeridos
        Call .WriteByte(QuestList(QuestIndex).RequiredNPCs)

        If QuestList(QuestIndex).RequiredNPCs Then

            'Si hay npcs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredNPCs
                Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).Amount)
                Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).NpcIndex)

                'Si es una quest ya empezada, entonces mandamos los NPCs que matí.
                If QuestSlot Then
                    Call .WriteInteger(UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i))

                End If

            Next i

        End If
        
        'Enviamos la cantidad de objs requeridos
        Call .WriteByte(QuestList(QuestIndex).RequiredOBJs)

        If QuestList(QuestIndex).RequiredOBJs Then

            'Si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredOBJs
                Call .WriteInteger(QuestList(QuestIndex).RequiredOBJ(i).Amount)
                Call .WriteInteger(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex)
            Next i

        End If
    
        'Enviamos la recompensa de oro y experiencia.
        Call .WriteLong(QuestList(QuestIndex).RewardGLD)
        Call .WriteLong(QuestList(QuestIndex).RewardEXP)
        
        'Enviamos la cantidad de objs de recompensa
        Call .WriteByte(QuestList(QuestIndex).RewardOBJs)

        If QuestList(QuestIndex).RewardOBJs Then

            'si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RewardOBJs
                Call .WriteInteger(QuestList(QuestIndex).RewardOBJ(i).Amount)
                Call .WriteInteger(QuestList(QuestIndex).RewardOBJ(i).ObjIndex)
            Next i

        End If

    End With

    Exit Sub
 
Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub
 
Public Sub WriteQuestListSend(ByVal UserIndex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Envía el paquete QuestList y la informaciín correspondiente.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i       As Integer

    Dim tmpStr  As String

    Dim tmpByte As Byte
 
    On Error GoTo Errhandler
 
    With UserList(UserIndex)
        .outgoingData.WriteByte ServerPacketID.QuestListSend
    
        For i = 1 To MAXUSERQUESTS

            If .QuestStats.Quests(i).QuestIndex Then
                tmpByte = tmpByte + 1
                tmpStr = tmpStr & QuestList(.QuestStats.Quests(i).QuestIndex).nombre & "-"

            End If

        Next i
        
        'Escribimos la cantidad de quests
        Call .outgoingData.WriteByte(tmpByte)
        
        'Escribimos la lista de quests (sacamos el íltimo caracter)
        If tmpByte Then
            Call .outgoingData.WriteASCIIString(Left$(tmpStr, Len(tmpStr) - 1))

        End If

    End With

    Exit Sub
 
Errhandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

