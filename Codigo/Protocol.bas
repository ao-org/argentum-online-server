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

    logged                  ' LOGGED  0
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    EquiteToggle
    CreateRenderText
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
    ChangeMap               ' CM
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
   'GM messages
    SpawnList               ' SPL
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
    EfectOverHEad '120
    EfectToScreen
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
    ExpOverHEad
    OroOverHEad
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
    Day
    SetTime
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
    Home
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

Public Function HandleIncomingData(ByVal Userindex As Integer) As Boolean
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
    packetID = CLng(UserList(Userindex).incomingData.PeekByte())

    'frmMain.listaDePaquetes.AddItem "Paq:" & PaquetesCount & ": " & packetID
    
    ' Debug.Print "Llego paquete ní" & packetID & " pesa: " & UserList(UserIndex).incomingData.length & "Bytes"
    
    'Does the packet requires a logged user??
    If Not (packetID = ClientPacketID.LoginExistingChar Or packetID = ClientPacketID.LoginNewChar Or packetID = ClientPacketID.CrearNuevaCuenta Or packetID = ClientPacketID.IngresarConCuenta Or packetID = ClientPacketID.RevalidarCuenta Or packetID = ClientPacketID.BorrarPJ Or packetID = ClientPacketID.RecuperandoContraseña Or packetID = ClientPacketID.BorrandoCuenta Or packetID = ClientPacketID.ValidarCuenta Or packetID = ClientPacketID.ThrowDice) Then
        
        'Is the user actually logged?
        If Not UserList(Userindex).flags.UserLogged Then
            Call CloseSocket(Userindex)
            Exit Function
        
            'He is logged. Reset idle counter if id is valid.
        ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
            UserList(Userindex).Counters.IdleCount = 0

        End If

    Else
    
        UserList(Userindex).Counters.IdleCount = 0
        
        ' Envió el primer paquete
        UserList(Userindex).flags.FirstPacket = True

    End If
    
    Select Case packetID
        
        Case ClientPacketID.Walk                    'M
            Call HandleWalk(Userindex)
    
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(Userindex)
        
        Case ClientPacketID.Yell                    '-
            Call HandleYell(Userindex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(Userindex)
    
        Case ClientPacketID.LoginExistingChar       'OLOGIN
            Call HandleLoginExistingChar(Userindex)
            
        Case ClientPacketID.CrearNuevaCuenta
            Call HandleCrearCuenta(Userindex)
        
        Case ClientPacketID.IngresarConCuenta
            Call HandleIngresarConCuenta(Userindex)
            
        Case ClientPacketID.ValidarCuenta
            Call HandleValidarCuenta(Userindex)
            
        Case ClientPacketID.RevalidarCuenta
            Call HandleReValidarCuenta(Userindex)
            
        Case ClientPacketID.BorrarPJ
            Call HandleBorrarPJ(Userindex)
            
        Case ClientPacketID.RecuperandoContraseña
            Call HandleRecuperandoContraseña(Userindex)
        
        Case ClientPacketID.BorrandoCuenta
         
            Call HandleBorrandoCuenta(Userindex)
        
        Case ClientPacketID.LoginNewChar            'NLOGIN
            Call HandleLoginNewChar(Userindex)
            
        Case ClientPacketID.ThrowDice
            Call HandleThrowDice(Userindex)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(Userindex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(Userindex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(Userindex)
        
        Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
            Call HandleSafeToggle(Userindex)
        
        Case ClientPacketID.PartySafeToggle
            Call HandlePartyToggle(Userindex)
        
        Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
            Call HandleRequestGuildLeaderInfo(Userindex)
        
        Case ClientPacketID.RequestAtributes        'ATR
            Call HandleRequestAtributes(Userindex)
           
        Case ClientPacketID.RequestSkills           'ESKI
            Call HandleRequestSkills(Userindex)
        
        Case ClientPacketID.RequestMiniStats        'FEST
            Call HandleRequestMiniStats(Userindex)
        
        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(Userindex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(Userindex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(Userindex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(Userindex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(Userindex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(Userindex)
        
        Case ClientPacketID.CastSpell               'LH
            Call HandleCastSpell(Userindex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(Userindex)
        
        Case ClientPacketID.DoubleClick             'RC
            Call HandleDoubleClick(Userindex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(Userindex)
        
        Case ClientPacketID.UseSpellMacro           'UMH
            Call HandleUseSpellMacro(Userindex)
        
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(Userindex)
        
        Case ClientPacketID.CraftBlacksmith         'CNS
            Call HandleCraftBlacksmith(Userindex)
        
        Case ClientPacketID.CraftCarpenter          'CNC
            Call HandleCraftCarpenter(Userindex)
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(Userindex)
        
        Case ClientPacketID.CreateNewGuild          'CIG
            Call HandleCreateNewGuild(Userindex)
        
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(Userindex)
        
        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(Userindex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(Userindex)
        
        Case ClientPacketID.ModifySkills            'SKSE
            Call HandleModifySkills(Userindex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(Userindex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(Userindex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(Userindex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(Userindex)
        
        Case ClientPacketID.BankDeposit             'DEPO
            Call HandleBankDeposit(Userindex)
        
        Case ClientPacketID.ForumPost               'DEMSG
            Call HandleForumPost(Userindex)
        
        Case ClientPacketID.MoveSpell               'DESPHE
            Call HandleMoveSpell(Userindex)
        
        Case ClientPacketID.ClanCodexUpdate         'DESCOD
            Call HandleClanCodexUpdate(Userindex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(Userindex)
        
        Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
            Call HandleGuildAcceptPeace(Userindex)
        
        Case ClientPacketID.GuildRejectAlliance     'RECPALIA
            Call HandleGuildRejectAlliance(Userindex)
        
        Case ClientPacketID.GuildRejectPeace        'RECPPEAT
            Call HandleGuildRejectPeace(Userindex)
        
        Case ClientPacketID.GuildAcceptAlliance     'ACEPALIA
            Call HandleGuildAcceptAlliance(Userindex)
        
        Case ClientPacketID.GuildOfferPeace         'PEACEOFF
            Call HandleGuildOfferPeace(Userindex)
        
        Case ClientPacketID.GuildOfferAlliance      'ALLIEOFF
            Call HandleGuildOfferAlliance(Userindex)
        
        Case ClientPacketID.GuildAllianceDetails    'ALLIEDET
            Call HandleGuildAllianceDetails(Userindex)
        
        Case ClientPacketID.GuildPeaceDetails       'PEACEDET
            Call HandleGuildPeaceDetails(Userindex)
        
        Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
            Call HandleGuildRequestJoinerInfo(Userindex)
        
        Case ClientPacketID.GuildAlliancePropList   'ENVALPRO
            Call HandleGuildAlliancePropList(Userindex)
        
        Case ClientPacketID.GuildPeacePropList      'ENVPROPP
            Call HandleGuildPeacePropList(Userindex)
        
        Case ClientPacketID.GuildDeclareWar         'DECGUERR
            Call HandleGuildDeclareWar(Userindex)
        
        Case ClientPacketID.GuildNewWebsite         'NEWWEBSI
            Call HandleGuildNewWebsite(Userindex)
        
        Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
            Call HandleGuildAcceptNewMember(Userindex)
        
        Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
            Call HandleGuildRejectNewMember(Userindex)
        
        Case ClientPacketID.GuildKickMember         'ECHARCLA
            Call HandleGuildKickMember(Userindex)
        
        Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
            Call HandleGuildUpdateNews(Userindex)
        
        Case ClientPacketID.GuildMemberInfo         '1HRINFO<
            Call HandleGuildMemberInfo(Userindex)
        
        Case ClientPacketID.GuildOpenElections      'ABREELEC
            Call HandleGuildOpenElections(Userindex)
        
        Case ClientPacketID.GuildRequestMembership  'SOLICITUD
            Call HandleGuildRequestMembership(Userindex)
        
        Case ClientPacketID.GuildRequestDetails     'CLANDETAILS
            Call HandleGuildRequestDetails(Userindex)
                  
        Case ClientPacketID.Online                  '/ONLINE
            Call HandleOnline(Userindex)
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(Userindex)
        
        Case ClientPacketID.GuildLeave              '/SALIRCLAN
            Call HandleGuildLeave(Userindex)
        
        Case ClientPacketID.RequestAccountState     '/BALANCE
            Call HandleRequestAccountState(Userindex)
        
        Case ClientPacketID.PetStand                '/QUIETO
            Call HandlePetStand(Userindex)
        
        Case ClientPacketID.PetFollow               '/ACOMPAÑAR
            Call HandlePetFollow(Userindex)
            
        Case ClientPacketID.PetLeave                '/LIBERAR
            Call HandlePetLeave(Userindex)
        
        Case ClientPacketID.GrupoMsg
            Call HandleGrupoMsg(Userindex)
        
        Case ClientPacketID.TrainList               '/ENTRENAR
            Call HandleTrainList(Userindex)
        
        Case ClientPacketID.Rest                    '/DESCANSAR
            Call HandleRest(Userindex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(Userindex)
        
        Case ClientPacketID.Resucitate              '/RESUCITAR
            Call HandleResucitate(Userindex)
        
        Case ClientPacketID.Heal                    '/CURAR
            Call HandleHeal(Userindex)
        
        Case ClientPacketID.Help                    '/AYUDA
            Call HandleHelp(Userindex)
        
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(Userindex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(Userindex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(Userindex)
        
        Case ClientPacketID.Enlist                  '/ENLISTAR
            Call HandleEnlist(Userindex)
        
        Case ClientPacketID.Information             '/INFORMACION
            Call HandleInformation(Userindex)
        
        Case ClientPacketID.Reward                  '/RECOMPENSA
            Call HandleReward(Userindex)
        
        Case ClientPacketID.RequestMOTD             '/MOTD
            Call HandleRequestMOTD(Userindex)
        
        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(Userindex)
                
        Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
            Call HandleInquiry(Userindex)
        
        Case ClientPacketID.GuildMessage            '/CMSG
            Call HandleGuildMessage(Userindex)
        
        Case ClientPacketID.CentinelReport          '/CENTINELA
            Call HandleCentinelReport(Userindex)
        
        Case ClientPacketID.GuildOnline             '/ONLINECLAN
            Call HandleGuildOnline(Userindex)
        
        Case ClientPacketID.CouncilMessage          '/BMSG
            Call HandleCouncilMessage(Userindex)
        
        Case ClientPacketID.RoleMasterRequest       '/ROL
            Call HandleRoleMasterRequest(Userindex)
        
        Case ClientPacketID.GMRequest               '/GM
            Call HandleGMRequest(Userindex)

        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(Userindex)
        
        Case ClientPacketID.GuildVote               '/VOTO
            Call HandleGuildVote(Userindex)
        
        Case ClientPacketID.punishments             '/PENAS
            Call HandlePunishments(Userindex)
        
        Case ClientPacketID.ChangePassword          '/Contraseña
            Call HandleChangePassword(Userindex)
        
        Case ClientPacketID.Gamble                  '/APOSTAR
            Call HandleGamble(Userindex)
        
        Case ClientPacketID.InquiryVote             '/ENCUESTA ( with parameters )
            Call HandleInquiryVote(Userindex)
        
        Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
            Call HandleLeaveFaction(Userindex)
        
        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
            Call HandleBankExtractGold(Userindex)
        
        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
            Call HandleBankDepositGold(Userindex)
        
        Case ClientPacketID.Denounce                '/DENUNCIAR
            Call HandleDenounce(Userindex)
        
        Case ClientPacketID.GuildMemberList         '/MIEMBROSCLAN
            Call HandleGuildMemberList(Userindex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(Userindex)
        
            'GM messages
        Case ClientPacketID.GMMessage               '/GMSG
            Call HandleGMMessage(Userindex)
        
        Case ClientPacketID.showName                '/SHOWNAME
            Call HandleShowName(Userindex)
        
        Case ClientPacketID.OnlineRoyalArmy         '/ONLINEREAL
            Call HandleOnlineRoyalArmy(Userindex)
        
        Case ClientPacketID.OnlineChaosLegion       '/ONLINECAOS
            Call HandleOnlineChaosLegion(Userindex)
        
        Case ClientPacketID.GoNearby                '/IRCERCA
            Call HandleGoNearby(Userindex)
        
        Case ClientPacketID.comment                 '/REM
            Call HandleComment(Userindex)
        
        Case ClientPacketID.serverTime              '/HORA
            Call HandleServerTime(Userindex)
        
        Case ClientPacketID.Where                   '/DONDE
            Call HandleWhere(Userindex)
        
        Case ClientPacketID.CreaturesInMap          '/NENE
            Call HandleCreaturesInMap(Userindex)
        
        Case ClientPacketID.WarpMeToTarget          '/TELEPLOC
            Call HandleWarpMeToTarget(Userindex)
        
        Case ClientPacketID.WarpChar                '/TELEP
            Call HandleWarpChar(Userindex)
        
        Case ClientPacketID.Silence                 '/SILENCIAR
            Call HandleSilence(Userindex)
        
        Case ClientPacketID.SOSShowList             '/SHOW SOS
            Call HandleSOSShowList(Userindex)
        
        Case ClientPacketID.SOSRemove               'SOSDONE
            Call HandleSOSRemove(Userindex)
        
        Case ClientPacketID.GoToChar                '/IRA
            Call HandleGoToChar(Userindex)
            
        Case ClientPacketID.Desbuggear              '/DESBUGGEAR
            Call HandleDesbuggear(Userindex)
            
        Case ClientPacketID.DarLlaveAUsuario        '/DARLLAVE
            Call HandleDarLlaveAUsuario(Userindex)
            
        Case ClientPacketID.SacarLlave              '/SACARLLAVE
            Call HandleSacarLlave(Userindex)
            
        Case ClientPacketID.VerLlaves               '/VERLLAVES
            Call HandleVerLlaves(Userindex)
            
        Case ClientPacketID.UseKey
            Call HandleUseKey(Userindex)
        
        Case ClientPacketID.invisible               '/INVISIBLE
            Call HandleInvisible(Userindex)
        
        Case ClientPacketID.GMPanel                 '/PANELGM
            Call HandleGMPanel(Userindex)
        
        Case ClientPacketID.RequestUserList         'LISTUSU
            Call HandleRequestUserList(Userindex)
        
        Case ClientPacketID.Working                 '/TRABAJANDO
            Call HandleWorking(Userindex)
        
        Case ClientPacketID.Hiding                  '/OCULTANDO
            Call HandleHiding(Userindex)
        
        Case ClientPacketID.Jail                    '/CARCEL
            Call HandleJail(Userindex)
        
        Case ClientPacketID.KillNPC                 '/RMATA
            Call HandleKillNPC(Userindex)
        
        Case ClientPacketID.WarnUser                '/ADVERTENCIA
            Call HandleWarnUser(Userindex)
        
        Case ClientPacketID.EditChar                '/MOD
            Call HandleEditChar(Userindex)
            
        Case ClientPacketID.RequestCharInfo         '/INFO
            Call HandleRequestCharInfo(Userindex)
        
        Case ClientPacketID.RequestCharStats        '/STAT
            Call HandleRequestCharStats(Userindex)
            
        Case ClientPacketID.RequestCharGold         '/BAL
            Call HandleRequestCharGold(Userindex)
            
        Case ClientPacketID.RequestCharInventory    '/INV
            Call HandleRequestCharInventory(Userindex)
            
        Case ClientPacketID.RequestCharBank         '/BOV
            Call HandleRequestCharBank(Userindex)
        
        Case ClientPacketID.RequestCharSkills       '/SKILLS
            Call HandleRequestCharSkills(Userindex)
        
        Case ClientPacketID.ReviveChar              '/REVIVIR
            Call HandleReviveChar(Userindex)
        
        Case ClientPacketID.OnlineGM                '/ONLINEGM
            Call HandleOnlineGM(Userindex)
        
        Case ClientPacketID.OnlineMap               '/ONLINEMAP
            Call HandleOnlineMap(Userindex)
        
        Case ClientPacketID.Forgive                 '/PERDON
            Call HandleForgive(Userindex)
            
        Case ClientPacketID.Kick                    '/ECHAR
            Call HandleKick(Userindex)
            
        Case ClientPacketID.Execute                 '/EJECUTAR
            Call HandleExecute(Userindex)
            
        Case ClientPacketID.BanChar                 '/BAN
            Call HandleBanChar(Userindex)
            
        Case ClientPacketID.SilenciarUser               '/BAN
            Call HandleSilenciarUser(Userindex)
            
        Case ClientPacketID.UnbanChar               '/UNBAN
            Call HandleUnbanChar(Userindex)
            
        Case ClientPacketID.NPCFollow               '/SEGUIR
            Call HandleNPCFollow(Userindex)
            
        Case ClientPacketID.SummonChar              '/SUM
            Call HandleSummonChar(Userindex)
            
        Case ClientPacketID.SpawnListRequest        '/CC
            Call HandleSpawnListRequest(Userindex)
            
        Case ClientPacketID.SpawnCreature           'SPA
            Call HandleSpawnCreature(Userindex)
            
        Case ClientPacketID.ResetNPCInventory       '/RESETINV
            Call HandleResetNPCInventory(Userindex)
            
        Case ClientPacketID.CleanWorld              '/LIMPIAR
            Call HandleCleanWorld(Userindex)
            
        Case ClientPacketID.ServerMessage           '/RMSG
            Call HandleServerMessage(Userindex)
            
        Case ClientPacketID.NickToIP                '/NICK2IP
            Call HandleNickToIP(Userindex)
        
        Case ClientPacketID.IPToNick                '/IP2NICK
            Call HandleIPToNick(Userindex)
            
        Case ClientPacketID.GuildOnlineMembers      '/ONCLAN
            Call HandleGuildOnlineMembers(Userindex)
        
        Case ClientPacketID.TeleportCreate          '/CT
            Call HandleTeleportCreate(Userindex)
            
        Case ClientPacketID.TeleportDestroy         '/DT
            Call HandleTeleportDestroy(Userindex)
            
        Case ClientPacketID.RainToggle              '/LLUVIA
            Call HandleRainToggle(Userindex)
        
        Case ClientPacketID.SetCharDescription      '/SETDESC
            Call HandleSetCharDescription(Userindex)
        
        Case ClientPacketID.ForceMIDIToMap          '/FORCEMIDIMAP
            Call HanldeForceMIDIToMap(Userindex)
            
        Case ClientPacketID.ForceWAVEToMap          '/FORCEWAVMAP
            Call HandleForceWAVEToMap(Userindex)
            
        Case ClientPacketID.RoyalArmyMessage        '/REALMSG
            Call HandleRoyalArmyMessage(Userindex)
                        
        Case ClientPacketID.ChaosLegionMessage      '/CAOSMSG
            Call HandleChaosLegionMessage(Userindex)
            
        Case ClientPacketID.CitizenMessage          '/CIUMSG
            Call HandleCitizenMessage(Userindex)
            
        Case ClientPacketID.CriminalMessage         '/CRIMSG
            Call HandleCriminalMessage(Userindex)
            
        Case ClientPacketID.TalkAsNPC               '/TALKAS
            Call HandleTalkAsNPC(Userindex)
        
        Case ClientPacketID.DestroyAllItemsInArea   '/MASSDEST
            Call HandleDestroyAllItemsInArea(Userindex)
            
        Case ClientPacketID.AcceptRoyalCouncilMember '/ACEPTCONSE
            Call HandleAcceptRoyalCouncilMember(Userindex)
            
        Case ClientPacketID.AcceptChaosCouncilMember '/ACEPTCONSECAOS
            Call HandleAcceptChaosCouncilMember(Userindex)
            
        Case ClientPacketID.ItemsInTheFloor         '/PISO
            Call HandleItemsInTheFloor(Userindex)
            
        Case ClientPacketID.MakeDumb                '/ESTUPIDO
            Call HandleMakeDumb(Userindex)
            
        Case ClientPacketID.MakeDumbNoMore          '/NOESTUPIDO
            Call HandleMakeDumbNoMore(Userindex)
            
        Case ClientPacketID.DumpIPTables            '/DUMPSECURITY"
            Call HandleDumpIPTables(Userindex)
            
        Case ClientPacketID.CouncilKick             '/KICKCONSE
            Call HandleCouncilKick(Userindex)
        
        Case ClientPacketID.SetTrigger              '/TRIGGER
            Call HandleSetTrigger(Userindex)
        
        Case ClientPacketID.AskTrigger               '/TRIGGER
            Call HandleAskTrigger(Userindex)
            
        Case ClientPacketID.BannedIPList            '/BANIPLIST
            Call HandleBannedIPList(Userindex)
        
        Case ClientPacketID.BannedIPReload          '/BANIPRELOAD
            Call HandleBannedIPReload(Userindex)
        
        Case ClientPacketID.GuildBan                '/BANCLAN
            Call HandleGuildBan(Userindex)
        
        Case ClientPacketID.BanIP                   '/BANIP
            Call HandleBanIP(Userindex)
        
        Case ClientPacketID.UnbanIP                 '/UNBANIP
            Call HandleUnbanIP(Userindex)
        
        Case ClientPacketID.CreateItem              '/CI
            Call HandleCreateItem(Userindex)
        
        Case ClientPacketID.DestroyItems            '/DEST
            Call HandleDestroyItems(Userindex)
        
        Case ClientPacketID.ChaosLegionKick         '/NOCAOS
            Call HandleChaosLegionKick(Userindex)
        
        Case ClientPacketID.RoyalArmyKick           '/NOREAL
            Call HandleRoyalArmyKick(Userindex)
        
        Case ClientPacketID.ForceMIDIAll            '/FORCEMIDI
            Call HandleForceMIDIAll(Userindex)
        
        Case ClientPacketID.ForceWAVEAll            '/FORCEWAV
            Call HandleForceWAVEAll(Userindex)
        
        Case ClientPacketID.RemovePunishment        '/BORRARPENA
            Call HandleRemovePunishment(Userindex)
        
        Case ClientPacketID.TileBlockedToggle       '/BLOQ
            Call HandleTileBlockedToggle(Userindex)
        
        Case ClientPacketID.KillNPCNoRespawn        '/MATA
            Call HandleKillNPCNoRespawn(Userindex)
        
        Case ClientPacketID.KillAllNearbyNPCs       '/MASSKILL
            Call HandleKillAllNearbyNPCs(Userindex)
        
        Case ClientPacketID.LastIP                  '/LASTIP
            Call HandleLastIP(Userindex)
        
        Case ClientPacketID.ChangeMOTD              '/MOTDCAMBIA
            Call HandleChangeMOTD(Userindex)
        
        Case ClientPacketID.SetMOTD                 'ZMOTD
            Call HandleSetMOTD(Userindex)
        
        Case ClientPacketID.SystemMessage           '/SMSG
            Call HandleSystemMessage(Userindex)
        
        Case ClientPacketID.CreateNPC               '/ACC
            Call HandleCreateNPC(Userindex)
        
        Case ClientPacketID.CreateNPCWithRespawn    '/RACC
            Call HandleCreateNPCWithRespawn(Userindex)
        
        Case ClientPacketID.ImperialArmour          '/AI1 - 4
            Call HandleImperialArmour(Userindex)
        
        Case ClientPacketID.ChaosArmour             '/AC1 - 4
            Call HandleChaosArmour(Userindex)
        
        Case ClientPacketID.NavigateToggle          '/NAVE
            Call HandleNavigateToggle(Userindex)
        
        Case ClientPacketID.ServerOpenToUsersToggle '/HABILITAR
            Call HandleServerOpenToUsersToggle(Userindex)
        
        Case ClientPacketID.Participar           '/APAGAR
            Call HandleParticipar(Userindex)
        
        Case ClientPacketID.TurnCriminal            '/CONDEN
            Call HandleTurnCriminal(Userindex)
        
        Case ClientPacketID.ResetFactions           '/RAJAR
            Call HandleResetFactions(Userindex)
        
        Case ClientPacketID.RemoveCharFromGuild     '/RAJARCLAN
            Call HandleRemoveCharFromGuild(Userindex)
        
        Case ClientPacketID.RequestCharMail         '/LASTEMAIL
            Call HandleRequestCharMail(Userindex)
        
        Case ClientPacketID.AlterPassword           '/APASS
            Call HandleAlterPassword(Userindex)
        
        Case ClientPacketID.AlterMail               '/AEMAIL
            Call HandleAlterMail(Userindex)
        
        Case ClientPacketID.AlterName               '/ANAME
            Call HandleAlterName(Userindex)
        
        Case ClientPacketID.ToggleCentinelActivated '/CENTINELAACTIVADO
            Call HandleToggleCentinelActivated(Userindex)
        
        Case ClientPacketID.DoBackUp                '/DOBACKUP
            Call HandleDoBackUp(Userindex)
        
        Case ClientPacketID.ShowGuildMessages       '/SHOWCMSG
            Call HandleShowGuildMessages(Userindex)
        
        Case ClientPacketID.SaveMap                 '/GUARDAMAPA
            Call HandleSaveMap(Userindex)
        
        Case ClientPacketID.ChangeMapInfoPK         '/MODMAPINFO PK
            Call HandleChangeMapInfoPK(Userindex)
        
        Case ClientPacketID.ChangeMapInfoBackup     '/MODMAPINFO BACKUP
            Call HandleChangeMapInfoBackup(Userindex)
    
        Case ClientPacketID.ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
            Call HandleChangeMapInfoRestricted(Userindex)
            
        Case ClientPacketID.ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
            Call HandleChangeMapInfoNoMagic(Userindex)
            
        Case ClientPacketID.ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
            Call HandleChangeMapInfoNoInvi(Userindex)
            
        Case ClientPacketID.ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
            Call HandleChangeMapInfoNoResu(Userindex)
            
        Case ClientPacketID.ChangeMapInfoLand       '/MODMAPINFO TERRENO
            Call HandleChangeMapInfoLand(Userindex)
            
        Case ClientPacketID.ChangeMapInfoZone       '/MODMAPINFO ZONA
            Call HandleChangeMapInfoZone(Userindex)
        
        Case ClientPacketID.SaveChars               '/GRABAR
            Call HandleSaveChars(Userindex)
        
        Case ClientPacketID.CleanSOS                '/BORRAR SOS
            Call HandleCleanSOS(Userindex)
        
        Case ClientPacketID.ShowServerForm          '/SHOW INT
            Call HandleShowServerForm(Userindex)
            
        Case ClientPacketID.night                   '/NOCHE
            Call HandleNight(Userindex)

        Case ClientPacketID.Day                     '/DIA
            Call HandleDay(Userindex)

        Case ClientPacketID.SetTime                 '/HORA X
            Call HandleSetTime(Userindex)

        Case ClientPacketID.KickAllChars            '/ECHARTODOSPJS
            Call HandleKickAllChars(Userindex)
        
        Case ClientPacketID.RequestTCPStats         '/TCPESSTATS
            Call HandleRequestTCPStats(Userindex)
        
        Case ClientPacketID.ReloadNPCs              '/RELOADNPCS
            Call HandleReloadNPCs(Userindex)
        
        Case ClientPacketID.ReloadServerIni         '/RELOADSINI
            Call HandleReloadServerIni(Userindex)
        
        Case ClientPacketID.ReloadSpells            '/RELOADHECHIZOS
            Call HandleReloadSpells(Userindex)
        
        Case ClientPacketID.ReloadObjects           '/RELOADOBJ
            Call HandleReloadObjects(Userindex)
        
        Case ClientPacketID.Restart                 '/REINICIAR
            Call HandleRestart(Userindex)
        
        Case ClientPacketID.ResetAutoUpdate         '/AUTOUPDATE
            Call HandleResetAutoUpdate(Userindex)
        
        Case ClientPacketID.ChatColor               '/CHATCOLOR
            Call HandleChatColor(Userindex)
        
        Case ClientPacketID.Ignored                 '/IGNORADO
            Call HandleIgnored(Userindex)
        
        Case ClientPacketID.CheckSlot               '/SLOT
            Call HandleCheckSlot(Userindex)
            
            'Nuevo Ladder
            
        Case ClientPacketID.GlobalMessage     '/CONSOLA
            Call HandleGlobalMessage(Userindex)
        
        Case ClientPacketID.GlobalOnOff        '/GLOBAL
            Call HandleGlobalOnOff(Userindex)
        
        Case ClientPacketID.NewPacketID    'Los Nuevos Packs ID
            Call HandleIncomingDataNewPacks(Userindex)

        Case Else
            'ERROR : Abort!
            Call CloseSocket(Userindex)

    End Select

    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(Userindex).incomingData.Length > 0 And Err.Number = 0 Then
        HandleIncomingData = True
  
    ElseIf Err.Number <> 0 And Not Err.Number = UserList(Userindex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & Err.Number & " [" & Err.description & "] " & _
                      " Source: " & Err.source & vbTab & _
                      " HelpFile: " & Err.HelpFile & vbTab & _
                      " HelpContext: " & Err.HelpContext & vbTab & _
                      " LastDllError: " & Err.LastDllError & vbTab & _
                      " - UserIndex: " & Userindex & " - producido al manejar el paquete: " & CStr(packetID))
        
        Call CloseSocket(Userindex)
  
        HandleIncomingData = False
    Else
        'Flush buffer - send everything that has been written
        
        HandleIncomingData = False

    End If

End Function

Public Sub HandleIncomingDataNewPacks(ByVal Userindex As Integer)
        
        On Error GoTo HandleIncomingDataNewPacks_Err
        

        '***************************************************
        'Los nuevos Pack ID
        'Creado por Ladder con gran ayuda de Maraxus
        '04.12.08
        '***************************************************
        Dim packetID As Integer
    
100     packetID = UserList(Userindex).incomingData.PeekInteger() \ &H100
    
102     Select Case packetID

            Case NewPacksID.OfertaInicial
104             Call HandleOfertaInicial(Userindex)
    
106         Case NewPacksID.OfertaDeSubasta
108             Call HandleOfertaDeSubasta(Userindex)
        
110         Case NewPacksID.CuentaRegresiva
112             Call HandleCuentaRegresiva(Userindex)

114         Case NewPacksID.QuestionGM
116             Call HandleQuestionGM(Userindex)

118         Case NewPacksID.PossUser
120             Call HandlePossUser(Userindex)

122         Case NewPacksID.Duelo
                'Call HandleDuelo(UserIndex)

124         Case NewPacksID.NieveToggle
126             Call HandleNieveToggle(Userindex)

128         Case NewPacksID.NieblaToggle
130             Call HandleNieblaToggle(Userindex)

132         Case NewPacksID.TransFerGold
134             Call HandleTransFerGold(Userindex)

136         Case NewPacksID.Moveitem
138             Call HandleMoveItem(Userindex)

140         Case NewPacksID.LlamadadeClan
142             Call HandleLlamadadeClan(Userindex)

144         Case NewPacksID.QuieroFundarClan
146             Call HandleQuieroFundarClan(Userindex)

148         Case NewPacksID.BovedaMoveItem
150             Call HandleBovedaMoveItem(Userindex)

152         Case NewPacksID.Genio
154             Call HandleGenio(Userindex)

156         Case NewPacksID.Casarse
158             Call HandleCasamiento(Userindex)

160         Case NewPacksID.EnviarCodigo
162             Call HandleEnviarCodigo(Userindex)

164         Case NewPacksID.CrearTorneo
166             Call HandleCrearTorneo(Userindex)
            
168         Case NewPacksID.ComenzarTorneo
170             Call HandleComenzarTorneo(Userindex)
            
172         Case NewPacksID.CancelarTorneo
174             Call HandleCancelarTorneo(Userindex)

176         Case NewPacksID.BusquedaTesoro
178             Call HandleBusquedaTesoro(Userindex)
            
180         Case NewPacksID.CrearEvento
182             Call HandleCrearEvento(Userindex)

184         Case NewPacksID.CraftAlquimista
186             Call HandleCraftAlquimia(Userindex)

188         Case NewPacksID.DropItem
190             Call HandleDropItem(Userindex)

192         Case NewPacksID.RequestFamiliar
194             Call HandleRequestFamiliar(Userindex)

196         Case NewPacksID.FlagTrabajar
198             Call HandleFlagTrabajar(Userindex)

200         Case NewPacksID.CraftSastre
202             Call HandleCraftSastre(Userindex)

204         Case NewPacksID.MensajeUser
206             Call HandleMensajeUser(Userindex)

208         Case NewPacksID.TraerBoveda
210             Call HandleTraerBoveda(Userindex)

212         Case NewPacksID.CompletarAccion
214             Call HandleCompletarAccion(Userindex)

216         Case NewPacksID.Escribiendo
218             Call HandleEscribiendo(Userindex)

220         Case NewPacksID.TraerRecompensas
222             Call HandleTraerRecompensas(Userindex)

224         Case NewPacksID.ReclamarRecompensa
226             Call HandleReclamarRecompensa(Userindex)

232         Case NewPacksID.Correo
234             Call HandleCorreo(Userindex)

236         Case NewPacksID.SendCorreo ' ok
238             Call HandleSendCorreo(Userindex)

240         Case NewPacksID.RetirarItemCorreo ' ok
242             Call HandleRetirarItemCorreo(Userindex)

244         Case NewPacksID.BorrarCorreo
246             Call HandleBorrarCorreo(Userindex) 'ok

248         Case NewPacksID.InvitarGrupo
250             Call HandleInvitarGrupo(Userindex) 'ok

252         Case NewPacksID.MarcaDeClanPack
254             Call HandleMarcaDeClan(Userindex)

256         Case NewPacksID.MarcaDeGMPack
258             Call HandleMarcaDeGM(Userindex)

260         Case NewPacksID.ResponderPregunta 'ok
262             Call HandleResponderPregunta(Userindex)

264         Case NewPacksID.RequestGrupo
266             Call HandleRequestGrupo(Userindex) 'ok

268         Case NewPacksID.AbandonarGrupo
270             Call HandleAbandonarGrupo(Userindex) ' ok

272         Case NewPacksID.HecharDeGrupo
274             Call HandleHecharDeGrupo(Userindex) 'ok

276         Case NewPacksID.MacroPossent
278             Call HandleMacroPos(Userindex)

280         Case NewPacksID.SubastaInfo
282             Call HandleSubastaInfo(Userindex)

284         Case NewPacksID.EventoInfo
286             Call HandleEventoInfo(Userindex)

288         Case NewPacksID.CrearEvento
290             Call HandleCrearEvento(Userindex)

292         Case NewPacksID.BanCuenta
294             Call HandleBanCuenta(Userindex)
            
296         Case NewPacksID.unBanCuenta
298             Call HandleUnBanCuenta(Userindex)
            
300         Case NewPacksID.BanSerial
302             Call HandleBanSerial(Userindex)
        
304         Case NewPacksID.unBanSerial
306             Call HandleUnBanSerial(Userindex)
            
308         Case NewPacksID.CerrarCliente
310             Call HandleCerrarCliente(Userindex)
            
312         Case NewPacksID.BanTemporal
314             Call HandleBanTemporal(Userindex)

316         Case NewPacksID.Traershop
318             Call HandleTraerShop(Userindex)

320         Case NewPacksID.TraerRanking
322             Call HandleTraerRanking(Userindex)

324         Case NewPacksID.Pareja
326             Call HandlePareja(Userindex)
            
328         Case NewPacksID.ComprarItem
330             Call HandleComprarItem(Userindex)
            
332         Case NewPacksID.CompletarViaje
334             Call HandleCompletarViaje(Userindex)
            
336         Case NewPacksID.ScrollInfo
338             Call HandleScrollInfo(Userindex)

340         Case NewPacksID.CancelarExit
342             Call HandleCancelarExit(Userindex)
            
344         Case NewPacksID.Quest
346             Call HandleQuest(Userindex)
            
348         Case NewPacksID.QuestAccept
350             Call HandleQuestAccept(Userindex)
        
352         Case NewPacksID.QuestListRequest
354             Call HandleQuestListRequest(Userindex)
        
356         Case NewPacksID.QuestDetailsRequest
358             Call HandleQuestDetailsRequest(Userindex)
        
360         Case NewPacksID.QuestAbandon
362             Call HandleQuestAbandon(Userindex)
            
364         Case NewPacksID.SeguroClan
366             Call HandleSeguroClan(Userindex)
            
368         Case NewPacksID.CreatePretorianClan     '/CREARPRETORIANOS
370             Call HandleCreatePretorianClan(Userindex)
         
372         Case NewPacksID.RemovePretorianClan     '/ELIMINARPRETORIANOS
374             Call HandleDeletePretorianClan(Userindex)

            Case NewPacksID.Home
                Call HandleHome(Userindex)
            
376         Case Else
                'ERROR : Abort!
378             Call CloseSocket(Userindex)
            
        End Select
    
380     If UserList(Userindex).incomingData.Length > 0 And Err.Number = 0 Then
382         Err.Clear
384         Call HandleIncomingData(Userindex)
    
386     ElseIf Err.Number <> 0 And Not Err.Number = UserList(Userindex).incomingData.NotEnoughDataErrCode Then
            'An error ocurred, log it and kick player.
388         Call LogError("Error: " & Err.Number & " [" & Err.description & "] " & _
                          " Source: " & Err.source & vbTab & _
                          " HelpFile: " & Err.HelpFile & vbTab & _
                          " HelpContext: " & Err.HelpContext & vbTab & _
                          " LastDllError: " & Err.LastDllError & vbTab & _
                          " - UserIndex: " & Userindex & " - producido al manejar el paquete: " & CStr(packetID))
                          
390         Call CloseSocket(Userindex)
    
        End If

        
        Exit Sub

HandleIncomingDataNewPacks_Err:
392     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleIncomingDataNewPacks", Erl)
394     Resume Next
        
End Sub

''
' Handles the "LoginExistingChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    ''Last Modification: 01/12/08 Ladder
    '***************************************************
    If UserList(Userindex).incomingData.Length < 16 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
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
        Call WriteShowMessageBox(Userindex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If

    If Not EntrarCuenta(Userindex, CuentaEmail, Password, MacAddress, HDserial) Then
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If
    
    If Not AsciiValidos(UserName) Then
        Call WriteShowMessageBox(Userindex, "Nombre invalido.")
        
        Call CloseSocket(Userindex)
        
        Exit Sub

    End If
    
    If Not PersonajeExiste(UserName) Then
        Call WriteShowMessageBox(Userindex, "El personaje no existe.")
        
        Call CloseSocket(Userindex)
        
        Exit Sub

    End If
    
    If BANCheck(UserName) Then

        Dim LoopC As Integer
        
        For LoopC = 1 To Baneos.Count

            If Baneos(LoopC).name = UCase$(UserName) Then
                Call WriteShowMessageBox(Userindex, "Se te ha prohibido la entrada a Argentum20 hasta el día " & Format(Baneos(LoopC).FechaLiberacion, "dddddd") & " a las " & Format(Baneos(LoopC).FechaLiberacion, "hh:mm am/pm") & " debido a " & Baneos(LoopC).Causa & " Esta decisión fue tomada por " & Baneos(LoopC).Baneador & ".")
                
                Call CloseSocket(Userindex)
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
        
        Call WriteShowMessageBox(Userindex, "Se te ha prohibido la entrada al juego debido a " & BaneoMotivo & ". Esta decisión fue tomada por " & BanNick & ".")
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If
        
    Call ConnectUser(Userindex, UserName, CuentaEmail)

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal Userindex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    If UserList(Userindex).incomingData.Length < 22 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim UserName As String

    Dim race     As eRaza

    Dim gender   As eGenero
    
    Dim Hogar   As eCiudad

    Dim Class As eClass

    Dim Head        As Integer

    Dim CuentaEmail As String

    Dim Password    As String

    Dim MacAddress  As String

    Dim HDserial    As Long

    Dim Version     As String
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(Userindex, "La creacion de personajes en este servidor se ha deshabilitado.")
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If

    If aClon.MaxPersonajes(UserList(Userindex).ip) Then
        Call WriteErrorMsg(Userindex, "Has creado demasiados personajes.")
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If
    
    If ObtenerCantidadDePersonajesByUserIndex(Userindex) >= MAX_PERSONAJES Then
        Call CloseSocket(Userindex)
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
    Hogar = buffer.ReadByte()
    MacAddress = buffer.ReadASCIIString()
    HDserial = buffer.ReadLong()
    
    If Not VersionOK(Version) Then
        Call WriteShowMessageBox(Userindex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If

    If Not EntrarCuenta(Userindex, CuentaEmail, Password, MacAddress, HDserial) Then
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If

    Call ConnectNewUser(Userindex, UserName, race, gender, Class, Head, CuentaEmail, Hogar)

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleThrowDice(ByVal Userindex As Integer)
        'Remove packet ID
        
        On Error GoTo HandleThrowDice_Err
        
100     Call UserList(Userindex).incomingData.ReadByte
    
102     With UserList(Userindex).Stats
104         .UserAtributos(eAtributos.Fuerza) = RandomNumber(DiceMinimum, DiceMaximum)
106         .UserAtributos(eAtributos.Agilidad) = RandomNumber(DiceMinimum, DiceMaximum)
108         .UserAtributos(eAtributos.Inteligencia) = RandomNumber(DiceMinimum, DiceMaximum)
            .UserAtributos(eAtributos.Carisma) = RandomNumber(DiceMinimum, DiceMaximum)
110         .UserAtributos(eAtributos.Constitucion) = RandomNumber(DiceMinimum, DiceMaximum)

        End With
    
112     Call WriteDiceRoll(Userindex)

        
        Exit Sub

HandleThrowDice_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleThrowDice", Erl)
        Resume Next
        
End Sub

''
' Handles the "Talk" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String: chat = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.name, "Dijo: " & chat)
        End If
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then

                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    .Char.Body = ObjData(.Invent.BarcoObjIndex).Ropaje

                    .Char.ShieldAnim = NingunEscudo
                    .Char.WeaponAnim = NingunArma
                    .Char.CascoAnim = NingunCasco

                    Call WriteConsoleMsg(Userindex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                    'Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(Userindex, "307", FontTypeNames.FONTTYPE_INFO)
    
                End If

            End If

        End If
       
        If .flags.Silenciado = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Los administradores te han impedido hablar durante los proximos " & .flags.MinutosRestantes & " minutos debido a tu comportamiento.", FontTypeNames.FONTTYPE_VENENO)
            Call WriteLocaleMsg(Userindex, "110", FontTypeNames.FONTTYPE_INFO, .flags.MinutosRestantes)
        Else

            If LenB(chat) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(chat)
                
                If .flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, CHAT_COLOR_DEAD_CHAR, UserList(Userindex).name))
                Else
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, .flags.ChatColor, UserList(Userindex).name))

                End If

            End If

        End If

    End With
    
ErrHandler:

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

Private Sub HandleYell(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String
        
        chat = buffer.ReadASCIIString()
        
        If UserList(Userindex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
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
                
                If .flags.Navegando = 1 Then
                    
                    'TODO: Revisar con WyroX
                    If .clase = eClass.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
                        .Char.Body = ObjData(.Invent.BarcoObjIndex).Ropaje
    
                        .Char.ShieldAnim = NingunEscudo
                        .Char.WeaponAnim = NingunArma
                        .Char.CascoAnim = NingunCasco
    
                        Call WriteConsoleMsg(Userindex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
    
                    End If
    
                Else
    
                    If .flags.invisible = 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                        Call WriteConsoleMsg(Userindex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
    
                    End If
    
                End If

            End If
            
            If .flags.Silenciado = 1 Then
                Call WriteLocaleMsg(Userindex, "110", FontTypeNames.FONTTYPE_VENENO, .flags.MinutosRestantes)
        
                'Call WriteConsoleMsg(UserIndex, "Los administradores te han impedido hablar durante los proximos " & .flags.MinutosRestantes & " minutos debido a tu comportamiento.", FontTypeNames.FONTTYPE_VENENO)
            Else

                If LenB(chat) <> 0 Then
                    'Analize chat...
                    Call Statistics.ParseChat(chat)

                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, vbRed, UserList(Userindex).name))
               
                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleWhisper(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat            As String

        Dim targetCharIndex As String

        Dim targetUserIndex As Integer

        Dim rank            As Integer
        
        rank = UserList(Userindex).flags.Privilegios

        targetCharIndex = buffer.ReadASCIIString()
        chat = buffer.ReadASCIIString()
        
        targetUserIndex = NameIndex(targetCharIndex)

        If targetUserIndex <= 0 Then 'existe el usuario destino?
            Call WriteConsoleMsg(Userindex, "Usuario offline o inexistente.", FontTypeNames.FONTTYPE_INFO)
        Else
        
            If rank = 1 And (UserList(targetUserIndex).flags.Privilegios) > 1 Then
                Call WriteConsoleMsg(Userindex, "No podes hablar por privado con administradores del juego.", FontTypeNames.FONTTYPE_WARNING)
            Else

                If EstaPCarea(Userindex, targetUserIndex) Then
                    If LenB(chat) <> 0 Then
                        'Analize chat...
                        Call Statistics.ParseChat(chat)
            
                        Call SendData(SendTarget.ToSuperiores, Userindex, PrepareMessageChatOverHead(chat, .Char.CharIndex, RGB(157, 226, 20)))
                        
                        Call WriteChatOverHead(Userindex, chat, .Char.CharIndex, RGB(157, 226, 20))
                        Call WriteChatOverHead(targetUserIndex, chat, .Char.CharIndex, RGB(157, 226, 20))
                        'Call WriteConsoleMsg(UserIndex, "[" & .Name & "] " & chat, FontTypeNames.FONTTYPE_MP)
                        'Call WriteConsoleMsg(targetUserIndex, "[" & .Name & "] " & chat, FontTypeNames.FONTTYPE_MP)
                        Call WritePlayWave(targetUserIndex, FXSound.MP_SOUND, NO_3D_SOUND, NO_3D_SOUND)
                        

                    End If

                Else
                    Call WriteConsoleMsg(Userindex, "[" & .name & "] " & chat, FontTypeNames.FONTTYPE_MP)
                    Call WriteConsoleMsg(targetUserIndex, "[" & .name & "] " & chat, FontTypeNames.FONTTYPE_MP)
                    Call WritePlayWave(targetUserIndex, FXSound.MP_SOUND, NO_3D_SOUND, NO_3D_SOUND)
                    
                    
                End If

            End If

        End If

        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleWalk(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        On Error GoTo HandleWalk_Err
        

        Dim demora      As Long

        Dim demorafinal As Long

100     demora = (timeGetTime And &H7FFFFFFF)

102     If UserList(Userindex).incomingData.Length < 2 Then
104         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
        Dim dummy    As Long

        Dim TempTick As Long

        Dim Heading  As eHeading
    
106     With UserList(Userindex)
            'Remove packet ID
108         Call .incomingData.ReadByte
        
110         Heading = .incomingData.ReadByte()
        
112         If .flags.Paralizado = 0 And .flags.Inmovilizado = 0 Then

114             If .flags.Meditando Then
                    'Stop meditating, next action will start movement.
116                 .flags.Meditando = False
120                 UserList(Userindex).Char.FX = 0
122                 Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageMeditateToggle(UserList(Userindex).Char.CharIndex, 0))
                End If
            
124             'If IntervaloPermiteCaminar(UserIndex) Then
            
                    'Move user
126                 Call MoveUserChar(Userindex, Heading)
                
128                 If UserList(Userindex).Grupo.EnGrupo = True Then
130                     Call CompartirUbicacion(Userindex)
                    End If

                    'Stop resting if needed
132                 If .flags.Descansar Then
134                     .flags.Descansar = False
                    
136                     Call WriteRestOK(Userindex)
                        'Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)
138                     Call WriteLocaleMsg(Userindex, "178", FontTypeNames.FONTTYPE_INFO)

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
170                             Call CloseSocket(Userindex)
                                
                                Exit Sub
                            Else
172                             .flags.CountSH = TempTick

                            End If

                        End If

174                     .flags.StartWalk = TempTick
176                     .flags.TimesWalk = 0

                    End If
                    
178                 .flags.TimesWalk = .flags.TimesWalk + 1
                
180                 Call CancelExit(Userindex)

                    'Esta usando el /HOGAR, no se puede mover
                    If .flags.Traveling = 1 Then
                        .flags.Traveling = 0
                        .Counters.goHome = 0
                        Call WriteConsoleMsg(Userindex, "Has cancelado el viaje a casa.", FontTypeNames.FONTTYPE_INFO)
                    End If

                'End If

            Else    'paralized

182             If Not .flags.UltimoMensaje = 1 Then
184                 .flags.UltimoMensaje = 1
                    'Call WriteConsoleMsg(UserIndex, "No podes moverte porque estas paralizado.", FontTypeNames.FONTTYPE_INFO)
186                 Call WriteLocaleMsg(Userindex, "54", FontTypeNames.FONTTYPE_INFO)

                End If
            
188             .flags.CountSH = 0

            End If
            
            'Can't move while hidden except he is a thief
190         If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
                
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
        
                            Call WriteConsoleMsg(Userindex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                            Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
    
                        End If
    
                    Else
    
                        'If not under a spell effect, show char
                        If .flags.invisible = 0 Then
                            Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteLocaleMsg(Userindex, "307", FontTypeNames.FONTTYPE_INFO)
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                        End If
    
                    End If
    
                End If
                
            End If

        End With
    
204     demorafinal = (timeGetTime And &H7FFFFFFF) - demora

        Exit Sub

HandleWalk_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleWalk", Erl)
        Resume Next
        
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleRequestPositionUpdate_Err
        
100     UserList(Userindex).incomingData.ReadByte
    
102     Call WritePosUpdate(Userindex)

        
        Exit Sub

HandleRequestPositionUpdate_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestPositionUpdate", Erl)
        Resume Next
        
End Sub

''
' Handles the "Attack" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal Userindex As Integer)
        
        On Error GoTo HandleAttack_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 10/01/08
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'If dead, can't attack
104         If .flags.Muerto = 1 Then
106             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "í¡No podes atacar a nadie porque estas muerto!!.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'If equiped weapon is ranged, can't attack this way
108         If .Invent.WeaponEqpObjIndex > 0 Then
110             If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
112                 Call WriteConsoleMsg(Userindex, "No podés usar así esta arma.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If

            End If
        
114         If .Invent.HerramientaEqpObjIndex > 0 Then
116             Call WriteConsoleMsg(Userindex, "Para atacar debes desequipar la herramienta.", FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If
        
118         If UserList(Userindex).flags.Meditando Then
120             UserList(Userindex).flags.Meditando = False
124             UserList(Userindex).Char.FX = 0
126             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageMeditateToggle(UserList(Userindex).Char.CharIndex, 0))
            End If
        
            'If exiting, cancel
128         Call CancelExit(Userindex)
        
            'Attack!
130         Call UsuarioAtaca(Userindex)
            
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
    
                        Call WriteConsoleMsg(Userindex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)
                    End If
    
                Else
    
                    If .flags.invisible = 0 Then
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(.Char.CharIndex, False))
                        'Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                        Call WriteLocaleMsg(Userindex, "307", FontTypeNames.FONTTYPE_INFOIAO)
    
                    End If
    
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

Private Sub HandlePickUp(ByVal Userindex As Integer)
        
        On Error GoTo HandlePickUp_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'If dead, it can't pick up objects
104         If .flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!! Los muertos no pueden tomar objetos.", FontTypeNames.FONTTYPE_INFO)
106             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
            
                Exit Sub

            End If
        
            'Lower rank administrators can't pick up items
108         If .flags.Privilegios And PlayerType.Consejero Then
110             If Not .flags.Privilegios And PlayerType.RoleMaster Then
112                 Call WriteConsoleMsg(Userindex, "No podés tomar ningun objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            End If
        
114         Call GetObj(Userindex)

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

Private Sub HandleSafeToggle(ByVal Userindex As Integer)
        
        On Error GoTo HandleSafeToggle_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Seguro Then
106             Call WriteSafeModeOff(Userindex)
            Else
108             Call WriteSafeModeOn(Userindex)

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

Private Sub HandlePartyToggle(ByVal Userindex As Integer)
        
        On Error GoTo HandlePartyToggle_Err
        

        '***************************************************
        'Author: Rapsodius
        'Creation Date: 10/10/07
        '***************************************************
100     With UserList(Userindex)
102         Call .incomingData.ReadByte
        
104         .flags.SeguroParty = Not .flags.SeguroParty
        
106         If .flags.SeguroParty Then
108             Call WritePartySafeOn(Userindex)
            Else
110             Call WritePartySafeOff(Userindex)

            End If

        End With

        
        Exit Sub

HandlePartyToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandlePartyToggle", Erl)
        Resume Next
        
End Sub

Private Sub HandleSeguroClan(ByVal Userindex As Integer)
        
        On Error GoTo HandleSeguroClan_Err
        

        '***************************************************
        'Author: Ladder
        'Date: 31/10/20
        '***************************************************
100     With UserList(Userindex)
102         Call .incomingData.ReadInteger 'Leemos paquete
                
104         .flags.SeguroClan = Not .flags.SeguroClan

106         Call WriteClanSeguro(Userindex, .flags.SeguroClan)

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

Private Sub HandleRequestGuildLeaderInfo(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleRequestGuildLeaderInfo_Err
        
100     UserList(Userindex).incomingData.ReadByte
    
102     Call modGuilds.SendGuildLeaderInfo(Userindex)

        
        Exit Sub

HandleRequestGuildLeaderInfo_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestGuildLeaderInfo", Erl)
        Resume Next
        
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleRequestAtributes_Err
        
100     Call UserList(Userindex).incomingData.ReadByte
    
102     Call WriteAttributes(Userindex)

        
        Exit Sub

HandleRequestAtributes_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestAtributes", Erl)
        Resume Next
        
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleRequestSkills_Err
        
100     Call UserList(Userindex).incomingData.ReadByte
    
102     Call WriteSendSkills(Userindex)

        
        Exit Sub

HandleRequestSkills_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestSkills", Erl)
        Resume Next
        
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleRequestMiniStats_Err
        
100     Call UserList(Userindex).incomingData.ReadByte
    
102     Call WriteMiniStats(Userindex)

        
        Exit Sub

HandleRequestMiniStats_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestMiniStats", Erl)
        Resume Next
        
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleCommerceEnd_Err
        
100     Call UserList(Userindex).incomingData.ReadByte
    
        'User quits commerce mode
102     If UserList(Userindex).flags.TargetNPC <> 0 Then
104         If Npclist(UserList(Userindex).flags.TargetNPC).SoundClose <> 0 Then
106             Call WritePlayWave(Userindex, Npclist(UserList(Userindex).flags.TargetNPC).SoundClose, NO_3D_SOUND, NO_3D_SOUND)

            End If

        End If

108     UserList(Userindex).flags.Comerciando = False
110     Call WriteCommerceEnd(Userindex)

        
        Exit Sub

HandleCommerceEnd_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCommerceEnd", Erl)
        Resume Next
        
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal Userindex As Integer)
        
        On Error GoTo HandleUserCommerceEnd_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Quits commerce mode with user
104         If .ComUsu.DestUsu > 0 And UserList(.ComUsu.DestUsu).ComUsu.DestUsu = Userindex Then
106             Call WriteConsoleMsg(.ComUsu.DestUsu, .name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
108             Call FinComerciarUsu(.ComUsu.DestUsu)
            
                'Send data in the outgoing buffer of the other user
            

            End If
        
110         Call FinComerciarUsu(Userindex)

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

Private Sub HandleBankEnd(ByVal Userindex As Integer)
        
        On Error GoTo HandleBankEnd_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'User exits banking mode
104         .flags.Comerciando = False
        
106         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave("171", UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
108         Call WriteBankEnd(Userindex)

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

Private Sub HandleUserCommerceOk(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleUserCommerceOk_Err
        
100     Call UserList(Userindex).incomingData.ReadByte
    
        'Trade accepted
102     Call AceptarComercioUsu(Userindex)

        
        Exit Sub

HandleUserCommerceOk_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleUserCommerceOk", Erl)
        Resume Next
        
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal Userindex As Integer)
        
        On Error GoTo HandleUserCommerceReject_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim otherUser As Integer
    
100     With UserList(Userindex)
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
        
114         Call WriteConsoleMsg(Userindex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
116         Call FinComerciarUsu(Userindex)

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

Private Sub HandleDrop(ByVal Userindex As Integer)
        
        On Error GoTo HandleDrop_Err
        
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 07/25/09
        '07/25/09: Marco - Agregue un checkeo para patear a los usuarios que tiran items mientras comercian.
        '***************************************************
        
100     If UserList(Userindex).incomingData.Length < 4 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub
        End If
    
        Dim slot   As Byte

        Dim Amount As Long
    
104     With UserList(Userindex)

            'Remove packet ID
106         Call .incomingData.ReadByte

108         slot = .incomingData.ReadByte()
110         Amount = .incomingData.ReadLong()

112         If Not IntervaloPermiteTirar(Userindex) Then Exit Sub

            'low rank admins can't drop item. Neither can the dead nor those sailing.
114         If .flags.Muerto = 1 Or ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub
            
            'If the user is trading, he can't drop items => He's cheating, we kick him.
            If .flags.Comerciando Then Exit Sub
    
            'Si esta navegando y no es pirata, no dejamos tirar items al agua.
            If .flags.Navegando = 1 And Not .clase = eClass.Pirat Then
                Call WriteConsoleMsg(Userindex, "Solo los Piratas pueden tirar items en altamar", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            'Are we dropping gold or other items??
116         If slot = FLAGORO Then

118             Call TirarOro(Amount, Userindex)
            
120             Call WriteUpdateGold(Userindex)
            Else
        
                '04-05-08 Ladder
122             If (.flags.Privilegios And PlayerType.Admin) <> 16 Then
124                 If ObjData(.Invent.Object(slot).ObjIndex).Newbie = 1 Then
126                     Call WriteConsoleMsg(Userindex, "No se pueden tirar los objetos Newbies.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
128                 If ObjData(.Invent.Object(slot).ObjIndex).Instransferible = 1 Then
130                     Call WriteConsoleMsg(Userindex, "Acción no permitida.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
            
132                 If ObjData(.Invent.Object(slot).ObjIndex).Intirable = 1 Then
134                     Call WriteConsoleMsg(Userindex, "Acción no permitida.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
136                 If UserList(Userindex).flags.BattleModo = 1 Then
138                     Call WriteConsoleMsg(Userindex, "No podes tirar items en este mapa.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                End If
        
140             If ObjData(.Invent.Object(slot).ObjIndex).OBJType = eOBJType.otBarcos And UserList(Userindex).flags.Navegando Then
142                 Call WriteConsoleMsg(Userindex, "Para tirar la barca deberias estar en tierra firme.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
        
144             If ObjData(.Invent.Object(slot).ObjIndex).OBJType = eOBJType.otMonturas And UserList(Userindex).flags.Montado Then
146                 Call WriteConsoleMsg(Userindex, "Para tirar tu montura deberias descender de ella.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                '04-05-08 Ladder
        
                'Only drop valid slots
148             If slot <= UserList(Userindex).CurrentInventorySlots And slot > 0 Then
150                 If .Invent.Object(slot).ObjIndex = 0 Then
                        Exit Sub

                    End If
                
152                 Call DropObj(Userindex, slot, Amount, .Pos.Map, .Pos.X, .Pos.Y)

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

Private Sub HandleCastSpell(ByVal Userindex As Integer)
        
        On Error GoTo HandleCastSpell_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Spell As Byte
        
108         Spell = .incomingData.ReadByte()
        
110         If .flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
112             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
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
            
126             uh = UserList(Userindex).Stats.UserHechizos(Spell)

128             If Hechizos(uh).AutoLanzar = 1 Then
130                 UserList(Userindex).flags.TargetUser = Userindex
132                 Call LanzarHechizo(.flags.Hechizo, Userindex)
                Else
134                 Call WriteWorkRequestTarget(Userindex, eSkill.magia)

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

Private Sub HandleLeftClick(ByVal Userindex As Integer)
        
        On Error GoTo HandleLeftClick_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 3 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex).incomingData
            'Remove packet ID
106         Call .ReadByte
        
            Dim X As Byte

            Dim Y As Byte
        
108         X = .ReadByte()
110         Y = .ReadByte()
        
112         Call LookatTile(Userindex, UserList(Userindex).Pos.Map, X, Y)

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

Private Sub HandleDoubleClick(ByVal Userindex As Integer)
        
        On Error GoTo HandleDoubleClick_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 3 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex).incomingData
            'Remove packet ID
106         Call .ReadByte
        
            Dim X As Byte

            Dim Y As Byte
        
108         X = .ReadByte()
110         Y = .ReadByte()
        
112         Call Accion(Userindex, UserList(Userindex).Pos.Map, X, Y)

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

Private Sub HandleWork(ByVal Userindex As Integer)
        
        On Error GoTo HandleWork_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 13/01/2010
        '13/01/2010: ZaMa - El pirata se puede ocultar en barca
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Skill As eSkill
        
108         Skill = .incomingData.ReadByte()
        
110         If UserList(Userindex).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
112             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'If exiting, cancel
114         Call CancelExit(Userindex)
        
116         Select Case Skill

                Case Robar, magia, Domar
118                 Call WriteWorkRequestTarget(Userindex, Skill)

120             Case Ocultarse

122                 If .flags.Navegando = 1 Then
                        
                        If .clase <> eClass.Pirat Then

                            If Not .flags.UltimoMensaje = 3 Then
                                'Call WriteConsoleMsg(UserIndex, "No podés ocultarte si estás navegando.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteLocaleMsg(Userindex, "56", FontTypeNames.FONTTYPE_INFO)
                                .flags.UltimoMensaje = 3
                            End If
                            
                            Exit Sub
                            
                        End If

                    End If
                
130                 If .flags.Montado = 1 Then

                        '[CDT 17-02-2004]
132                     If Not .flags.UltimoMensaje = 3 Then
134                         Call WriteConsoleMsg(Userindex, "No podés ocultarte si estás montado.", FontTypeNames.FONTTYPE_INFO)
136                         .flags.UltimoMensaje = 3

                        End If

                        '[/CDT]
                        Exit Sub

                    End If

138                 If .flags.Oculto = 1 Then

                        '[CDT 17-02-2004]
140                     If Not .flags.UltimoMensaje = 2 Then
142                         Call WriteLocaleMsg(Userindex, "55", FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "Ya estás oculto.", FontTypeNames.FONTTYPE_INFO)
144                         .flags.UltimoMensaje = 2

                        End If

                        '[/CDT]
                        Exit Sub

                    End If
                
146                 Call DoOcultarse(Userindex)

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

Private Sub HandleUseSpellMacro(ByVal Userindex As Integer)
        
        On Error GoTo HandleUseSpellMacro_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
104         Call SendData(SendTarget.ToAdmins, Userindex, PrepareMessageConsoleMsg(.name & " fue expulsado por Anti-macro de hechizos", FontTypeNames.FONTTYPE_VENENO))
106         Call WriteErrorMsg(Userindex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros")
        
108         Call CloseSocket(Userindex)

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

Private Sub HandleUseItem(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        On Error GoTo HandleUseItem_Err
        

100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim slot As Byte
        
108         slot = .incomingData.ReadByte()
        
110         If slot <= UserList(Userindex).CurrentInventorySlots And slot > 0 Then
112             If .Invent.Object(slot).ObjIndex = 0 Then Exit Sub

114             Call UseInvItem(Userindex, slot)

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

Private Sub HandleCraftBlacksmith(ByVal Userindex As Integer)
        
        On Error GoTo HandleCraftBlacksmith_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 3 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex).incomingData
            'Remove packet ID
106         Call .ReadByte
        
            Dim Item As Integer
        
108         Item = .ReadInteger()
        
110         If Item < 1 Then Exit Sub
        
            ' If ObjData(Item).SkHerreria = 0 Then Exit Sub
        
112         Call HerreroConstruirItem(Userindex, Item)

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

Private Sub HandleCraftCarpenter(ByVal Userindex As Integer)
        
        On Error GoTo HandleCraftCarpenter_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 3 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex).incomingData
            'Remove packet ID
106         Call .ReadByte
        
            Dim Item As Integer
        
108         Item = .ReadInteger()
        
110         If Item < 1 Then Exit Sub
        
            'If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        
112         Call CarpinteroConstruirItem(Userindex, Item)

        End With

        
        Exit Sub

HandleCraftCarpenter_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCraftCarpenter", Erl)
        Resume Next
        
End Sub

Private Sub HandleCraftAlquimia(ByVal Userindex As Integer)
        
        On Error GoTo HandleCraftAlquimia_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 4 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex).incomingData
            'Remove packet ID
106         Call .ReadInteger
        
            Dim Item As Integer
        
108         Item = .ReadInteger()
        
110         If Item < 1 Then Exit Sub

112         Call AlquimistaConstruirItem(Userindex, Item)

        End With

        
        Exit Sub

HandleCraftAlquimia_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCraftAlquimia", Erl)
        Resume Next
        
End Sub

Private Sub HandleCraftSastre(ByVal Userindex As Integer)
        
        On Error GoTo HandleCraftSastre_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 4 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex).incomingData
            'Remove packet ID
106         Call .ReadInteger
        
            Dim Item As Integer
        
108         Item = .ReadInteger()
        
110         If Item < 1 Then Exit Sub
            'If ObjData(Item).SkMAGOria = 0 Then Exit Sub

112         Call SastreConstruirItem(Userindex, Item)

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

Private Sub HandleWorkLeftClick(ByVal Userindex As Integer)
        
    On Error GoTo HandleWorkLeftClick_Err
        

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(Userindex)
            
        'Remove packet ID
        Call .incomingData.ReadByte
        
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

        If Not InRangoVision(Userindex, X, Y) Then
            Call WritePosUpdate(Userindex)
            Exit Sub
        End If
            
        If .flags.Meditando Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageMeditateToggle(.Char.CharIndex, 0))
        End If
        
        'If exiting, cancel
        Call CancelExit(Userindex)
        
        Select Case Skill

            Dim fallo As Boolean

            Case eSkill.Proyectiles
            
                'Check attack interval
                If Not IntervaloPermiteMagiaGolpe(Userindex, False) Then Exit Sub

                'Check Magic interval
                If Not IntervaloPermiteGolpeMagia(Userindex, False) Then Exit Sub

                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(Userindex) Then Exit Sub
                
                'Make sure the item is valid and there is ammo equipped.
                With .Invent

                    If .WeaponEqpObjIndex = 0 Then
                        DummyInt = 1
                    ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > UserList(Userindex).CurrentInventorySlots Then
                        DummyInt = 1
                    ElseIf .MunicionEqpSlot < 1 Or .MunicionEqpSlot > UserList(Userindex).CurrentInventorySlots Then
                        DummyInt = 1
                    ElseIf .MunicionEqpObjIndex = 0 Then
                        DummyInt = 1
                    ElseIf ObjData(.WeaponEqpObjIndex).proyectil <> 1 Then
                        DummyInt = 2
                    ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
                        DummyInt = 1
                    ElseIf .Object(.MunicionEqpSlot).Amount < 1 Then
                        DummyInt = 1

                    End If
                    
                    If DummyInt <> 0 Then
                        If DummyInt = 1 Then
                            Call WriteConsoleMsg(Userindex, "No tenés municiones.", FontTypeNames.FONTTYPE_INFO)
                            
                            'Call Desequipar(UserIndex, .WeaponEqpSlot)
                            Call WriteWorkRequestTarget(Userindex, 0)

                        End If
                        
                        Call Desequipar(Userindex, .MunicionEqpSlot)
                        Call WriteWorkRequestTarget(Userindex, 0)
                        Exit Sub

                    End If

                End With
                
                'Quitamos stamina
                If .Stats.MinSta >= 10 Then
                    Call QuitarSta(Userindex, RandomNumber(1, 10))
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageArmaMov(UserList(Userindex).Char.CharIndex))
                Else
                    Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
                    ' Call WriteConsoleMsg(UserIndex, "Estís muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteWorkRequestTarget(Userindex, 0)
                    Exit Sub

                End If
                
                Call LookatTile(Userindex, .Pos.Map, X, Y)
                
                tU = .flags.TargetUser
                tN = .flags.TargetNPC
                fallo = True

                'Validate target
                If tU > 0 Then

                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                        Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                        Call WriteWorkRequestTarget(Userindex, 0)
                        Exit Sub

                    End If
                    
                    'Prevent from hitting self
                    If tU = Userindex Then
                        Call WriteConsoleMsg(Userindex, "¡No podés atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(Userindex, 0)
                        Exit Sub

                    End If
                    
                    'Attack!
                    If Not PuedeAtacar(Userindex, tU) Then Exit Sub 'TODO: Por ahora pongo esto para solucionar lo anterior.
                    
                    Dim backup    As Byte
                    Dim envie     As Boolean
                    Dim Particula As Integer
                    Dim Tiempo    As Long
                    
                    Select Case ObjData(.Invent.MunicionEqpObjIndex).Subtipo

                        Case 1 'Paraliza
                            backup = UserList(Userindex).flags.Paraliza
                            UserList(Userindex).flags.Paraliza = 1

                        Case 2 ' Incinera
                            backup = UserList(Userindex).flags.incinera
                            UserList(Userindex).flags.incinera = 1

                        Case 3 ' envenena
                            backup = UserList(Userindex).flags.Envenena
                            UserList(Userindex).flags.Envenena = 1

                        Case 4 ' Explosiva

                    End Select

                    Call UsuarioAtacaUsuario(Userindex, tU)
                    
                    Select Case ObjData(.Invent.MunicionEqpObjIndex).Subtipo

                        Case 0

                        Case 1 'Paraliza
                            UserList(Userindex).flags.Paraliza = backup

                        Case 2 ' Incinera
                            UserList(Userindex).flags.incinera = backup

                        Case 3 ' envenena
                            UserList(Userindex).flags.Envenena = backup

                        Case 4 ' Explosiva

                    End Select
                    
                    If ObjData(.Invent.MunicionEqpObjIndex).CreaFX <> 0 Then
                        Call SendData(SendTarget.ToPCArea, tU, PrepareMessageCreateFX(UserList(tU).Char.CharIndex, ObjData(.Invent.MunicionEqpObjIndex).CreaFX, 0))

                    End If
                    
                    If ObjData(.Invent.MunicionEqpObjIndex).CreaParticula <> "" Then
                    
                        Particula = val(ReadField(1, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
                        Tiempo = val(ReadField(2, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
                        Call SendData(SendTarget.ToPCArea, tU, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, Particula, Tiempo, False))

                    End If
                    
                    fallo = False
                    
                ElseIf tN > 0 Then

                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(Npclist(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(Npclist(tN).Pos.X - .Pos.X) > RANGO_VISION_X Then
                        Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(Userindex, 0)
                        'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)
                        Exit Sub

                    End If
                    
                    'Is it attackable???
                    If Npclist(tN).Attackable <> 0 Then
                    
                        fallo = False
                        
                        'Attack!
                        
                        Select Case UsuarioAtacaNpcFunction(Userindex, tN)
                        
                            Case 0 ' no se puede pegar
                            
                            Case 1 ' le pego
                
                                If Npclist(tN).flags.Snd2 > 0 Then
                                    Call SendData(SendTarget.ToNPCArea, tN, PrepareMessagePlayWave(Npclist(tN).flags.Snd2, Npclist(tN).Pos.X, Npclist(tN).Pos.Y))
                                Else
                                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(tN).Pos.X, Npclist(tN).Pos.Y))

                                End If
                                
                                If ObjData(.Invent.MunicionEqpObjIndex).Subtipo = 1 And UserList(Userindex).flags.TargetNPC > 0 Then
                                    If Npclist(tN).flags.Paralizado = 0 Then

                                        Dim Probabilidad As Byte

                                        Probabilidad = RandomNumber(1, 2)

                                        If Probabilidad = 1 Then
                                            If Npclist(tN).flags.AfectaParalisis = 0 Then
                                                Npclist(tN).flags.Paralizado = 1
                                                
                                                Npclist(tN).Contadores.Paralisis = IntervaloParalizado

                                                If UserList(Userindex).ChatCombate = 1 Then
                                                    'Call WriteConsoleMsg(UserIndex, "Tu golpe a paralizado a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
                                                    Call WriteLocaleMsg(Userindex, "136", FontTypeNames.FONTTYPE_FIGHT)

                                                End If

                                                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(Npclist(tN).Char.CharIndex, 8, 0))
                                                envie = True
                                            Else

                                                If UserList(Userindex).ChatCombate = 1 Then
                                                    'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
                                                    Call WriteLocaleMsg(Userindex, "381", FontTypeNames.FONTTYPE_INFO)

                                                End If

                                            End If

                                        End If

                                    End If

                                End If
                                
                                If ObjData(.Invent.MunicionEqpObjIndex).CreaFX <> 0 Then
                                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(Npclist(tN).Char.CharIndex, ObjData(.Invent.MunicionEqpObjIndex).CreaFX, 0))

                                End If
                    
                                If ObjData(.Invent.MunicionEqpObjIndex).CreaParticula <> "" Then
                                    Particula = val(ReadField(1, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
                                    Tiempo = val(ReadField(2, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
                                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(Npclist(tN).Char.CharIndex, Particula, Tiempo, False))

                                End If
                                
                            Case 2 ' Fallo
                            
                        End Select
                        
                    End If

                End If
                
                With .Invent
                    DummyInt = .MunicionEqpSlot
                    
                    'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                    
                    If Not fallo Then
                        Call QuitarUserInvItem(Userindex, DummyInt, 1)

                    End If
                    
                    'Call DropObj(UserIndex, .MunicionEqpSlot, 1, UserList(UserIndex).Pos.Map, x, y)
                   
                    ' If fallo And MapData(UserList(UserIndex).Pos.Map, x, Y).Blocked = 0 Then
                    '  Dim flecha As obj
                    ' flecha.Amount = 1
                    'flecha.ObjIndex = .MunicionEqpObjIndex
                    ' Call MakeObj(flecha, UserList(UserIndex).Pos.Map, x, Y)
                    ' End If
                    
                    If .Object(DummyInt).Amount > 0 Then
                        'QuitarUserInvItem unequipps the ammo, so we equip it again
                        .MunicionEqpSlot = DummyInt
                        .MunicionEqpObjIndex = .Object(DummyInt).ObjIndex
                        .Object(DummyInt).Equipped = 1
                    Else
                        .MunicionEqpSlot = 0
                        .MunicionEqpObjIndex = 0

                    End If

                    Call UpdateUserInv(False, Userindex, DummyInt)

                End With

                '-----------------------------------
            
            Case eSkill.magia
                'Check the map allows spells to be casted.
                '  If MapInfo(.Pos.map).MagiaSinEfecto > 0 Then
                ' Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energía", FontTypeNames.FONTTYPE_FIGHT)
                '  Exit Sub
                ' End If
                
                'Target whatever is in that tile
                Call LookatTile(Userindex, .Pos.Map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .ip & " a la posicion (" & .Pos.Map & "/" & X & "/" & Y & ")")
                    Exit Sub

                End If
                
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(Userindex, False) Then Exit Sub
                
                'Check attack-spell interval
                If Not IntervaloPermiteGolpeMagia(Userindex, False) Then Exit Sub
                
                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(Userindex) Then Exit Sub
                
                'Check intervals and cast
                If .flags.Hechizo > 0 Then
                    Call LanzarHechizo(.flags.Hechizo, Userindex)
                    .flags.Hechizo = 0
                Else
                    Call WriteConsoleMsg(Userindex, "¡Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eSkill.Pescar
                
                If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                    
                If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> eOBJType.otHerramientas Then Exit Sub
                    
                'Check interval
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub

                Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                    Case 1      ' Subtipo: Caña de Pescar

                        If (MapData(.Pos.Map, X, Y).Blocked And FLAG_AGUA) <> 0 Then
                            Call DoPescar(Userindex, False, ObjData(.Invent.HerramientaEqpObjIndex).Dorada = 1)
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
                            
                        Else
                            Call WriteConsoleMsg(Userindex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteMacroTrabajoToggle(Userindex, False)
    
                        End If
                    
                    Case 2      ' Subtipo: Red de Pesca
    
                        If (MapData(.Pos.Map, X, Y).Blocked And FLAG_AGUA) <> 0 Then
                            
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 8 Then
                                Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos para pescar.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)
                                Exit Sub
    
                            End If
                                
                            If UserList(Userindex).Stats.UserSkills(eSkill.Pescar) < 80 Then
                                Call WriteConsoleMsg(Userindex, "Para utilizar la red de pesca debes tener 80 skills en recoleccion.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)
                                Exit Sub
    
                            End If
                                    
                            If MapInfo(UserList(Userindex).Pos.Map).Seguro = 1 Then
                                Call WriteConsoleMsg(Userindex, "Esta prohibida la pesca masiva en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)
                                Exit Sub
    
                            End If
                                    
                            If UserList(Userindex).flags.Navegando = 0 Then
                                Call WriteConsoleMsg(Userindex, "Necesitas estar sobre tu barca para utilizar la red de pesca.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)
                                Exit Sub
    
                            End If
                                    
                            Call DoPescar(Userindex, True, True)
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
                        
                        Else
                        
                            Call WriteConsoleMsg(Userindex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(Userindex, 0)
    
                        End If
                
                End Select
                
                    
            Case eSkill.Talar
            
                If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub

                If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> eOBJType.otHerramientas Then Exit Sub
        
                'Check interval
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub

                Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                    Case 6      ' Herramientas de Carpinteria - Hacha
                        
                        'Target whatever is in the tile
                        Call LookatTile(Userindex, .Pos.Map, X, Y)

                        ' Ahora se puede talar en la ciudad
                        'If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
                        '    Call WriteConsoleMsg(UserIndex, "Esta prohibido talar arboles en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                        '    Call WriteWorkRequestTarget(UserIndex, 0)
                        '    Exit Sub
                        'End If
                            
                        DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                            
                        If DummyInt > 0 Then
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 1 Then
                                Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                                'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)
                                Exit Sub

                            End If
                                
                            If .Pos.X = X And .Pos.Y = Y Then
                                Call WriteConsoleMsg(Userindex, "No podés talar desde allí.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)
                                Exit Sub

                            End If
                                
                            If MapData(.Pos.Map, X, Y).ObjInfo.Amount <= 0 Then
                                Call WriteConsoleMsg(Userindex, "El árbol ya no te puede entregar mas leña.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)
                                Call WriteMacroTrabajoToggle(Userindex, False)
                                Exit Sub

                            End If

                            '¡Hay un arbol donde clickeo?
                            If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
                                Call DoTalar(Userindex, X, Y, ObjData(.Invent.HerramientaEqpObjIndex).Dorada = 1)
                            End If

                        Else
                            
                            Call WriteConsoleMsg(Userindex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(Userindex, 0)

                            If UserList(Userindex).Counters.Trabajando > 1 Then
                                Call WriteMacroTrabajoToggle(Userindex, False)

                            End If

                        End If
                
                End Select
            
            Case eSkill.Alquimia
            
                If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                    
                If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> eOBJType.otHerramientas Then Exit Sub
                    
                'Check interval
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub

                Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                    Case 3  ' Herramientas de Alquimia - Tijeras

                        If MapInfo(UserList(Userindex).Pos.Map).Seguro = 1 Then
                            Call WriteWorkRequestTarget(Userindex, 0)
                            Call WriteConsoleMsg(Userindex, "Esta prohibido cortar raices en las ciudades.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                            
                        If MapData(.Pos.Map, X, Y).ObjInfo.Amount <= 0 Then
                            Call WriteConsoleMsg(Userindex, "El árbol ya no te puede entregar mas raices.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(Userindex, 0)
                            Call WriteMacroTrabajoToggle(Userindex, False)
                            Exit Sub

                        End If
                
                        DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                            
                        If DummyInt > 0 Then
                            
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                                'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)
                                Exit Sub

                            End If
                                
                            If .Pos.X = X And .Pos.Y = Y Then
                                Call WriteConsoleMsg(Userindex, "No podés quitar raices allí.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)
                                Exit Sub

                            End If
                                
                            '¡Hay un arbol donde clickeo?
                            If ObjData(DummyInt).OBJType = eOBJType.otArboles Then
                                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_TIJERAS, .Pos.X, .Pos.Y))
                                Call DoRaices(Userindex, X, Y)

                            End If

                        Else
                            Call WriteConsoleMsg(Userindex, "No hay ningún árbol ahí.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(Userindex, 0)
                            Call WriteMacroTrabajoToggle(Userindex, False)

                        End If
                
                End Select
                
            Case eSkill.Mineria
            
                If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                    
                If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> eOBJType.otHerramientas Then Exit Sub
                    
                'Check interval
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub

                Select Case .Invent.HerramientaEqpObjIndex
                
                    Case 8  ' Herramientas de Mineria - Piquete
                
                        'Target whatever is in the tile
                        Call LookatTile(Userindex, .Pos.Map, X, Y)
                            
                        DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                            
                        If DummyInt > 0 Then

                            'Check distance
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)
                                Exit Sub

                            End If
                                
                            If MapData(.Pos.Map, X, Y).ObjInfo.Amount <= 0 Then
                                Call WriteConsoleMsg(Userindex, "Este yacimiento no tiene mas minerales para entregar.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)
                                Call WriteMacroTrabajoToggle(Userindex, False)
                                Exit Sub

                            End If
                                
                            DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex 'CHECK

                            '¡Hay un yacimiento donde clickeo?
                            If ObjData(DummyInt).OBJType = eOBJType.otYacimiento Then
                                Call DoMineria(Userindex, X, Y, ObjData(.Invent.HerramientaEqpObjIndex).Dorada = 1)

                            Else
                                Call WriteConsoleMsg(Userindex, "Ahí no hay ningún yacimiento.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)

                            End If

                        Else
                            Call WriteConsoleMsg(Userindex, "Ahí no hay ningun yacimiento.", FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(Userindex, 0)

                        End If

                End Select

            Case eSkill.Robar

                'Does the map allow us to steal here?
                If MapInfo(.Pos.Map).Seguro = 0 Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call LookatTile(Userindex, UserList(Userindex).Pos.Map, X, Y)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> Userindex Then

                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.user Then
                            If UserList(tU).flags.Muerto = 0 Then
                                If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                    Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                    Call WriteWorkRequestTarget(Userindex, 0)
                                    Exit Sub

                                End If
                                 
                                '17/09/02
                                'Check the trigger
                                If MapData(UserList(tU).Pos.Map, X, Y).trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(Userindex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                    Call WriteWorkRequestTarget(Userindex, 0)
                                    Exit Sub

                                End If
                                 
                                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(Userindex, "No podés robar aquí.", FontTypeNames.FONTTYPE_WARNING)
                                    Call WriteWorkRequestTarget(Userindex, 0)
                                    Exit Sub

                                End If
                                 
                                Call DoRobar(Userindex, tU)

                            End If

                        End If

                    Else
                        Call WriteConsoleMsg(Userindex, "No a quien robarle!", FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(Userindex, 0)

                    End If

                Else
                    Call WriteConsoleMsg(Userindex, "¡No podés robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)
                    Call WriteWorkRequestTarget(Userindex, 0)

                End If
                    
            Case eSkill.Domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de
                'criaturas hostiles.
                    
                'Target whatever is that tile
                Call LookatTile(Userindex, .Pos.Map, X, Y)
                tN = .flags.TargetNPC
                    
                If tN > 0 Then
                    If Npclist(tN).flags.Domable > 0 Then
                        If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                            Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
    
                        End If
                            
                        If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                            Call WriteConsoleMsg(Userindex, "No puedes domar una criatura que esta luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
    
                        End If
                            
                        Call DoDomar(Userindex, tN)
                    Else
                        Call WriteConsoleMsg(Userindex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)
    
                    End If
    
                Else
                    Call WriteConsoleMsg(Userindex, "No hay ninguna criatura alli!", FontTypeNames.FONTTYPE_INFO)
    
                End If
               
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
            
                'Check interval
                If Not IntervaloPermiteTrabajar(Userindex) Then Exit Sub
                
                Call LookatTile(Userindex, .Pos.Map, X, Y)
                
                'Check there is a proper item there
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then

                        'Validate other items
                        If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > UserList(Userindex).CurrentInventorySlots Then
                            Exit Sub

                        End If
                        
                        ''chequeamos que no se zarpe duplicando oro
                        If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                            If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                                Call WriteConsoleMsg(Userindex, "No tienes más minerales", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)
                                Exit Sub

                            End If
                            
                            ''FUISTE
                            Call WriteErrorMsg(Userindex, "Has sido expulsado por el sistema anti cheats.")
                            
                            Call CloseSocket(Userindex)
                            Exit Sub

                        End If
                        
                        Call FundirMineral(Userindex)
                        
                    Else
                    
                        Call WriteConsoleMsg(Userindex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(Userindex, 0)

                        If UserList(Userindex).Counters.Trabajando > 1 Then
                            Call WriteMacroTrabajoToggle(Userindex, False)

                        End If

                    End If

                Else
                
                    Call WriteConsoleMsg(Userindex, "Ahí no hay ninguna fragua.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteWorkRequestTarget(Userindex, 0)

                    If UserList(Userindex).Counters.Trabajando > 1 Then
                        Call WriteMacroTrabajoToggle(Userindex, False)

                    End If

                End If

            Case eSkill.Grupo
                'If UserList(UserIndex).Grupo.EnGrupo = False Then
                'Target whatever is in that tile
                Call LookatTile(Userindex, UserList(Userindex).Pos.Map, X, Y)
                    
                tU = .flags.TargetUser
                    
                'Call WritePreguntaBox(UserIndex, UserList(UserIndex).name & " te invitó a unirte a su grupo. ¿Deseas unirte?")
                    
                If tU > 0 And tU <> Userindex Then

                    'Can't steal administrative players
                    If UserList(Userindex).Grupo.EnGrupo = False Then
                        If UserList(tU).flags.Muerto = 0 Then
                            If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 8 Then
                                Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(Userindex, 0)
                                Exit Sub

                            End If
                                         
                            If UserList(Userindex).Grupo.CantidadMiembros = 0 Then
                                UserList(Userindex).Grupo.Lider = Userindex
                                UserList(Userindex).Grupo.Miembros(1) = Userindex
                                UserList(Userindex).Grupo.CantidadMiembros = 1
                                Call InvitarMiembro(Userindex, tU)
                            Else
                                UserList(Userindex).Grupo.Lider = Userindex
                                Call InvitarMiembro(Userindex, tU)

                            End If
                                         
                        Else
                            Call WriteLocaleMsg(Userindex, "7", FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "El usuario esta muerto.", FontTypeNames.FONTTYPE_INFOIAO)
                            Call WriteWorkRequestTarget(Userindex, 0)

                        End If

                    Else

                        If UserList(Userindex).Grupo.Lider = Userindex Then
                            Call InvitarMiembro(Userindex, tU)
                        Else
                            Call WriteConsoleMsg(Userindex, "Tu no podés invitar usuarios, debe hacerlo " & UserList(UserList(Userindex).Grupo.Lider).name & ".", FontTypeNames.FONTTYPE_INFOIAO)
                            Call WriteWorkRequestTarget(Userindex, 0)

                        End If

                    End If

                Else
                    Call WriteLocaleMsg(Userindex, "261", FontTypeNames.FONTTYPE_INFO)

                End If

                ' End If
            Case eSkill.MarcaDeClan

                'If UserList(UserIndex).Grupo.EnGrupo = False Then
                'Target whatever is in that tile
                Dim clan_nivel As Byte
                
                If UserList(Userindex).GuildIndex = 0 Then
                    Call WriteConsoleMsg(Userindex, "Servidor> No Perteneces a ningun clan.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If
                
                clan_nivel = modGuilds.NivelDeClan(UserList(Userindex).GuildIndex)

                If clan_nivel < 4 Then
                    Call WriteConsoleMsg(Userindex, "Servidor> El nivel de tu clan debe ser 4 para utilizar esta opción.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If
                                
                Call LookatTile(Userindex, UserList(Userindex).Pos.Map, X, Y)
                    
                tU = .flags.TargetUser

                If tU = 0 Then Exit Sub
                    
                If UserList(Userindex).GuildIndex = UserList(tU).GuildIndex Then
                    Call WriteConsoleMsg(Userindex, "Servidor> No podes marcar a un miembro de tu clan.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If
                    
                'Call WritePreguntaBox(UserIndex, UserList(UserIndex).name & " te invitó a unirte a su grupo. ¿Deseas unirte?")
                    
                If tU > 0 And tU <> Userindex Then

                    'Can't steal administrative players
                    If UserList(tU).flags.Muerto = 0 Then
                        'call marcar
                        Call SendData(SendTarget.ToClanArea, Userindex, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, 210, 700, False))
                        Call SendData(SendTarget.ToClanArea, Userindex, PrepareMessageConsoleMsg("Clan> [" & UserList(Userindex).name & "] marco a " & UserList(tU).name & ".", FontTypeNames.FONTTYPE_GUILD))
                    Else
                        Call WriteLocaleMsg(Userindex, "7", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "El usuario esta muerto.", FontTypeNames.FONTTYPE_INFOIAO)
                        Call WriteWorkRequestTarget(Userindex, 0)

                    End If

                Else
                    Call WriteLocaleMsg(Userindex, "261", FontTypeNames.FONTTYPE_INFO)

                End If

            Case eSkill.MarcaDeGM
                Call LookatTile(Userindex, UserList(Userindex).Pos.Map, X, Y)
                    
                tU = .flags.TargetUser

                If tU > 0 Then
                    Call WriteConsoleMsg(Userindex, "Servidor> [" & UserList(tU).name & "] seleccionado.", FontTypeNames.FONTTYPE_SERVER)
                Else
                    Call WriteLocaleMsg(Userindex, "261", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleCreateNewGuild(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
        
        If modGuilds.CrearNuevoClan(Userindex, Desc, GuildName, Alineacion, errorStr) Then

            Call QuitarObjetos(407, 1, Userindex)
            Call QuitarObjetos(408, 1, Userindex)
            Call QuitarObjetos(409, 1, Userindex)
            Call QuitarObjetos(411, 1, Userindex)
            
            Call SendData(SendTarget.ToAll, Userindex, PrepareMessageConsoleMsg(.name & " fundó el clan " & GuildName & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
            'Update tag
            Call RefreshCharStatus(Userindex)
        Else
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleSpellInfo(ByVal Userindex As Integer)
        
        On Error GoTo HandleSpellInfo_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim spellSlot As Byte

            Dim Spell     As Integer
        
108         spellSlot = .incomingData.ReadByte()
        
            'Validate slot
110         If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
112             Call WriteConsoleMsg(Userindex, "¡Primero selecciona el hechizo!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate spell in the slot
114         Spell = .Stats.UserHechizos(spellSlot)

116         If Spell > 0 And Spell < NumeroHechizos + 1 Then

118             With Hechizos(Spell)
                    'Send information
120                 Call WriteConsoleMsg(Userindex, "HECINF*" & Spell, FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleEquipItem(ByVal Userindex As Integer)
        
        On Error GoTo HandleEquipItem_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim itemSlot As Byte
        
108         itemSlot = .incomingData.ReadByte()
        
            'Dead users can't equip items
110         If .flags.Muerto = 1 Then
112             Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!! Sólo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate item slot
114         If itemSlot > UserList(Userindex).CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
116         If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
118         Call EquiparInvItem(Userindex, itemSlot)

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

Private Sub HandleChangeHeading(ByVal Userindex As Integer)
        
        On Error GoTo HandleChangeHeading_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 10/01/08
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Heading As eHeading
        
108         Heading = .incomingData.ReadByte()
        
            'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
110         If Heading > 0 And Heading < 5 Then
112             .Char.Heading = Heading
114             Call ChangeUserChar(Userindex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

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

Private Sub HandleModifySkills(ByVal Userindex As Integer)
        
        On Error GoTo HandleModifySkills_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 1 + NUMSKILLS Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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
118                 Call CloseSocket(Userindex)
                    Exit Sub

                End If
            
120             Count = Count + points(i)
122         Next i
        
124         If Count > .Stats.SkillPts Then
126             Call LogHackAttemp(.name & " IP:" & .ip & " trató de hackear los skills.")
128             Call CloseSocket(Userindex)
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

Private Sub HandleTrain(ByVal Userindex As Integer)
        
        On Error GoTo HandleTrain_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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
126             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))

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

Private Sub HandleCommerceBuy(ByVal Userindex As Integer)
        
        On Error GoTo HandleCommerceBuy_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 4 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim slot   As Byte

            Dim Amount As Integer
        
108         slot = .incomingData.ReadByte()
110         Amount = .incomingData.ReadInteger()
        
            'Dead people can't commerce...
112         If .flags.Muerto = 1 Then
114             Call WriteConsoleMsg(Userindex, "¡¡Estís muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'íEl target es un NPC valido?
116         If .flags.TargetNPC < 1 Then Exit Sub
            
            'íEl NPC puede comerciar?
118         If Npclist(.flags.TargetNPC).Comercia = 0 Then
120             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
                Exit Sub

            End If
        
            'Only if in commerce mode....
122         If Not .flags.Comerciando Then
124             Call WriteConsoleMsg(Userindex, "No estás comerciando", FontTypeNames.FONTTYPE_INFO)
126             Call WriteCommerceEnd(Userindex)
                Exit Sub

            End If
        
            'User compra el item
128         Call Comercio(eModoComercio.Compra, Userindex, .flags.TargetNPC, slot, Amount)

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

Private Sub HandleBankExtractItem(ByVal Userindex As Integer)
        
        On Error GoTo HandleBankExtractItem_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 5 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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
116             Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            '¿El target es un NPC valido?
118         If .flags.TargetNPC < 1 Then Exit Sub
        
            '¿Es el banquero?
120         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
                Exit Sub

            End If
        
            'User retira el item del slot
122         Call UserRetiraItem(Userindex, slot, Amount, slotdestino)

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

Private Sub HandleCommerceSell(ByVal Userindex As Integer)
        
        On Error GoTo HandleCommerceSell_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 4 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim slot   As Byte

            Dim Amount As Integer
        
108         slot = .incomingData.ReadByte()
110         Amount = .incomingData.ReadInteger()
        
            'Dead people can't commerce...
112         If .flags.Muerto = 1 Then
114             Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'íEl target es un NPC valido?
116         If .flags.TargetNPC < 1 Then Exit Sub
        
            'íEl NPC puede comerciar?
118         If Npclist(.flags.TargetNPC).Comercia = 0 Then
120             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
                Exit Sub

            End If
        
            'User compra el item del slot
122         Call Comercio(eModoComercio.Venta, Userindex, .flags.TargetNPC, slot, Amount)

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

Private Sub HandleBankDeposit(ByVal Userindex As Integer)
        
        On Error GoTo HandleBankDeposit_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 5 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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
116             Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'íEl target es un NPC valido?
118         If .flags.TargetNPC < 1 Then Exit Sub
        
            'íEl NPC puede comerciar?
120         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
                Exit Sub

            End If
        
            'User deposita el item del slot rdata
122         Call UserDepositaItem(Userindex, slot, Amount, slotdestino)

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

Private Sub HandleForumPost(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
    
ErrHandler:

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

Private Sub HandleMoveSpell(ByVal Userindex As Integer)
        
        On Error GoTo HandleMoveSpell_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 3 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex).incomingData
            'Remove packet ID
106         Call .ReadByte
        
            Dim dir As Integer
        
108         If .ReadBoolean() Then
110             dir = 1
            Else
112             dir = -1

            End If
        
114         Call DesplazarHechizo(Userindex, dir, .ReadByte())

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

Private Sub HandleClanCodexUpdate(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
    
ErrHandler:

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

Private Sub HandleUserCommerceOffer(ByVal Userindex As Integer)
        
        On Error GoTo HandleUserCommerceOffer_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 6 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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
114         If ((slot < 1 Or slot > UserList(Userindex).CurrentInventorySlots) And slot <> FLAGORO) Or Amount <= 0 Then Exit Sub
        
            'Is the other player valid??
116         If tUser < 1 Or tUser > MaxUsers Then Exit Sub
        
            'Is the commerce attempt valid??
118         If UserList(tUser).ComUsu.DestUsu <> Userindex Then
120             Call FinComerciarUsu(Userindex)
                Exit Sub

            End If
        
            'Is he still logged??
122         If Not UserList(tUser).flags.UserLogged Then
124             Call FinComerciarUsu(Userindex)
                Exit Sub
            Else

                'Is he alive??
126             If UserList(tUser).flags.Muerto = 1 Then
128                 Call FinComerciarUsu(Userindex)
                    Exit Sub

                End If
            
                'Has he got enough??
130             If slot = FLAGORO Then

                    'gold
132                 If Amount > .Stats.GLD Then
134                     Call WriteConsoleMsg(Userindex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                        Exit Sub

                    End If

                Else

                    'inventory
136                 If Amount > .Invent.Object(slot).Amount Then
138                     Call WriteConsoleMsg(Userindex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                        Exit Sub

                    End If

                End If
            
                'Prevent offer changes (otherwise people would ripp off other players)
140             If .ComUsu.Objeto > 0 Then
142                 Call WriteConsoleMsg(Userindex, "No podés cambiar tu oferta.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub

                End If
            
                'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
144             If .flags.Navegando = 1 Then
146                 If .Invent.BarcoSlot = slot Then
148                     Call WriteConsoleMsg(Userindex, "No podés vender tu barco mientras lo estás usando.", FontTypeNames.FONTTYPE_TALK)
                        Exit Sub

                    End If

                End If
            
150             If .flags.Montado = 1 Then
152                 If .Invent.MonturaSlot = slot Then
154                     Call WriteConsoleMsg(Userindex, "No podés vender tu montura mientras la estás usando.", FontTypeNames.FONTTYPE_TALK)
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

Private Sub HandleGuildAcceptPeace(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(Userindex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildRejectAlliance(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(Userindex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildRejectPeace(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(Userindex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildAcceptAlliance(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        guild = buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(Userindex, guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildOfferPeace(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
        
        If modGuilds.r_ClanGeneraPropuesta(Userindex, guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(Userindex, "Propuesta de paz enviada", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildOfferAlliance(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
        
        If modGuilds.r_ClanGeneraPropuesta(Userindex, guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(Userindex, "Propuesta de alianza enviada", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildAllianceDetails(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild    As String

        Dim errorStr As String

        Dim details  As String
        
        guild = buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(Userindex, guild, RELACIONES_GUILD.ALIADOS, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(Userindex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildPeaceDetails(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild    As String

        Dim errorStr As String

        Dim details  As String
        
        guild = buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(Userindex, guild, RELACIONES_GUILD.PAZ, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(Userindex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildRequestJoinerInfo(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim user    As String

        Dim details As String
        
        user = buffer.ReadASCIIString()
        
        details = modGuilds.a_DetallesAspirante(Userindex, user)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(Userindex, "El personaje no ha mandado solicitud, o no estás habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(Userindex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildAlliancePropList(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleGuildAlliancePropList_Err
        
100     Call UserList(Userindex).incomingData.ReadByte
    
102     Call WriteAlianceProposalsList(Userindex, r_ListaDePropuestas(Userindex, RELACIONES_GUILD.ALIADOS))

        
        Exit Sub

HandleGuildAlliancePropList_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleGuildAlliancePropList", Erl)
        Resume Next
        
End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleGuildPeacePropList_Err
        
100     Call UserList(Userindex).incomingData.ReadByte
    
102     Call WritePeaceProposalsList(Userindex, r_ListaDePropuestas(Userindex, RELACIONES_GUILD.PAZ))

        
        Exit Sub

HandleGuildPeacePropList_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleGuildPeacePropList", Erl)
        Resume Next
        
End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild           As String

        Dim errorStr        As String

        Dim otherGuildIndex As Integer
        
        guild = buffer.ReadASCIIString()
        
        otherGuildIndex = modGuilds.r_DeclararGuerra(Userindex, guild, errorStr)
        
        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
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
    
ErrHandler:

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

Private Sub HandleGuildNewWebsite(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarWebSite(Userindex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildAcceptNewMember(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim errorStr As String

        Dim UserName As String

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        
        If Not modGuilds.a_AceptarAspirante(Userindex, UserName, errorStr) Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
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
    
ErrHandler:

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

Private Sub HandleGuildRejectNewMember(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
        
        If Not modGuilds.a_RechazarAspirante(Userindex, UserName, errorStr) Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
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
    
ErrHandler:

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

Private Sub HandleGuildKickMember(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName   As String

        Dim GuildIndex As Integer
        
        UserName = buffer.ReadASCIIString()
        
        GuildIndex = modGuilds.m_EcharMiembroDeClan(Userindex, UserName)
        
        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        Else
            Call WriteConsoleMsg(Userindex, "No podés expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildUpdateNews(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.ActualizarNoticias(Userindex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildMemberInfo(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendDetallesPersonaje(Userindex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildOpenElections(ByVal Userindex As Integer)
        
        On Error GoTo HandleGuildOpenElections_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            Dim Error As String
        
104         If Not modGuilds.v_AbrirElecciones(Userindex, Error) Then
106             Call WriteConsoleMsg(Userindex, Error, FontTypeNames.FONTTYPE_GUILD)
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

Private Sub HandleGuildRequestMembership(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
        
        If Not modGuilds.a_NuevoAspirante(Userindex, guild, application, errorStr) Then
            Call WriteConsoleMsg(Userindex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, "Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & guild & ".", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildRequestDetails(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Call modGuilds.SendGuildDetails(Userindex, buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleOnline(ByVal Userindex As Integer)
        
        On Error GoTo HandleOnline_Err
        

        '***************************************************
        Dim i     As Long

        Dim Count As Long
    
100     With UserList(Userindex)
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
112             Call WriteConsoleMsg(Userindex, "Número de usuarios: " & CStr(Count) & " conectados.", FontTypeNames.FONTTYPE_INFOIAO)
            Else
                Call WriteConsoleMsg(Userindex, "Número de usuarios: " & CStr(Count) & " conectados: " & nombres & ".", FontTypeNames.FONTTYPE_INFOIAO)
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

Private Sub HandleQuit(ByVal Userindex As Integer)
        
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
    
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(Userindex, "No podés salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub

        End If
        
        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = Userindex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                    Call FinComerciarUsu(tUser)

                End If

            End If
            
            Call WriteConsoleMsg(Userindex, "Comercio cancelado. ", FontTypeNames.FONTTYPE_TALK)
            Call FinComerciarUsu(Userindex)

        End If
        
        isNotVisible = (.flags.Oculto Or .flags.invisible)

        If isNotVisible Then
            .flags.Oculto = 0
            .flags.invisible = 0

            .Counters.Invisibilidad = 0
            .Counters.TiempoOculto = 0
                
            'Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(Userindex, "307", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(.Char.CharIndex, False))

        End If
        
        Rem   Call WritePersonajesDeCuenta(UserIndex, .Cuenta)
        Call Cerrar_Usuario(Userindex)

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

Private Sub HandleGuildLeave(ByVal Userindex As Integer)
        
        On Error GoTo HandleGuildLeave_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim GuildIndex As Integer
    
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'obtengo el guildindex
104         GuildIndex = m_EcharMiembroDeClan(Userindex, .name)
        
106         If GuildIndex > 0 Then
108             Call WriteConsoleMsg(Userindex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
110             Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
            Else
112             Call WriteConsoleMsg(Userindex, "Tu no podés salir de ningún clan.", FontTypeNames.FONTTYPE_GUILD)

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

Private Sub HandleRequestAccountState(ByVal Userindex As Integer)
        
        On Error GoTo HandleRequestAccountState_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim earnings   As Integer

        Dim percentage As Integer
    
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Dead people can't check their accounts
104         If .flags.Muerto = 1 Then
106             Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
108         If .flags.TargetNPC = 0 Then
110             Call WriteConsoleMsg(Userindex, "Primero tenes que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
112         If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
114             Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
116         Select Case Npclist(.flags.TargetNPC).NPCtype

                Case eNPCType.Banquero
118                 Call WriteChatOverHead(Userindex, "Tenes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
120             Case eNPCType.Timbero

122                 If Not .flags.Privilegios And PlayerType.user Then
124                     earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
126                     If earnings >= 0 And Apuestas.Ganancias <> 0 Then
128                         percentage = Int(earnings * 100 / Apuestas.Ganancias)

                        End If
                    
130                     If earnings < 0 And Apuestas.Perdidas <> 0 Then
132                         percentage = Int(earnings * 100 / Apuestas.Perdidas)

                        End If
                    
134                     Call WriteConsoleMsg(Userindex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)

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
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's his pet
        If Npclist(.flags.TargetNPC).MaestroUser <> Userindex Then Exit Sub
        
        'Do it!
        Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO
        
        Call Expresar(.flags.TargetNPC, Userindex)
    End With
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal Userindex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> Userindex Then Exit Sub
        
        'Do it
        Call FollowAmo(.flags.TargetNPC)
        
        Call Expresar(.flags.TargetNPC, Userindex)
    End With
End Sub

''
' Handles the "PetLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetLeave(ByVal Userindex As Integer)
'***************************************************
    With UserList(Userindex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> Userindex Then Exit Sub

        Call QuitarNPC(.flags.TargetNPC)
    End With
End Sub

''
' Handles the "GrupoMsg" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGrupoMsg(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                    Call WriteChatOverHead(UserList(.Grupo.Lider).Grupo.Miembros(i), chat, UserList(Userindex).Char.CharIndex, &HFF8000)
                  
                Next i
            
            Else
                'Call WriteConsoleMsg(UserIndex, "[" & .Name & "] " & chat, FontTypeNames.FONTTYPE_New_GRUPO)
                Call WriteConsoleMsg(Userindex, "Grupo> No estas en ningun grupo.", FontTypeNames.FONTTYPE_New_GRUPO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleTrainList(ByVal Userindex As Integer)
        
        On Error GoTo HandleTrainList_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Dead users can't use pets
104         If .flags.Muerto = 1 Then
106             Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
108         If .flags.TargetNPC = 0 Then
110             Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make sure it's close enough
112         If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
114             Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make sure it's the trainer
116         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
118         Call WriteTrainerCreatureList(Userindex, .flags.TargetNPC)

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

Private Sub HandleRest(ByVal Userindex As Integer)
        
        On Error GoTo HandleRest_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Dead users can't use pets
104         If .flags.Muerto = 1 Then
106             Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!! Solo podés usar items cuando estás vivo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If HayOBJarea(.Pos, FOGATA) Then
110             Call WriteRestOK(Userindex)
            
112             If Not .flags.Descansar Then
114                 Call WriteConsoleMsg(Userindex, "Te acomodás junto a la fogata y comenzís a descansar.", FontTypeNames.FONTTYPE_INFO)
                Else
116                 Call WriteConsoleMsg(Userindex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)

                End If
            
118             .flags.Descansar = Not .flags.Descansar
            Else

120             If .flags.Descansar Then
122                 Call WriteRestOK(Userindex)
124                 Call WriteConsoleMsg(Userindex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
126                 .flags.Descansar = False
                    Exit Sub

                End If
            
128             Call WriteConsoleMsg(Userindex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleMeditate(ByVal Userindex As Integer)
        
        On Error GoTo HandleMeditate_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 04/15/08 (NicoNZ)
        'Arreglí un bug que mandaba un index de la meditacion diferente
        'al que decia el server.
        '***************************************************
        
100     With UserList(Userindex)

            'Remove packet ID
102         Call .incomingData.ReadByte
            
            'Si ya tiene el mana completo, no lo dejamos meditar.
104         If .Stats.MinMAN = .Stats.MaxMAN Then Exit Sub
                           
            'Las clases NO MAGICAS no meditan...
106         If .clase = eClass.Hunter Or _
               .clase = eClass.Trabajador Or _
               .clase = eClass.Warrior Then Exit Sub

108         If .flags.Muerto = 1 Then
110             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
112         If .flags.Montado = 1 Then
114             Call WriteConsoleMsg(Userindex, "No podes meditar estando montado.", FontTypeNames.FONTTYPE_INFO)
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

140         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageMeditateToggle(.Char.CharIndex, .Char.FX))

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

Private Sub HandleResucitate(ByVal Userindex As Integer)
        
        On Error GoTo HandleResucitate_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Se asegura que el target es un npc
104         If .flags.TargetNPC = 0 Then
106             Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate NPC and make sure player is dead
108         If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(Userindex))) Or .flags.Muerto = 0 Then Exit Sub
        
            'Make sure it's close enough
110         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
112             Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         Call RevivirUsuario(Userindex)
116         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Curar, 100, False))
118         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave("104", UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
120         Call WriteConsoleMsg(Userindex, "¡Has sido resucitado!!", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleHeal(ByVal Userindex As Integer)
        
        On Error GoTo HandleHeal_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Se asegura que el target es un npc
104         If .flags.TargetNPC = 0 Then
106             Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie) Or .flags.Muerto <> 0 Then Exit Sub
        
110         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
112             Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         .Stats.MinHp = .Stats.MaxHp
        
116         Call WriteUpdateHP(Userindex)
        
118         Call WriteConsoleMsg(Userindex, "ííHís sido curado!!", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleRequestStats(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleRequestStats_Err
        
100     Call UserList(Userindex).incomingData.ReadByte
    
102     Call SendUserStatsTxt(Userindex, Userindex)

        
        Exit Sub

HandleRequestStats_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestStats", Erl)
        Resume Next
        
End Sub

''
' Handles the "Help" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleHelp_Err
        
100     Call UserList(Userindex).incomingData.ReadByte
    
102     Call SendHelp(Userindex)

        
        Exit Sub

HandleHelp_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleHelp", Erl)
        Resume Next
        
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal Userindex As Integer)
        
        On Error GoTo HandleCommerceStart_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Dead people can't commerce
104         If .flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
106             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Is it already in commerce mode??
108         If .flags.Comerciando Then
110             Call WriteConsoleMsg(Userindex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
112         If .flags.TargetNPC > 0 Then

                'Does the NPC want to trade??
114             If Npclist(.flags.TargetNPC).Comercia = 0 Then
116                 If LenB(Npclist(.flags.TargetNPC).Desc) <> 0 Then
118                     Call WriteChatOverHead(Userindex, "No tengo ningún interés en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

                    End If
                
                    Exit Sub

                End If
            
120             If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
122                 Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Start commerce....
124             Call IniciarComercioNPC(Userindex)
                '[Alejo]
126         ElseIf .flags.TargetUser > 0 Then

                'User commerce...
                'Can he commerce??
128             If .flags.Privilegios And PlayerType.Consejero Then
130                 Call WriteConsoleMsg(Userindex, "No podés vender items.", FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub

                End If
            
                'Is the other one dead??
132             If UserList(.flags.TargetUser).flags.Muerto = 1 Then
134                 Call WriteConsoleMsg(Userindex, "¡¡No podés comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Is it me??
136             If .flags.TargetUser = Userindex Then
138                 Call WriteConsoleMsg(Userindex, "No podés comerciar con vos mismo...", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Check distance
140             If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
142                 Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Is he already trading?? is it with me or someone else??
144             If UserList(.flags.TargetUser).flags.Comerciando = True And UserList(.flags.TargetUser).ComUsu.DestUsu <> Userindex Then
146                 Call WriteConsoleMsg(Userindex, "No podés comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Initialize some variables...
148             .ComUsu.DestUsu = .flags.TargetUser
150             .ComUsu.DestNick = UserList(.flags.TargetUser).name
152             .ComUsu.cant = 0
154             .ComUsu.Objeto = 0
156             .ComUsu.Acepto = False
            
                'Rutina para comerciar con otro usuario
158             Call IniciarComercioConUsuario(Userindex, .flags.TargetUser)
            Else
160             Call WriteConsoleMsg(Userindex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleBankStart(ByVal Userindex As Integer)
        
        On Error GoTo HandleBankStart_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Dead people can't commerce
104         If .flags.Muerto = 1 Then
106             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If .flags.Comerciando Then
110             Call WriteConsoleMsg(Userindex, "Ya estás comerciando", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
112         If .flags.TargetNPC > 0 Then
114             If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 6 Then
116                 Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'If it's the banker....
118             If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
120                 Call IniciarDeposito(Userindex)

                End If

            Else
122             Call WriteConsoleMsg(Userindex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleEnlist(ByVal Userindex As Integer)
        
        On Error GoTo HandleEnlist_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Validate target NPC
104         If .flags.TargetNPC = 0 Then
106             Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
110         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
112             Call WriteConsoleMsg(Userindex, "Debes acercarte mís.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
116             Call EnlistarArmadaReal(Userindex)
            Else
118             Call EnlistarCaos(Userindex)

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

Private Sub HandleInformation(ByVal Userindex As Integer)
        
        On Error GoTo HandleInformation_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Validate target NPC
104         If .flags.TargetNPC = 0 Then
106             Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
110         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
112             Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
116             If .Faccion.ArmadaReal = 0 Then
118                 Call WriteChatOverHead(Userindex, "No perteneces a las tropas reales!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

120             Call WriteChatOverHead(Userindex, "Tu deber es combatir criminales, cada 100 criminales que derrotes te darí una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else

122             If .Faccion.FuerzasCaos = 0 Then
124                 Call WriteChatOverHead(Userindex, "No perteneces a la legiín oscura!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

126             Call WriteChatOverHead(Userindex, "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te darí una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

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

Private Sub HandleReward(ByVal Userindex As Integer)
        
        On Error GoTo HandleReward_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Validate target NPC
104         If .flags.TargetNPC = 0 Then
106             Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
110         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
112             Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
116             If .Faccion.ArmadaReal = 0 Then
118                 Call WriteChatOverHead(Userindex, "No perteneces a las tropas reales!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

120             Call RecompensaArmadaReal(Userindex)
            Else

122             If .Faccion.FuerzasCaos = 0 Then
124                 Call WriteChatOverHead(Userindex, "No perteneces a la legiín oscura!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

126             Call RecompensaCaos(Userindex)

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

Private Sub HandleRequestMOTD(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleRequestMOTD_Err
        
100     Call UserList(Userindex).incomingData.ReadByte
    
102     Call SendMOTD(Userindex)

        
        Exit Sub

HandleRequestMOTD_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestMOTD", Erl)
        Resume Next
        
End Sub

''
' Handles the "UpTime" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 01/10/08
        '01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleUpTime_Err
        
100     Call UserList(Userindex).incomingData.ReadByte
    
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
    
122     Call WriteConsoleMsg(Userindex, "Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)

        
        Exit Sub

HandleUpTime_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleUpTime", Erl)
        Resume Next
        
End Sub

''
' Handles the "Inquiry" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleInquiry(ByVal Userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        'Remove packet ID
        
        On Error GoTo HandleInquiry_Err
        
100     Call UserList(Userindex).incomingData.ReadByte
    
102     ConsultaPopular.SendInfoEncuesta (Userindex)

        
        Exit Sub

HandleInquiry_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleInquiry", Erl)
        Resume Next
        
End Sub

''
' Handles the "GuildMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
    
ErrHandler:

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

Private Sub HandleCentinelReport(ByVal Userindex As Integer)
        
        On Error GoTo HandleCentinelReport_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 3 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
108         Call CentinelaCheckClave(Userindex, .incomingData.ReadInteger())

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

Private Sub HandleGuildOnline(ByVal Userindex As Integer)
        
        On Error GoTo HandleGuildOnline_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            Dim onlineList As String
        
104         onlineList = modGuilds.m_ListaDeMiembrosOnline(Userindex, .GuildIndex)
        
106         If .GuildIndex <> 0 Then
108             Call WriteConsoleMsg(Userindex, "Compaíeros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
            Else
110             Call WriteConsoleMsg(Userindex, "No pertences a ningún clan.", FontTypeNames.FONTTYPE_GUILDMSG)

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

Private Sub HandleCouncilMessage(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call SendData(SendTarget.ToConsejo, Userindex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, Userindex, PrepareMessageConsoleMsg("(Consejero) " & .name & "> " & chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleRoleMasterRequest(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim request As String
        
        request = buffer.ReadASCIIString()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(Userindex, "Su solicitud ha sido enviada", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGMRequest(ByVal Userindex As Integer)
        
        On Error GoTo HandleGMRequest_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If Not Ayuda.Existe(.name) Then
106             Call WriteConsoleMsg(Userindex, "El mensaje ha sido entregado, ahora sílo debes esperar que se desocupe algín GM.", FontTypeNames.FONTTYPE_INFO)
                'Call Ayuda.Push(.name)
            Else
                'Call Ayuda.Quitar(.name)
                'Call Ayuda.Push(.name)
108             Call WriteConsoleMsg(Userindex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleChangeDescription(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim description As String
        
        description = buffer.ReadASCIIString()
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "No podés cambiar la descripción estando muerto.", FontTypeNames.FONTTYPE_INFOIAO)
        Else

            If Not AsciiValidos(description) Then
                Call WriteConsoleMsg(Userindex, "La descripción tiene carácteres inválidos.", FontTypeNames.FONTTYPE_INFOIAO)
            Else
                .Desc = Trim$(description)
                Call WriteConsoleMsg(Userindex, "La descripción a cambiado.", FontTypeNames.FONTTYPE_INFOIAO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGuildVote(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim vote     As String

        Dim errorStr As String
        
        vote = buffer.ReadASCIIString()
        
        If Not modGuilds.v_UsuarioVota(Userindex, vote, errorStr) Then
            Call WriteConsoleMsg(Userindex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(Userindex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandlePunishments(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                    Call WriteConsoleMsg(Userindex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                Else

                    While Count > 0

                        Call WriteConsoleMsg(Userindex, Count & " - " & GetVar(CharPath & name & ".chr", "PENAS", "P" & Count), FontTypeNames.FONTTYPE_INFO)
                        Count = Count - 1
                    Wend

                End If

            Else
                Call WriteConsoleMsg(Userindex, "Personaje """ & name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleChangePassword(ByVal Userindex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Creation Date: 10/10/07
    'Last Modified By: Ladder
    'Ahora cambia la password de la cuenta y no del PJ.
    '***************************************************

    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
            Call ChangePasswordDatabase(Userindex, SDesencriptar(oldPass), SDesencriptar(newPass))
        
        Else

            If LenB(SDesencriptar(newPass)) = 0 Then
                Call WriteConsoleMsg(Userindex, "Debe especificar una contraseña nueva, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Else
                oldPass2 = GetVar(CuentasPath & UserList(Userindex).Cuenta & ".act", "INIT", "PASSWORD")
                
                If SDesencriptar(oldPass2) <> SDesencriptar(oldPass) Then
                    Call WriteConsoleMsg(Userindex, "La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteVar(CuentasPath & UserList(Userindex).Cuenta & ".act", "INIT", "PASSWORD", newPass)
                    Call WriteConsoleMsg(Userindex, "La contraseña de su cuenta fue cambiada con éxito.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGamble(ByVal Userindex As Integer)
        
        On Error GoTo HandleGamble_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 3 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Amount As Integer
        
108         Amount = .incomingData.ReadInteger()
        
110         If .flags.Muerto = 1 Then
112             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
114         ElseIf .flags.TargetNPC = 0 Then
                'Validate target NPC
116             Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
118         ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
120             Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                ' Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
122         ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
124             Call WriteChatOverHead(Userindex, "No tengo ningún interés en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
126         ElseIf Amount < 1 Then
128             Call WriteChatOverHead(Userindex, "El mínimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
130         ElseIf Amount > 10000 Then
132             Call WriteChatOverHead(Userindex, "El míximo de apuesta es 10000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
134         ElseIf .Stats.GLD < Amount Then
136             Call WriteChatOverHead(Userindex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Else

138             If RandomNumber(1, 100) <= 45 Then
140                 .Stats.GLD = .Stats.GLD + Amount
142                 Call WriteChatOverHead(Userindex, "Felicidades! Has ganado " & CStr(Amount) & " monedas de oro!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
144                 Apuestas.Perdidas = Apuestas.Perdidas + Amount
146                 Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
148                 .Stats.GLD = .Stats.GLD - Amount
150                 Call WriteChatOverHead(Userindex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                
152                 Apuestas.Ganancias = Apuestas.Ganancias + Amount
154                 Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))

                End If
            
156             Apuestas.Jugadas = Apuestas.Jugadas + 1
            
158             Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
160             Call WriteUpdateGold(Userindex)

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

Private Sub HandleInquiryVote(ByVal Userindex As Integer)
        
        On Error GoTo HandleInquiryVote_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim opt As Byte
        
108         opt = .incomingData.ReadByte()
        
110         Call WriteConsoleMsg(Userindex, ConsultaPopular.doVotar(Userindex, opt), FontTypeNames.FONTTYPE_GUILD)

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

Private Sub HandleBankExtractGold(ByVal Userindex As Integer)
        
        On Error GoTo HandleBankExtractGold_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 5 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Amount As Long
        
108         Amount = .incomingData.ReadLong()
        
            'Dead people can't leave a faction.. they can't talk...
110         If .flags.Muerto = 1 Then
112             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
114         If .flags.TargetNPC = 0 Then
116             Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
118         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
120         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
122             Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
124         If Amount > 0 And Amount <= .Stats.Banco Then
126             .Stats.Banco = .Stats.Banco - Amount
128             .Stats.GLD = .Stats.GLD + Amount
130             'Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

                Call WriteUpdateGold(Userindex)
                Call WriteGoliathInit(Userindex)
            Else
132             Call WriteChatOverHead(Userindex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

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

Private Sub HandleLeaveFaction(ByVal Userindex As Integer)
        
        On Error GoTo HandleLeaveFaction_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Dead people can't leave a faction.. they can't talk...
104         If .flags.Muerto = 1 Then
106             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
108         If .Faccion.ArmadaReal = 0 And .Faccion.FuerzasCaos = 0 Then
110             If .Faccion.Status = 1 Then
112                 Call VolverCriminal(Userindex)
114                 Call WriteConsoleMsg(Userindex, "Ahora sos un criminal.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If

            Else

                ' Call WriteConsoleMsg(UserIndex, "Ya sos un criminal.", FontTypeNames.FONTTYPE_INFOIAO)
                ' Exit Sub
            End If
        
            'Validate target NPC
116         If .flags.TargetNPC = 0 Then
118             If .Faccion.ArmadaReal = 1 Then
120                 Call WriteConsoleMsg(Userindex, "Para salir del ejercito debes ir a visitar al rey.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
122             ElseIf .Faccion.FuerzasCaos = 1 Then
124                 Call WriteConsoleMsg(Userindex, "Para salir de la legion debes ir a visitar al diablo.", FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub

                End If

            End If
        
126         If .flags.TargetNPC = 0 Then Exit Sub
        
128         If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Enlistador Then

                'Quit the Royal Army?
130             If .Faccion.ArmadaReal = 1 Then
132                 If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
134                     Call ExpulsarFaccionReal(Userindex)
136                     Call WriteChatOverHead(Userindex, "Serís bienvenido a las fuerzas imperiales si deseas regresar.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                        Exit Sub
                    Else
138                     Call WriteChatOverHead(Userindex, "¡¡¡Sal de aquí bufón!!!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                   
                    End If

                    'Quit the Chaos Legion??
140             ElseIf .Faccion.FuerzasCaos = 1 Then

142                 If Npclist(.flags.TargetNPC).flags.Faccion = 1 Then
144                     Call ExpulsarFaccionCaos(Userindex)
146                     Call WriteChatOverHead(Userindex, "Ya volverís arrastrandote.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                    Else
148                     Call WriteChatOverHead(Userindex, "Sal de aquí maldito criminal", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

                    End If

                Else
150                 Call WriteChatOverHead(Userindex, "¡No perteneces a ninguna facción!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

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

Private Sub HandleBankDepositGold(ByVal Userindex As Integer)
        
        On Error GoTo HandleBankDepositGold_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 5 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Amount As Long
        
108         Amount = .incomingData.ReadLong()
        
            'Dead people can't leave a faction.. they can't talk...
110         If .flags.Muerto = 1 Then
112             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
114         If .flags.TargetNPC = 0 Then
116             Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
118         If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
120             Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
122         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
124         If Amount > 0 And Amount <= .Stats.GLD Then
126             .Stats.Banco = .Stats.Banco + Amount
128             .Stats.GLD = .Stats.GLD - Amount
130             'Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            
132             Call WriteUpdateGold(Userindex)
                Call WriteGoliathInit(Userindex)
            Else
134             Call WriteChatOverHead(Userindex, "No tenés esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)

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

Private Sub HandleDenounce(ByVal Userindex As Integer)
        
        On Error GoTo HandleDenounce_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte

104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
106         If EventoActivo Then
108             Call FinalizarEvento
            Else
110             Call WriteConsoleMsg(Userindex, "No hay ningun evento activo.", FontTypeNames.FONTTYPE_New_Eventos)
        
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

Private Sub HandleGuildMemberList(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "No existe el clan: " & guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
                For i = 1 To memberCount
                    UserName = GetVar(App.Path & "\Guilds\" & guild & "-Members" & ".mem", "Members", "Member" & i)
                    
                    Call WriteConsoleMsg(Userindex, UserName & "<" & guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleGMMessage(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String
        
        message = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.user Then
            Call LogGM(.name, "Mensaje a Gms:" & message)
        
            If LenB(message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(message)
            
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & "> " & message, FontTypeNames.FONTTYPE_GMMSG))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleShowName(ByVal Userindex As Integer)
        
        On Error GoTo HandleShowName_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
106             .showName = Not .showName 'Show / Hide the name
            
108             Call RefreshCharStatus(Userindex)

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

Private Sub HandleOnlineRoyalArmy(ByVal Userindex As Integer)
        
        On Error GoTo HandleOnlineRoyalArmy_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
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
120         Call WriteConsoleMsg(Userindex, "Armadas conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
        Else
122         Call WriteConsoleMsg(Userindex, "No hay Armadas conectados", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleOnlineChaosLegion(ByVal Userindex As Integer)
        
        On Error GoTo HandleOnlineChaosLegion_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
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
120         Call WriteConsoleMsg(Userindex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
        Else
122         Call WriteConsoleMsg(Userindex, "No hay Caos conectados", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleGoNearby(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/10/07
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String
        
        UserName = buffer.ReadASCIIString()
        
        Dim tIndex As Integer

        Dim X      As Long

        Dim Y      As Long

        Dim i      As Long

        Dim found  As Boolean
        
        tIndex = NameIndex(UserName)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then

            'Si es dios o Admins no podemos salvo que nosotros tambiín lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex <= 0 Then 'existe el usuario destino?
                    Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else

                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i

                                If MapData(UserList(tIndex).Pos.Map, X, Y).Userindex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.Map, X, Y, True, True) Then
                                        Call WarpUserChar(Userindex, UserList(tIndex).Pos.Map, X, Y, True)
                                        found = True
                                        Exit For

                                    End If

                                End If

                            Next Y
                            
                            If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next X
                        
                        If found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next i
                    
                    'No space found??
                    If Not found Then
                        Call WriteConsoleMsg(Userindex, "Todos los lugares estín ocupados.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleComment(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim comment As String

        comment = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.user Then
            Call LogGM(.name, "Comentario: " & comment)
            Call WriteConsoleMsg(Userindex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleServerTime(ByVal Userindex As Integer)
        
        On Error GoTo HandleServerTime_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 01/08/07
        'Last Modification by: (liquid)
        '***************************************************
100     With UserList(Userindex)
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

Private Sub HandleWhere(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                If (UserList(tUser).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(Userindex, "Ubicaciín  " & UserName & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "/Donde " & UserName)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleCreaturesInMap(ByVal Userindex As Integer)
        
        On Error GoTo HandleCreaturesInMap_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 30/07/06
        'Pablo (ToxicWaste): modificaciones generales para simplificar la visualizaciín.
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 3 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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
128                             List1(0) = Npclist(i).name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
130                             NPCcant1(0) = 1
                            Else

132                             For j = 0 To NPCcount1 - 1

134                                 If Left$(List1(j), Len(Npclist(i).name)) = Npclist(i).name Then
136                                     List1(j) = List1(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
138                                     NPCcant1(j) = NPCcant1(j) + 1
                                        Exit For

                                    End If

140                             Next j

142                             If j = NPCcount1 Then
144                                 ReDim Preserve List1(0 To NPCcount1) As String
146                                 ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
148                                 NPCcount1 = NPCcount1 + 1
150                                 List1(j) = Npclist(i).name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
152                                 NPCcant1(j) = 1

                                End If

                            End If

                        Else

154                         If NPCcount2 = 0 Then
156                             ReDim List2(0) As String
158                             ReDim NPCcant2(0) As Integer
160                             NPCcount2 = 1
162                             List2(0) = Npclist(i).name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
164                             NPCcant2(0) = 1
                            Else

166                             For j = 0 To NPCcount2 - 1

168                                 If Left$(List2(j), Len(Npclist(i).name)) = Npclist(i).name Then
170                                     List2(j) = List2(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
172                                     NPCcant2(j) = NPCcant2(j) + 1
                                        Exit For

                                    End If

174                             Next j

176                             If j = NPCcount2 Then
178                                 ReDim Preserve List2(0 To NPCcount2) As String
180                                 ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
182                                 NPCcount2 = NPCcount2 + 1
184                                 List2(j) = Npclist(i).name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
186                                 NPCcant2(j) = 1

                                End If

                            End If

                        End If

                    End If

188             Next i
            
190             Call WriteConsoleMsg(Userindex, "Npcs Hostiles en mapa: ", FontTypeNames.FONTTYPE_WARNING)

192             If NPCcount1 = 0 Then
194                 Call WriteConsoleMsg(Userindex, "No hay NPCS Hostiles", FontTypeNames.FONTTYPE_INFO)
                Else

196                 For j = 0 To NPCcount1 - 1
198                     Call WriteConsoleMsg(Userindex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
200                 Next j

                End If

202             Call WriteConsoleMsg(Userindex, "Otros Npcs en mapa: ", FontTypeNames.FONTTYPE_WARNING)

204             If NPCcount2 = 0 Then
206                 Call WriteConsoleMsg(Userindex, "No hay mís NPCS", FontTypeNames.FONTTYPE_INFO)
                Else

208                 For j = 0 To NPCcount2 - 1
210                     Call WriteConsoleMsg(Userindex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
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

Private Sub HandleWarpMeToTarget(ByVal Userindex As Integer)
        
        On Error GoTo HandleWarpMeToTarget_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And PlayerType.user Then Exit Sub
        
106         Call WarpUserChar(Userindex, .flags.TargetMap, .flags.TargetX, .flags.TargetY, True)
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

Private Sub HandleWarpChar(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 7 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim Map      As Integer

        Dim X        As Byte

        Dim Y        As Byte

        Dim tUser    As Integer
        
        UserName = buffer.ReadASCIIString()
        Map = buffer.ReadInteger()
        X = buffer.ReadByte()
        Y = buffer.ReadByte()
        
        If Not .flags.Privilegios And PlayerType.user Then
            If MapaValido(Map) And LenB(UserName) <> 0 Then
                If UCase$(UserName) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(UserName)

                    End If

                Else
                    tUser = Userindex

                End If
            
                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                ElseIf InMapBounds(Map, X, Y) Then
                    Call FindLegalPos(tUser, Map, X, Y)
                    Call WarpUserChar(tUser, Map, X, Y, True)
                    Call WriteConsoleMsg(Userindex, UserList(tUser).name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                    If tUser <> Userindex Then Call LogGM(.name, "Transportí a " & UserList(tUser).name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleSilence(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(Userindex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "ESTIMADO USUARIO, ud ha sido silenciado por los administradores. Sus denuncias serín ignoradas por el servidor de aquí en mís. Utilice /GM para contactar un administrador.")
                    Call LogGM(.name, "/silenciar " & UserList(tUser).name)
                
                    'Flush the other user's buffer
                    
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(Userindex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.name, "/DESsilenciar " & UserList(tUser).name)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleSOSShowList(ByVal Userindex As Integer)
        
        On Error GoTo HandleSOSShowList_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And PlayerType.user Then Exit Sub
106         Call WriteShowSOSForm(Userindex)

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

Private Sub HandleSOSRemove(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
    
ErrHandler:

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

Private Sub HandleGoToChar(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim X        As Byte

        Dim Y        As Byte
        
        UserName = buffer.ReadASCIIString()
        tUser = NameIndex(UserName)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then

            'Si es dios o Admins no podemos salvo que nosotros tambiín lo seamos
            If Not (EsDios(UserName) Or EsAdmin(UserName)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                
                    X = UserList(tUser).Pos.X
                    Y = UserList(tUser).Pos.Y + 1
                    Call FindLegalPos(Userindex, UserList(tUser).Pos.Map, X, Y)
                
                    Call WarpUserChar(Userindex, UserList(tUser).Pos.Map, X, Y, True)
                    
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                        

                    End If
                    
                    Call LogGM(.name, "/IRA " & UserName & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleDesbuggear(ByVal Userindex As Integer)

    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                    Call WriteConsoleMsg(Userindex, "El usuario debe estar offline.", FontTypeNames.FONTTYPE_INFO)
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
                            Call WriteConsoleMsg(Userindex, "Hay un usuario de la cuenta conectado. Se actualizaron solo los usuarios online.", FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call ResetLoggedDatabase(AccountID)
                            Call WriteConsoleMsg(Userindex, "Cuenta del personaje desbuggeada y usuarios online actualizados.", FontTypeNames.FONTTYPE_INFO)

                        End If
    
                        Call LogGM(.name, "/DESBUGGEAR " & UserName)
                    Else
                        Call WriteConsoleMsg(Userindex, "El usuario no existe.", FontTypeNames.FONTTYPE_INFO)

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
                
                Call WriteConsoleMsg(Userindex, "Se actualizaron los usuarios online.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleDarLlaveAUsuario(ByVal Userindex As Integer)

    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Es necesario que el juego esté corriendo con base de datos.", FontTypeNames.FONTTYPE_INFO)
            
            ' Me aseguro que el objeto sea una llave válida
            ElseIf Llave < 1 Or Llave > NumObjDatas Then
                Call WriteConsoleMsg(Userindex, "El número ingresado no es el de una llave válida.", FontTypeNames.FONTTYPE_INFO)
            ElseIf ObjData(Llave).OBJType <> eOBJType.otLlaves Then ' vb6 no tiene short-circuit evaluation :(
                Call WriteConsoleMsg(Userindex, "El número ingresado no es el de una llave válida.", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                
                If tUser > 0 Then
                    ' Es un user online, guardamos la llave en la db
                    If DarLlaveAUsuarioDatabase(UserName, Llave) Then
                        ' Actualizamos su llavero
                        If MeterLlaveEnLLavero(tUser, Llave) Then
                            Call WriteConsoleMsg(Userindex, "Llave número " & Llave & " entregada a " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(Userindex, "No se pudo entregar la llave. El usuario no tiene más espacio en su llavero.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
                        Call WriteConsoleMsg(Userindex, "No se pudo entregar la llave. Asegúrese de que la llave esté disponible.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    ' No es un usuario online, nos fijamos si es un email
                    If CheckMailString(UserName) Then
                        ' Es un email, intentamos guardarlo en la db
                        If DarLlaveACuentaDatabase(UserName, Llave) Then
                            Call WriteConsoleMsg(Userindex, "Llave número " & Llave & " entregada a " & LCase$(UserName) & ".", FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(Userindex, "No se pudo entregar la llave. Asegúrese de que la llave esté disponible y que el email sea correcto.", FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
                        Call WriteConsoleMsg(Userindex, "El usuario no está online. Ingrese el email de la cuenta para otorgar la llave offline.", FontTypeNames.FONTTYPE_INFO)
                    End If
    
                End If
                
                Call LogGM(.name, "/DARLLAVE " & UserName & " " & Llave)
            End If
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleSacarLlave(ByVal Userindex As Integer)

    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)

        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Llave As Integer
        
        Llave = .incomingData.ReadInteger()
        
        ' Solo dios o admin
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            ' Me aseguro que esté activada la db
            If Not Database_Enabled Then
                Call WriteConsoleMsg(Userindex, "Es necesario que el juego esté corriendo con base de datos.", FontTypeNames.FONTTYPE_INFO)

            Else
                ' Intento borrarla de la db
                If SacarLlaveDatabase(Llave) Then
                    Call WriteConsoleMsg(Userindex, "La llave " & Llave & " fue removida.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "No se pudo sacar la llave. Asegúrese de que esté en uso.", FontTypeNames.FONTTYPE_INFO)
                End If

                Call LogGM(.name, "/SACARLLAVE " & Llave)
            End If
        End If

    End With

End Sub

Private Sub HandleVerLlaves(ByVal Userindex As Integer)

    With UserList(Userindex)
    
        Call .incomingData.ReadByte

        ' Sólo GMs
        If Not (.flags.Privilegios And PlayerType.user) Then
            ' Me aseguro que esté activada la db
            If Not Database_Enabled Then
                Call WriteConsoleMsg(Userindex, "Es necesario que el juego esté corriendo con base de datos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Leo y muestro todas las llaves usadas
            Call VerLlavesDatabase(Userindex)
        End If
                
    End With

End Sub

Private Sub HandleUseKey(ByVal Userindex As Integer)

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(Userindex)
    
        Call .incomingData.ReadByte
        
        Dim slot As Byte
        slot = .incomingData.ReadByte

        Call UsarLlave(Userindex, slot)
                
    End With

End Sub

''
' Handles the "Invisible" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal Userindex As Integer)
        
        On Error GoTo HandleInvisible_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And PlayerType.user Then Exit Sub
        
106         Call DoAdminInvisible(Userindex)
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

Private Sub HandleGMPanel(ByVal Userindex As Integer)
        
        On Error GoTo HandleGMPanel_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And PlayerType.user Then Exit Sub
        
106         Call WriteShowGMPanelForm(Userindex)

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

Private Sub HandleRequestUserList(ByVal Userindex As Integer)
        
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
    
100     With UserList(Userindex)
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
        
120         If Count > 1 Then Call WriteUserNameList(Userindex, names(), Count - 1)

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

Private Sub HandleWorking(ByVal Userindex As Integer)
        
        On Error GoTo HandleWorking_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim i     As Long

        Dim Users As String
    
100     With UserList(Userindex)
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
120             Call WriteConsoleMsg(Userindex, "Usuarios trabajando: " & Users, FontTypeNames.FONTTYPE_INFO)
            Else
122             Call WriteConsoleMsg(Userindex, "No hay usuarios trabajando", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleHiding(ByVal Userindex As Integer)
        
        On Error GoTo HandleHiding_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim i     As Long

        Dim Users As String
    
100     With UserList(Userindex)
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
118             Call WriteConsoleMsg(Userindex, "Usuarios ocultandose: " & Users, FontTypeNames.FONTTYPE_INFO)
            Else
120             Call WriteConsoleMsg(Userindex, "No hay usuarios ocultandose", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleJail(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(UserName)
                
                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "El usuario no está online.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If Not UserList(tUser).flags.Privilegios And PlayerType.user Then
                        Call WriteConsoleMsg(Userindex, "No podés encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > 60 Then
                        Call WriteConsoleMsg(Userindex, "No podés encarcelar por mís de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
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
    
ErrHandler:

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

Private Sub HandleKillNPC(ByVal Userindex As Integer)
        
    On Error GoTo HandleKillNPC_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(Userindex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.user Then Exit Sub

        'Si estamos en el mapa pretoriano...
        If .Pos.Map = MAPA_PRETORIANO Then

            '... solo los Dioses y Administradores pueden usar este comando en el mapa pretoriano.
            If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) = 0 Then
                
                Call WriteConsoleMsg(Userindex, "Solo los Administradores y Dioses pueden usar este comando en el mapa pretoriano.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If
        
        Dim tNPC As Integer: tNPC = .flags.TargetNPC
        
        If tNPC > 0 Then

            Call WriteConsoleMsg(Userindex, "RMatas (con posible respawn) a: " & Npclist(tNPC).name, FontTypeNames.FONTTYPE_INFO)
            
            Dim auxNPC As npc: auxNPC = Npclist(tNPC)
            
            Call QuitarNPC(tNPC)
            Call ReSpawnNpc(auxNPC)
            
        Else
            Call WriteConsoleMsg(Userindex, "Debes hacer click sobre el NPC antes", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleWarnUser(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/26/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                privs = UserDarPrivilegioLevel(UserName)
                
                If Not privs And PlayerType.user Then
                    Call WriteConsoleMsg(Userindex, "No podés advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
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
                        
                        Call WriteConsoleMsg(Userindex, "Has advertido a " & UCase$(UserName), FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.name, " advirtio a " & UserName)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleMensajeUser(ByVal Userindex As Integer)

    '***************************************************
    'Author: Ladder
    'Last Modification: 04/jul/2014
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Utilice /MENSAJEINFORMACION nick@mensaje", FontTypeNames.FONTTYPE_INFO)
            Else
                privs = UserDarPrivilegioLevel(UserName)
                
                If (InStrB(UserName, "\") <> 0) Then
                    UserName = Replace(UserName, "\", "")

                End If

                If (InStrB(UserName, "/") <> 0) Then
                    UserName = Replace(UserName, "/", "")

                End If
                    
                AddCorreo Userindex, UserName, LCase$(Mensaje), 0, 0
                    
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
    
ErrHandler:

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
Private Sub HandleTraerBoveda(ByVal Userindex As Integer)

    '***************************************************
    'Author: Ladder
    'Last Modification: 04/jul/2014
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        Call UpdateUserHechizos(True, Userindex, 0)
       
        Call UpdateUserInv(True, Userindex, 0)
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleEditChar(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/28/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 8 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
            tUser = Userindex
            
        Else
            tUser = NameIndex(UserName)

        End If
        
        opcion = buffer.ReadByte()
        Arg1 = buffer.ReadASCIIString()
        Arg2 = buffer.ReadASCIIString()
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
        ' Si no es GM, no hacemos nada.
        If Not EsGM(Userindex) Then Exit Sub
        
        ' Si NO sos Dios o Admin,
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) = 0 Then

            ' Si te editas a vos mismo esta bien ;)
            If Userindex <> tUser Then Exit Sub
            
        End If
        
        Select Case opcion

            Case eEditOptions.eo_Gold

                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    UserList(tUser).Stats.GLD = val(Arg1)
                    Call WriteUpdateGold(tUser)

                End If
                
            Case eEditOptions.eo_Experience

                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else

                    If UserList(tUser).Stats.ELV < STAT_MAXELV Then
                        UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
                        Call CheckUserLevel(tUser)
                        Call WriteUpdateExp(tUser)
                            
                    Else
                        Call WriteConsoleMsg(Userindex, "El usuario es nivel máximo.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
                
            Case eEditOptions.eo_Body

                If tUser <= 0 Then
                    
                    If Database_Enabled Then
                        Call SaveUserBodyDatabase(UserName, val(Arg1))
                    Else
                        Call WriteVar(CharPath & UserName & ".chr", "INIT", "Body", Arg1)

                    End If

                    Call WriteConsoleMsg(Userindex, "Usuario Offline Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

                End If
                
            Case eEditOptions.eo_Head

                If tUser <= 0 Then
                    
                    If Database_Enabled Then
                        Call SaveUserHeadDatabase(UserName, val(Arg1))
                    Else
                        Call WriteVar(CharPath & UserName & ".chr", "INIT", "Head", Arg1)

                    End If

                    Call WriteConsoleMsg(Userindex, "Usuario Offline Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call ChangeUserChar(tUser, UserList(tUser).Char.Body, val(Arg1), UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim)

                End If
                
            Case eEditOptions.eo_CriminalsKilled

                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else

                    If val(Arg1) > MAXUSERMATADOS Then
                        UserList(tUser).Faccion.CriminalesMatados = MAXUSERMATADOS
                    Else
                        UserList(tUser).Faccion.CriminalesMatados = val(Arg1)

                    End If

                End If
                
            Case eEditOptions.eo_CiticensKilled

                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else

                    If val(Arg1) > MAXUSERMATADOS Then
                        UserList(tUser).Faccion.CiudadanosMatados = MAXUSERMATADOS
                    Else
                        UserList(tUser).Faccion.CiudadanosMatados = val(Arg1)

                    End If

                End If
                
            Case eEditOptions.eo_Level

                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else

                    If val(Arg1) > STAT_MAXELV Then
                        Arg1 = CStr(STAT_MAXELV)
                        Call WriteConsoleMsg(Userindex, "No podés tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)

                    End If
                        
                    UserList(tUser).Stats.ELV = val(Arg1)

                End If
                    
                Call WriteUpdateUserStats(Userindex)
                
            Case eEditOptions.eo_Class

                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else

                    For LoopC = 1 To NUMCLASES

                        If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then Exit For
                    Next LoopC
                        
                    If LoopC > NUMCLASES Then
                        Call WriteConsoleMsg(Userindex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        UserList(tUser).clase = LoopC

                    End If

                End If
                
            Case eEditOptions.eo_Skills

                For LoopC = 1 To NUMSKILLS

                    If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg1) Then Exit For
                Next LoopC
                    
                If LoopC > NUMSKILLS Then
                    Call WriteConsoleMsg(Userindex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                Else

                    If tUser <= 0 Then
                        
                        If Database_Enabled Then
                            Call SaveUserSkillDatabase(UserName, LoopC, val(Arg2))
                        Else
                            Call WriteVar(CharPath & UserName & ".chr", "Skills", "SK" & LoopC, Arg2)

                        End If

                        Call WriteConsoleMsg(Userindex, "Usuario Offline Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
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
                        
                    Call WriteConsoleMsg(Userindex, "Usuario Offline Alterado: " & UserName, FontTypeNames.FONTTYPE_INFO)
                Else
                    UserList(tUser).Stats.SkillPts = val(Arg1)

                End If
                
            Case eEditOptions.eo_Sex

                If tUser <= 0 Then
                    Call WriteConsoleMsg(Userindex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                    
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
                    Call WriteConsoleMsg(Userindex, "Usuario offline: " & UserName, FontTypeNames.FONTTYPE_INFO)
                        
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
                
                Call WriteConsoleMsg(Userindex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)

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
                
            Case Else
                commandString = commandString & "UNKOWN "

        End Select
        
        commandString = commandString & Arg1 & " " & Arg2
        
        Call LogGM(.name, commandString & " " & UserName)

    End With

ErrHandler:

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

Private Sub HandleRequestCharInfo(ByVal Userindex As Integer)

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid).. alto bug zapallo..
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                    Call WriteConsoleMsg(Userindex, "Usuario offline, Buscando en Charfile.", FontTypeNames.FONTTYPE_INFO)
                    Call SendUserStatsTxtOFF(Userindex, targetName)

                End If

            Else

                'don't allow to retrieve administrator's info
                If UserList(targetIndex).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(Userindex, targetIndex)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Private Sub HandleRequestCharStats(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Usuario offline. Leyendo Charfile... ", FontTypeNames.FONTTYPE_INFO)
                
                Call SendUserMiniStatsTxtFromChar(Userindex, UserName)
            Else
                Call SendUserMiniStatsTxt(Userindex, tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleRequestCharGold(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserOROTxtFromChar(Userindex, UserName)
            Else
                Call WriteConsoleMsg(Userindex, "El usuario " & UserName & " tiene " & UserList(tUser).Stats.Banco & " en el banco", FontTypeNames.FONTTYPE_TALK)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleRequestCharInventory(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Usuario offline. Leyendo del charfile...", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserInvTxtFromChar(Userindex, UserName)
            Else
                Call SendUserInvTxt(Userindex, tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleRequestCharBank(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Usuario offline. Leyendo charfile... ", FontTypeNames.FONTTYPE_TALK)
                
                Call SendUserBovedaTxtFromChar(Userindex, UserName)
            Else
                Call SendUserBovedaTxt(Userindex, tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleRequestCharSkills(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim tUser    As Integer

        Dim LoopC    As Long

        Dim message  As String
        
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
                    message = message & "CHAR>" & SkillsNames(LoopC) & " = " & GetVar(CharPath & UserName & ".chr", "SKILLS", "SK" & LoopC) & vbCrLf
                Next LoopC
                
                Call WriteConsoleMsg(Userindex, message & "CHAR> Libres:" & GetVar(CharPath & UserName & ".chr", "STATS", "SKILLPTSLIBRES"), FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendUserSkillsTxt(Userindex, tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleReviveChar(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                tUser = Userindex

            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                With UserList(tUser)

                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0
                        
                        'Call DarCuerpoDesnudo(tUser)
                        
                        'Call ChangeUserChar(tUser, .Char.body, .OrigChar.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                        Call RevivirUsuario(tUser)
                        
                        Call WriteConsoleMsg(tUser, UserList(Userindex).name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(Userindex).name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)

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

ErrHandler:

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

Private Sub HandleOnlineGM(ByVal Userindex As Integer)
        
        On Error GoTo HandleOnlineGM_Err
        

        '***************************************************
        'Author: Fredy Horacio Treboux (liquid)
        'Last Modification: 12/28/06
        '
        '***************************************************
        Dim i    As Long

        Dim list As String

        Dim priv As PlayerType
    
100     With UserList(Userindex)
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
122             Call WriteConsoleMsg(Userindex, list & ".", FontTypeNames.FONTTYPE_INFO)
            Else
124             Call WriteConsoleMsg(Userindex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleOnlineMap(ByVal Userindex As Integer)
        
        On Error GoTo HandleOnlineMap_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(Userindex)
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
        
120         Call WriteConsoleMsg(Userindex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleForgive(ByVal Userindex As Integer)
        
        On Error GoTo HandleForgive_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            'Se asegura que el target es un npc
104         If .flags.TargetNPC = 0 Then
106             Call WriteConsoleMsg(Userindex, "Primero tenés que seleccionar al sacerdote.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate NPC and make sure player is dead
108         If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor And (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.ResucitadorNewbie Or Not EsNewbie(Userindex))) Or .flags.Muerto = 1 Then Exit Sub
        
            'Make sure it's close enough
110         If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
                'Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
112             Call WriteConsoleMsg(Userindex, "El sacerdote no puede escuchar tus pecados debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
114         If UserList(Userindex).Faccion.Status = 1 Or UserList(Userindex).Faccion.ArmadaReal = 1 Then
                'Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)
116             Call WriteChatOverHead(Userindex, "Tu alma ya esta libre de pecados hijo mio.", Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If
        
118         If UserList(Userindex).Faccion.CiudadanosMatados > 0 Or UserList(Userindex).Faccion.ArmadaReal > 0 Then
120             Call WriteChatOverHead(Userindex, "Has matado gente inocente, lamentablemente no podre concebirte el perdon.", Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If
        
            Dim Clanalineacion As Byte
                        
122         If .GuildIndex <> 0 Then
124             Clanalineacion = modGuilds.Alineacion(.GuildIndex)

126             If Clanalineacion = 1 Then
128                 Call WriteChatOverHead(Userindex, "Te encuentras en un clan criminal... debes retirarte para que pueda perdonarte.", Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

            End If
        
130         Call WriteChatOverHead(Userindex, "Con estas palabras, te libero de todo tipo de pecados. íQue dios te acompaíe hijo mio!", Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, vbYellow)

132         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, "80", 100, False))
134         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave("100", UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
136         UserList(Userindex).Faccion.Status = 1
138         Call RefreshCharStatus(Userindex)

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

Private Sub HandleKick(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else

                If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(Userindex, "No podes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
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

ErrHandler:

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

Private Sub HandleExecute(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "No está online", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleBanChar(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
            Call BanCharacter(Userindex, UserName, Reason)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleSilenciarUser(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim UserName As String

        Dim Time     As Byte
        
        UserName = buffer.ReadASCIIString()
        Time = buffer.ReadByte()
    
        Call SilenciarUserName(Userindex, UserName, Time)
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleUnbanChar(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Charfile inexistente (no use +)", FontTypeNames.FONTTYPE_INFO)
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
                    Call WriteConsoleMsg(Userindex, UserName & " desbaneado.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, UserName & " no esta baneado. Imposible unbanear", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleNPCFollow(ByVal Userindex As Integer)
        
        On Error GoTo HandleNPCFollow_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(Userindex)
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

Private Sub HandleSummonChar(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "El jugador no esta online.", FontTypeNames.FONTTYPE_INFO)
            Else

                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.user)) <> 0 Then
                    Call WriteConsoleMsg(tUser, .name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                    Call WarpToLegalPos(tUser, .Pos.Map, .Pos.X, .Pos.Y + 1, True)
                    
                    If UserList(tUser).flags.BattleModo = 1 Then
                        Call WriteConsoleMsg(Userindex, "¡¡¡ATENCIÓN!!! [" & UCase(UserList(tUser).name) & "] SE ENCUENTRA EN MODO BATTLE.", FontTypeNames.FONTTYPE_WARNING)
                        Call LogGM(.name, "¡¡¡ATENCIÓN /SUM EN MODO BATTLE " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                    Else
                        Call LogGM(.name, "/SUM " & UserName & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)

                    End If
                    
                Else
                    Call WriteConsoleMsg(Userindex, "No podés invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleSpawnListRequest(ByVal Userindex As Integer)
        
        On Error GoTo HandleSpawnListRequest_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub
        
106         Call EnviarSpawnList(Userindex)

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

Private Sub HandleSpawnCreature(ByVal Userindex As Integer)
        
        On Error GoTo HandleSpawnCreature_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 3 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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

Private Sub HandleResetNPCInventory(ByVal Userindex As Integer)
        
        On Error GoTo HandleResetNPCInventory_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(Userindex)
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

Private Sub HandleCleanWorld(ByVal Userindex As Integer)
        
        On Error GoTo HandleCleanWorld_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte

104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

            Call LimpiezaForzada
            
            Call WriteConsoleMsg(Userindex, "Se han limpiado los items del suelo.", FontTypeNames.FONTTYPE_INFO)
            
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

Private Sub HandleServerMessage(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue
        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String

        message = buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(message) <> 0 Then
                Call LogGM(.name, "Mensaje Broadcast:" & message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & "> " & message, FontTypeNames.FONTTYPE_SERVER))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleNickToIP(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 24/07/07
    'Pablo (ToxicWaste): Agrego para uqe el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                    Call WriteConsoleMsg(Userindex, "El ip de " & UserName & " es " & UserList(tUser).ip, FontTypeNames.FONTTYPE_INFO)

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
                    Call WriteConsoleMsg(Userindex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(Userindex, "No hay ningun personaje con ese nick", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleIPToNick(ByVal Userindex As Integer)
        
        On Error GoTo HandleIPToNick_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 5 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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
140         Call WriteConsoleMsg(Userindex, "Los personajes con ip " & ip & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleGuildOnlineMembers(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Clan " & UCase(GuildName) & ": " & modGuilds.m_ListaDeMiembrosOnline(Userindex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleTeleportCreate(ByVal Userindex As Integer)
        
        On Error GoTo HandleTeleportCreate_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 5 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim Mapa As Integer

            Dim X    As Byte

            Dim Y    As Byte
        
108         Mapa = .incomingData.ReadInteger()
110         X = .incomingData.ReadByte()
112         Y = .incomingData.ReadByte()
        
114         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
116         Call LogGM(.name, "/CT " & Mapa & "," & X & "," & Y)
        
118         If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
120         If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then Exit Sub
        
122         If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then Exit Sub
        
124         If MapData(Mapa, X, Y).ObjInfo.ObjIndex > 0 Then
126             Call WriteConsoleMsg(Userindex, "Hay un objeto en el piso en ese lugar", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
128         If MapData(Mapa, X, Y).TileExit.Map > 0 Then
130             Call WriteConsoleMsg(Userindex, "No podés crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            Rem Call WriteParticleFloorCreate(UserIndex, 37, -1, .Pos.map, .Pos.X, .Pos.Y - 1)
        
            Dim Objeto As obj
        
132         Objeto.Amount = 1
134         Objeto.ObjIndex = 378
136         Call MakeObj(Objeto, .Pos.Map, .Pos.X, .Pos.Y - 1)
        
138         With MapData(.Pos.Map, .Pos.X, .Pos.Y - 1)
140             .TileExit.Map = Mapa
142             .TileExit.X = X
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

Private Sub HandleTeleportDestroy(ByVal Userindex As Integer)
        
        On Error GoTo HandleTeleportDestroy_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(Userindex)

            Dim Mapa As Integer

            Dim X    As Byte

            Dim Y    As Byte
        
            'Remove packet ID
102         Call .incomingData.ReadByte
        
            '/dt
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
106         Mapa = .flags.TargetMap
108         X = .flags.TargetX
110         Y = .flags.TargetY
        
112         If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
114         With MapData(Mapa, X, Y)

116             If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
118             If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.Map > 0 Then
120                 Call LogGM(UserList(Userindex).name, "/DT: " & Mapa & "," & X & "," & Y)
                
122                 Call EraseObj(.ObjInfo.Amount, Mapa, X, Y)
                
124                 If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
126                     Call EraseObj(1, .TileExit.Map, .TileExit.X, .TileExit.Y)

                    End If
                
128                 .TileExit.Map = 0
130                 .TileExit.X = 0
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

Private Sub HandleRainToggle(ByVal Userindex As Integer)
        
        On Error GoTo HandleRainToggle_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     With UserList(Userindex)
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

Private Sub HandleSetCharDescription(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Haz click sobre un personaje antes!", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HanldeForceMIDIToMap(ByVal Userindex As Integer)
        
        On Error GoTo HanldeForceMIDIToMap_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 4 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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

Private Sub HandleForceWAVEToMap(ByVal Userindex As Integer)
        
        On Error GoTo HandleForceWAVEToMap_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 6 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim waveID As Byte

            Dim Mapa   As Integer

            Dim X      As Byte

            Dim Y      As Byte
        
108         waveID = .incomingData.ReadByte()
110         Mapa = .incomingData.ReadInteger()
112         X = .incomingData.ReadByte()
114         Y = .incomingData.ReadByte()
        
            'Solo dioses, admins y RMS
116         If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

                'Si el mapa no fue enviado tomo el actual
118             If Not InMapBounds(Mapa, X, Y) Then
120                 Mapa = .Pos.Map
122                 X = .Pos.X
124                 Y = .Pos.Y

                End If
            
                'Ponemos el pedido por el GM
126             Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayWave(waveID, X, Y))

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

Private Sub HandleRoyalArmyMessage(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String

        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("ARMADA REAL> " & message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleChaosLegionMessage(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String

        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleCitizenMessage(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String

        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleCriminalMessage(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String

        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleTalkAsNPC(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String

        message = buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(message, Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite))
            Else
                Call WriteConsoleMsg(Userindex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleDestroyAllItemsInArea(ByVal Userindex As Integer)
        
        On Error GoTo HandleDestroyAllItemsInArea_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
            Dim X As Long

            Dim Y As Long
        
106         For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
108             For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

110                 If X > 0 And Y > 0 And X < 101 And Y < 101 Then
112                     If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then
114                         If ItemNoEsDeMapa(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex) Then
116                             Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, X, Y)

                            End If

                        End If

                    End If

118             Next X
120         Next Y
        
122         Call LogGM(UserList(Userindex).name, "/MASSDEST")

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

Private Sub HandleAcceptRoyalCouncilMember(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", FontTypeNames.FONTTYPE_CONSEJO))

                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)

                End With

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleAcceptChaosCouncilMember(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el Consejo de la Legiín Oscura.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)

                End With

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleItemsInTheFloor(ByVal Userindex As Integer)
        
        On Error GoTo HandleItemsInTheFloor_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
            Dim tObj  As Integer

            Dim lista As String

            Dim X     As Long

            Dim Y     As Long
        
106         For X = 5 To 95
108             For Y = 5 To 95
110                 tObj = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex

112                 If tObj > 0 Then
114                     If ObjData(tObj).OBJType <> eOBJType.otArboles Then
116                         Call WriteConsoleMsg(Userindex, "(" & X & "," & Y & ") " & ObjData(tObj).name, FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

118             Next Y
120         Next X

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

Private Sub HandleMakeDumb(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleMakeDumbNoMore(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)
                

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleDumpIPTables(ByVal Userindex As Integer)
        
        On Error GoTo HandleDumpIPTables_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(Userindex)
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

Private Sub HandleCouncilKick(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                    Call WriteConsoleMsg(Userindex, "Usuario offline, echando de los consejos", FontTypeNames.FONTTYPE_INFO)
                    
                    If Database_Enabled Then
                        Call EcharConsejoDatabase(UserName)
                    Else
                        Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECE", 0)
                        Call WriteVar(CharPath & UserName & ".chr", "CONSEJO", "PERTENECECAOS", 0)

                    End If

                Else
                    Call WriteConsoleMsg(Userindex, "No existe el personaje.", FontTypeNames.FONTTYPE_INFO)

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
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de la Legiín Oscura", FontTypeNames.FONTTYPE_CONSEJO))

                    End If

                End With

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleSetTrigger(ByVal Userindex As Integer)
        
        On Error GoTo HandleSetTrigger_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim tTrigger As Byte

            Dim tLog     As String
        
108         tTrigger = .incomingData.ReadByte()
        
110         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
112         If tTrigger >= 0 Then
114             MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = tTrigger
116             tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y
            
118             Call LogGM(.name, tLog)
120             Call WriteConsoleMsg(Userindex, tLog, FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleAskTrigger(ByVal Userindex As Integer)
        
        On Error GoTo HandleAskTrigger_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 04/13/07
        '
        '***************************************************
        Dim tTrigger As Byte
    
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger
        
108         Call LogGM(.name, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
110         Call WriteConsoleMsg(Userindex, "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleBannedIPList(ByVal Userindex As Integer)
        
        On Error GoTo HandleBannedIPList_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(Userindex)
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
        
116         Call WriteConsoleMsg(Userindex, lista, FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleBannedIPReload(ByVal Userindex As Integer)
        
        On Error GoTo HandleBannedIPReload_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(Userindex)
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

Private Sub HandleGuildBan(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
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

ErrHandler:

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

Private Sub HandleBanIP(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "El personaje no está online.", FontTypeNames.FONTTYPE_INFO)
            Else
                bannedIP = UserList(tUser).ip

            End If

        End If
        
        Reason = buffer.ReadASCIIString()
        
        If LenB(bannedIP) > 0 Then
            If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                Call LogGM(.name, "/BanIP " & bannedIP & " por " & Reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(Userindex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                Call BanIpAgrega(bannedIP)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " baneí la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))
                
                'Find every player with that ip and ban him!
                For i = 1 To LastUser

                    If UserList(i).ConnIDValida Then
                        If UserList(i).ip = bannedIP Then
                            Call BanCharacter(Userindex, UserList(i).name, "IP POR " & Reason)

                        End If

                    End If

                Next i

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleUnbanIP(ByVal Userindex As Integer)
        
        On Error GoTo HandleUnbanIP_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 5 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadByte
        
            Dim bannedIP As String
        
108         bannedIP = .incomingData.ReadByte() & "."
110         bannedIP = bannedIP & .incomingData.ReadByte() & "."
112         bannedIP = bannedIP & .incomingData.ReadByte() & "."
114         bannedIP = bannedIP & .incomingData.ReadByte()
        
116         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
118         If BanIpQuita(bannedIP) Then
120             Call WriteConsoleMsg(Userindex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
            Else
122             Call WriteConsoleMsg(Userindex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)

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

Private Sub HandleCreateItem(ByVal Userindex As Integer)
        
    On Error GoTo HandleCreateItem_Err
        

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(Userindex)

        'Remove packet ID
        Call .incomingData.ReadByte

        Dim tObj    As Integer
        Dim Cuantos As Integer
        
        tObj = .incomingData.ReadInteger()
        Cuantos = .incomingData.ReadInteger()
        
        ' Si es usuario o consejero, lo sacamos cagando.
        If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub
        
        ' Si es Semi-Dios, dejamos crear un item siempre y cuando pueda estar en el inventario.
        If (.flags.Privilegios And PlayerType.SemiDios) And ObjData(tObj).Agarrable = 1 Then Exit Sub
        
        If ObjData(tObj).donador = 1 Then
            ' Si es usuario, consejero o Semi-Dios y trata de crear un objeto para donadores, lo sacamos cagando.
            If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios) Then Exit Sub
        End If
        
        ' Si hace mas de 10000, lo sacamos cagando.
        If Cuantos > MAX_INVENTORY_OBJS Then
            Call WriteConsoleMsg(Userindex, "Solo podés crear hasta " & CStr(MAX_INVENTORY_OBJS) & " unidades", FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        End If
        
        ' Hay un TileExit donde estoy creando el objeto?
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then Exit Sub
        
        ' El indice proporcionado supera la cantidad minima o total de items existentes en el juego?
        If tObj < 1 Or tObj > NumObjDatas Then Exit Sub
        
        ' El nombre del objeto es nulo?
        If LenB(ObjData(tObj).name) = 0 Then Exit Sub
        
        Dim Objeto As obj
            Objeto.Amount = Cuantos
            Objeto.ObjIndex = tObj

        ' Chequeo si el objeto es AGARRABLE(para las puertas, arboles y demas objs. que no deberian estar en el inventario)
        '   0 = SI
        '   1 = NO
        If ObjData(tObj).Agarrable = 0 Then
            
            ' Trato de meterlo en el inventario.
            If MeterItemEnInventario(Userindex, Objeto) Then
                Call WriteConsoleMsg(Userindex, "Has creado " & Objeto.Amount & " unidades de " & ObjData(tObj).name & ".", FontTypeNames.FONTTYPE_INFO)
            
            Else
                ' Si no hay espacio, lo tiro al piso.
                Call TirarItemAlPiso(.Pos, Objeto)
                Call WriteConsoleMsg(Userindex, "No tenes espacio en tu inventario para crear el item.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(Userindex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
                
            End If
        
        Else
        
            ' Crear el item NO AGARRARBLE y tirarlo al piso.
            Call TirarItemAlPiso(.Pos, Objeto)
            Call WriteConsoleMsg(Userindex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
            
        End If
        
        ' Lo registro en los logs.
        Call LogGM(.name, "/CI: " & tObj & " Cantidad : " & Cuantos)

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

Private Sub HandleDestroyItems(ByVal Userindex As Integer)
        
        On Error GoTo HandleDestroyItems_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
106         If MapData(.Pos.Map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex = 0 Then Exit Sub
        
108         Call LogGM(.name, "/DEST")
        
            ' If ObjData(MapData(.Pos.map, .Pos.X, .Pos.Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then
            ''  Call WriteConsoleMsg(UserIndex, "No puede destruir teleports así. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            '  Exit Sub
            ' End If
        
110         Call EraseObj(10000, .Pos.Map, .Pos.X, .Pos.Y)

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

Private Sub HandleChaosLegionKick(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
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
                    
                    Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas del caos y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "El personaje " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleRoyalArmyKick(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
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

                    Call WriteConsoleMsg(Userindex, UserName & " expulsado de las fuerzas reales y prohibida la reenlistada", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, UserName & ".chr inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleForceMIDIAll(ByVal Userindex As Integer)
        
        On Error GoTo HandleForceMIDIAll_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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

Private Sub HandleForceWAVEAll(ByVal Userindex As Integer)
        
        On Error GoTo HandleForceWAVEAll_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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

Private Sub HandleRemovePunishment(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 1/05/07
    'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
    '***************************************************
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
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
                    
                    Call WriteConsoleMsg(Userindex, "Pena Modificada.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Private Sub HandleTileBlockedToggle(ByVal Userindex As Integer)
        
        On Error GoTo HandleTileBlockedToggle_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub

106         Call LogGM(.name, "/BLOQ")
        
108         If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0 Then
110             MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = eBlock.ALL_SIDES Or eBlock.GM
            Else
112             MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0

            End If
        
114         Call Bloquear(True, .Pos.Map, .Pos.X, .Pos.Y, IIf(MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked > 0, eBlock.ALL_SIDES, 0))

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

Private Sub HandleKillNPCNoRespawn(ByVal Userindex As Integer)
        
        On Error GoTo HandleKillNPCNoRespawn_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(Userindex)
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

Private Sub HandleKillAllNearbyNPCs(ByVal Userindex As Integer)
        
        On Error GoTo HandleKillAllNearbyNPCs_Err
        

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
            
            'Si está en el mapa pretoriano, me aseguro de que los saque correctamente antes que nada.
106         If .Pos.Map = MAPA_PRETORIANO Then Call EliminarPretorianos(MAPA_PRETORIANO)

            Dim X As Long
            Dim Y As Long
        
108         For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
110             For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

112                 If X > 0 And Y > 0 And X < 101 And Y < 101 Then

114                     If MapData(.Pos.Map, X, Y).NpcIndex > 0 Then
116                         Call QuitarNPC(MapData(.Pos.Map, X, Y).NpcIndex)
                        End If

                    End If

118             Next X
120         Next Y

122         Call LogGM(.name, "/MASSKILL")

        End With

        
        Exit Sub

HandleKillAllNearbyNPCs_Err:
124     Call RegistrarError(Err.Number, Err.description, "Protocol.HandleKillAllNearbyNPCs", Erl)
126     Resume Next
        
End Sub

''
' Handles the "LastIP" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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

                    Call WriteConsoleMsg(Userindex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "Charfile """ & UserName & """ inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(Userindex, UserName & " es de mayor jerarquía que vos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Public Sub HandleChatColor(ByVal Userindex As Integer)
        
        On Error GoTo HandleChatColor_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
        'Change the user`s chat color
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 4 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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

Public Sub HandleIgnored(ByVal Userindex As Integer)
        
        On Error GoTo HandleIgnored_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Ignore the user
        '***************************************************
100     With UserList(Userindex)
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

Public Sub HandleCheckSlot(ByVal Userindex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Check one Users Slot in Particular from Inventory
    '***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo ErrHandler

    With UserList(Userindex)

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
           
        If tIndex > 0 And UserList(Userindex).flags.BattleModo = 0 Then
            If slot > 0 And slot <= UserList(Userindex).CurrentInventorySlots Then
                If UserList(tIndex).Invent.Object(slot).ObjIndex > 0 Then
                    Call WriteConsoleMsg(Userindex, " Objeto " & slot & ") " & ObjData(UserList(tIndex).Invent.Object(slot).ObjIndex).name & " Cantidad:" & UserList(tIndex).Invent.Object(slot).Amount, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "No hay Objeto en slot seleccionado", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(Userindex, "Slot Invílido.", FontTypeNames.FONTTYPE_TALK)

            End If

        Else
            Call WriteConsoleMsg(Userindex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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

Public Sub HandleResetAutoUpdate(ByVal Userindex As Integer)
        
        On Error GoTo HandleResetAutoUpdate_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Reset the AutoUpdate
        '***************************************************
100     With UserList(Userindex)
            'Remove packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
106         Call WriteConsoleMsg(Userindex, "TID: " & CStr(ReiniciarAutoUpdate()), FontTypeNames.FONTTYPE_INFO)

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

Public Sub HandleRestart(ByVal Userindex As Integer)
        
        On Error GoTo HandleRestart_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Restart the game
        '***************************************************
100     With UserList(Userindex)
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

Public Sub HandleReloadObjects(ByVal Userindex As Integer)
        
        On Error GoTo HandleReloadObjects_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Reload the objects
        '***************************************************
100     With UserList(Userindex)
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         Call LogGM(.name, .name & " ha recargado a los objetos.")
        
108         Call LoadOBJData
110         Call LoadPesca
112         Call LoadRecursosEspeciales
114         Call WriteConsoleMsg(Userindex, "Obj.dat recargado exitosamente.", FontTypeNames.FONTTYPE_SERVER)

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

Public Sub HandleReloadSpells(ByVal Userindex As Integer)
        
        On Error GoTo HandleReloadSpells_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Reload the spells
        '***************************************************
100     With UserList(Userindex)
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

Public Sub HandleReloadServerIni(ByVal Userindex As Integer)
        
        On Error GoTo HandleReloadServerIni_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Reload the Server`s INI
        '***************************************************
100     With UserList(Userindex)
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

Public Sub HandleReloadNPCs(ByVal Userindex As Integer)
        
        On Error GoTo HandleReloadNPCs_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Reload the Server`s NPC
        '***************************************************
100     With UserList(Userindex)
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
         
106         Call LogGM(.name, .name & " ha recargado los NPCs.")
    
108         Call CargaNpcsDat
    
110         Call WriteConsoleMsg(Userindex, "Npcs.dat recargado exitosamente.", FontTypeNames.FONTTYPE_SERVER)

        End With

        
        Exit Sub

HandleReloadNPCs_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleReloadNPCs", Erl)
        Resume Next
        
End Sub

''
' Handle the "RequestTCPStats" message
' @param UserIndex The index of the user sending the message

Public Sub HandleRequestTCPStats(ByVal Userindex As Integer)
        
        On Error GoTo HandleRequestTCPStats_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
        'Send the TCP`s stadistics
        '***************************************************
100     With UserList(Userindex)
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
                
            Dim list  As String

            Dim Count As Long

            Dim i     As Long
        
106         Call LogGM(.name, .name & " ha pedido las estadisticas del TCP.")
    
108         Call WriteConsoleMsg(Userindex, "Los datos estín en BYTES.", FontTypeNames.FONTTYPE_INFO)
        
            'Send the stats
110         With TCPESStats
112             Call WriteConsoleMsg(Userindex, "IN/s: " & .BytesRecibidosXSEG & " OUT/s: " & .BytesEnviadosXSEG, FontTypeNames.FONTTYPE_INFO)
114             Call WriteConsoleMsg(Userindex, "IN/s MAX: " & .BytesRecibidosXSEGMax & " -> " & .BytesRecibidosXSEGCuando, FontTypeNames.FONTTYPE_INFO)
116             Call WriteConsoleMsg(Userindex, "OUT/s MAX: " & .BytesEnviadosXSEGMax & " -> " & .BytesEnviadosXSEGCuando, FontTypeNames.FONTTYPE_INFO)

            End With
        
            'Search for users that are working
118         For i = 1 To LastUser

120             With UserList(i)

122                 If .flags.UserLogged And .ConnID >= 0 And .ConnIDValida Then
124                     If .outgoingData.Length > 0 Then
126                         list = list & .name & " (" & CStr(.outgoingData.Length) & "), "
128                         Count = Count + 1

                        End If

                    End If

                End With

130         Next i
        
132         Call WriteConsoleMsg(Userindex, "Posibles pjs trabados: " & CStr(Count), FontTypeNames.FONTTYPE_INFO)
134         Call WriteConsoleMsg(Userindex, list, FontTypeNames.FONTTYPE_INFO)

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

Public Sub HandleKickAllChars(ByVal Userindex As Integer)
        
        On Error GoTo HandleKickAllChars_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Kick all the chars that are online
        '***************************************************
100     With UserList(Userindex)
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

Public Sub HandleNight(ByVal Userindex As Integer)
        
        On Error GoTo HandleNight_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
106
            HoraMundo = (timeGetTime And &H7FFFFFFF)

            Call SendData(SendTarget.ToAll, 0, PrepareMessageHora())
    
        End With

        
        Exit Sub

HandleNight_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleNight", Erl)
        Resume Next
        
End Sub

''
' Handle the "Day" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleDay(ByVal Userindex As Integer)
        
        On Error GoTo HandleDay_Err

100     With UserList(Userindex)
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
106
            HoraMundo = (timeGetTime And &H7FFFFFFF) - DuracionDia \ 2

            Call SendData(SendTarget.ToAll, 0, PrepareMessageHora())
    
        End With

        
        Exit Sub

HandleDay_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleDay", Erl)
        Resume Next
        
End Sub

''
' Handle the "SetTime" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSetTime(ByVal Userindex As Integer)
        
        On Error GoTo HandleSetTime_Err

100     With UserList(Userindex)
            'Remove Packet ID
102         Call .incomingData.ReadByte

            Dim HoraDia As Long
            HoraDia = .incomingData.ReadLong
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
106
            HoraMundo = (timeGetTime And &H7FFFFFFF) - HoraDia
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageHora())
    
        End With

        
        Exit Sub

HandleSetTime_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleSetTime", Erl)
        Resume Next
        
End Sub

''
' Handle the "ShowServerForm" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal Userindex As Integer)
        
        On Error GoTo HandleShowServerForm_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Show the server form
        '***************************************************
100     With UserList(Userindex)
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

Public Sub HandleCleanSOS(ByVal Userindex As Integer)
        
        On Error GoTo HandleCleanSOS_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Clean the SOS
        '***************************************************
100     With UserList(Userindex)
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

Public Sub HandleSaveChars(ByVal Userindex As Integer)
        
        On Error GoTo HandleSaveChars_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Save the characters
        '***************************************************
100     With UserList(Userindex)
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

Public Sub HandleChangeMapInfoBackup(ByVal Userindex As Integer)
        
        On Error GoTo HandleChangeMapInfoBackup_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/24/06
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
        'Change the backup`s info of the map
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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
        
122         Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " Backup: " & MapInfo(.Pos.Map).backup_mode, FontTypeNames.FONTTYPE_INFO)

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

Public Sub HandleChangeMapInfoPK(ByVal Userindex As Integer)
        
        On Error GoTo HandleChangeMapInfoPK_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/24/06
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
        'Change the pk`s info of the  map
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove Packet ID
106         Call .incomingData.ReadByte
        
            Dim isMapPk As Boolean
        
108         isMapPk = .incomingData.ReadBoolean()
        
110         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
112         Call LogGM(.name, .name & " ha cambiado la informacion sobre si es seguro el mapa.")
        
114         MapInfo(.Pos.Map).Seguro = isMapPk
        
            'Change the boolean to string in a fast way
            Rem Call WriteVar(App.Path & MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "Pk", IIf(isMapPk, "1", "0"))

116         Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " PK: " & MapInfo(.Pos.Map).Seguro, FontTypeNames.FONTTYPE_INFO)

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

Public Sub HandleChangeMapInfoRestricted(ByVal Userindex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    Dim tStr As String
    
    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.name, .name & " ha cambiado la informacion sobre si es Restringido el mapa.")
                MapInfo(UserList(Userindex).Pos.Map).restrict_mode = tStr
                Call WriteVar(MapPath & "mapa" & UserList(Userindex).Pos.Map & ".dat", "Mapa" & UserList(Userindex).Pos.Map, "Restringir", tStr)
                Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " Restringido: " & MapInfo(.Pos.Map).restrict_mode, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Public Sub HandleChangeMapInfoNoMagic(ByVal Userindex As Integer)
        
        On Error GoTo HandleChangeMapInfoNoMagic_Err
        

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'MagiaSinEfecto -> Options: "1" , "0".
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
        Dim nomagic As Boolean
    
104     With UserList(Userindex)
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

Public Sub HandleChangeMapInfoNoInvi(ByVal Userindex As Integer)
        
        On Error GoTo HandleChangeMapInfoNoInvi_Err
        

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'InviSinEfecto -> Options: "1", "0"
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
        Dim noinvi As Boolean
    
104     With UserList(Userindex)
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

Public Sub HandleChangeMapInfoNoResu(ByVal Userindex As Integer)
        
        On Error GoTo HandleChangeMapInfoNoResu_Err
        

        '***************************************************
        'Author: Pablo (ToxicWaste)
        'Last Modification: 26/01/2007
        'ResuSinEfecto -> Options: "1", "0"
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 2 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
        Dim noresu As Boolean
    
104     With UserList(Userindex)
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

Public Sub HandleChangeMapInfoLand(ByVal Userindex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    Dim tStr As String
    
    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la informacion del Terreno del mapa.")
                MapInfo(UserList(Userindex).Pos.Map).terrain = tStr
                Call WriteVar(MapPath & "mapa" & UserList(Userindex).Pos.Map & ".dat", "Mapa" & UserList(Userindex).Pos.Map, "Terreno", tStr)
                Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " Terreno: " & MapInfo(.Pos.Map).terrain, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(Userindex, "Igualmente, el ínico ítil es 'NIEVE' ya que al ingresarlo, la gente muere de frio en el Mapa", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Public Sub HandleChangeMapInfoZone(ByVal Userindex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    Dim tStr As String
    
    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call buffer.ReadByte
        
        tStr = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la informacion de la Zona del mapa.")
                MapInfo(UserList(Userindex).Pos.Map).zone = tStr
                Call WriteVar(MapPath & "mapa" & UserList(Userindex).Pos.Map & ".dat", "Mapa" & UserList(Userindex).Pos.Map, "Zona", tStr)
                Call WriteConsoleMsg(Userindex, "Mapa " & .Pos.Map & " Zona: " & MapInfo(.Pos.Map).zone, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(Userindex, "Igualmente, el ínico ítil es 'DUNGEON' ya que al ingresarlo, NO se sentirí el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Public Sub HandleSaveMap(ByVal Userindex As Integer)
        
        On Error GoTo HandleSaveMap_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/24/06
        'Saves the map
        '***************************************************
100     With UserList(Userindex)
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         Call LogGM(.name, .name & " ha guardado el mapa " & CStr(.Pos.Map))
        
            ' Call GrabarMapa(.Pos.Map, App.Path & "\WorldBackUp\Mapa" & .Pos.Map)
        
108         Call WriteConsoleMsg(Userindex, "Mapa Guardado", FontTypeNames.FONTTYPE_INFO)

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

Public Sub HandleShowGuildMessages(ByVal Userindex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Allows admins to read guild messages
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim guild As String
        
        guild = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(Userindex, guild)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Public Sub HandleDoBackUp(ByVal Userindex As Integer)
        
        On Error GoTo HandleDoBackUp_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/24/06
        'Show guilds messages
        '***************************************************
100     With UserList(Userindex)
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

Public Sub HandleToggleCentinelActivated(ByVal Userindex As Integer)
        
        On Error GoTo HandleToggleCentinelActivated_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/26/06
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
        'Activate or desactivate the Centinel
        '***************************************************
100     With UserList(Userindex)
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

Public Sub HandleAlterName(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Change user name
    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(UserName)
                
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(Userindex, "El Pj esta online, debe salir para el cambio", FontTypeNames.FONTTYPE_WARNING)
                Else

                    If Not FileExist(CharPath & UserName & ".chr") Then
                        Call WriteConsoleMsg(Userindex, "El pj " & UserName & " es inexistente ", FontTypeNames.FONTTYPE_INFO)
                    Else
                        GuildIndex = val(GetVar(CharPath & UserName & ".chr", "GUILD", "GUILDINDEX"))
                        
                        If GuildIndex > 0 Then
                            Call WriteConsoleMsg(Userindex, "El pj " & UserName & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else

                            If Not FileExist(CharPath & newName & ".chr") Then
                                Call FileCopy(CharPath & UserName & ".chr", CharPath & UCase$(newName) & ".chr")
                                
                                Call WriteConsoleMsg(Userindex, "Transferencia exitosa", FontTypeNames.FONTTYPE_INFO)
                                
                                Call WriteVar(CharPath & UserName & ".chr", "BAN", "Baneado", "1")
                                Call WriteVar(CharPath & UserName & ".chr", "BAN", "BanMotivo", "BAN POR Cambio de nick a " & UCase$(newName) & " " & Date & " " & Time)
                                Call WriteVar(CharPath & UserName & ".chr", "BAN", "BannedBy", .name)

                                Dim cantPenas As Byte
                                
                                cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", CStr(cantPenas + 1))
                                
                                Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & CStr(cantPenas + 1), LCase$(.name) & ": BAN POR Cambio de nick a " & UCase$(newName) & " " & Date & " " & Time)
                                
                                Call LogGM(.name, "Ha cambiado de nombre al usuario " & UserName & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(Userindex, "El nick solicitado ya existe", FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Public Sub HandleAlterMail(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Change user password
    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "usar /AEMAIL <pj>-<nuevomail>", FontTypeNames.FONTTYPE_INFO)
            Else

                If Not FileExist(CharPath & UserName & ".chr") Then
                    Call WriteConsoleMsg(Userindex, "No existe el charfile " & UserName & ".chr", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteVar(CharPath & UserName & ".chr", "CONTACTO", "Email", newMail)
                    Call WriteConsoleMsg(Userindex, "Email de " & UserName & " cambiado a: " & newMail, FontTypeNames.FONTTYPE_INFO)

                End If
                
                Call LogGM(.name, "Le ha cambiado el mail a " & UserName)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Public Sub HandleAlterPassword(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Change user password
    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "usar /APASS <pjsinpass>@<pjconpass>", FontTypeNames.FONTTYPE_INFO)
            Else

                If Not FileExist(CharPath & UserName & ".chr") Or Not FileExist(CharPath & copyFrom & ".chr") Then
                    Call WriteConsoleMsg(Userindex, "Alguno de los PJs no existe " & UserName & "@" & copyFrom, FontTypeNames.FONTTYPE_INFO)
                Else
                    Password = GetVar(CharPath & copyFrom & ".chr", "INIT", "Password")
                    Call WriteVar(CharPath & UserName & ".chr", "INIT", "Password", Password)
                    
                    Call WriteConsoleMsg(Userindex, "Password de " & UserName & " cambiado a: " & Password, FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Public Sub HandleCreateNPC(ByVal Userindex As Integer)
        
    On Error GoTo HandleCreateNPC_Err
        

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(Userindex)
    
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NpcIndex As Integer: NpcIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        'Nos fijamos si es pretoriano.
        If Npclist(NpcIndex).NPCtype = eNPCType.Pretoriano Then
            Call WriteConsoleMsg(Userindex, "No puedes sumonear miembros del clan pretoriano de esta forma, utiliza /CrearPretoianos MAPA X Y.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub

        End If
        
        NpcIndex = SpawnNpc(NpcIndex, .Pos, True, False)
        
        If NpcIndex <> 0 Then
            Call LogGM(.name, "Sumoneo a " & Npclist(NpcIndex).name & " en mapa " & .Pos.Map)
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

Public Sub HandleCreateNPCWithRespawn(ByVal Userindex As Integer)
        
        On Error GoTo HandleCreateNPCWithRespawn_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/24/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 3 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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

Public Sub HandleImperialArmour(ByVal Userindex As Integer)
        
        On Error GoTo HandleImperialArmour_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/24/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 4 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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

Public Sub HandleChaosArmour(ByVal Userindex As Integer)
        
        On Error GoTo HandleChaosArmour_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/24/06
        '
        '***************************************************
100     If UserList(Userindex).incomingData.Length < 4 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
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

Public Sub HandleNavigateToggle(ByVal Userindex As Integer)
        
        On Error GoTo HandleNavigateToggle_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 01/12/07
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then Exit Sub
        
106         If .flags.Navegando = 1 Then
108             .flags.Navegando = 0
            Else
110             .flags.Navegando = 1

            End If
        
            'Tell the client that we are navigating.
112         Call WriteNavigateToggle(Userindex)

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

Public Sub HandleServerOpenToUsersToggle(ByVal Userindex As Integer)
        
        On Error GoTo HandleServerOpenToUsersToggle_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/24/06
        '
        '***************************************************
100     With UserList(Userindex)
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
106         If ServerSoloGMs > 0 Then
108             Call WriteConsoleMsg(Userindex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
110             ServerSoloGMs = 0
            Else
112             Call WriteConsoleMsg(Userindex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
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

Public Sub HandleParticipar(ByVal Userindex As Integer)
        
        On Error GoTo HandleParticipar_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/24/06
        'Turns off the server
        '***************************************************
        Dim handle As Integer
    
100     With UserList(Userindex)
            'Remove Packet ID
102         Call .incomingData.ReadByte
        
104         If Torneo.HayTorneoaActivo = False Then
106             Call WriteConsoleMsg(Userindex, "No hay ningún evento disponible.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
                
108         If .flags.BattleModo = 1 Then
110             Call WriteConsoleMsg(Userindex, "No podes participar desde aquí.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
    
112         If .flags.EnTorneo Then
114             Call WriteConsoleMsg(Userindex, "Ya estás participando.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
    
116         If .Stats.ELV > Torneo.nivelmaximo Then
118             Call WriteConsoleMsg(Userindex, "El nivel míximo para participar es " & Torneo.nivelmaximo & ".", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
    
120         If .Stats.ELV < Torneo.NivelMinimo Then
122             Call WriteConsoleMsg(Userindex, "El nivel mínimo para participar es " & Torneo.NivelMinimo & ".", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
    
124         If .Stats.GLD < Torneo.costo Then
126             Call WriteConsoleMsg(Userindex, "No tienes suficiente oro para ingresar.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
128         If .clase = Mage And Torneo.mago = 0 Then
130             Call WriteConsoleMsg(Userindex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
132         If .clase = Cleric And Torneo.clerico = 0 Then
134             Call WriteConsoleMsg(Userindex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
136         If .clase = Warrior And Torneo.guerrero = 0 Then
138             Call WriteConsoleMsg(Userindex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
140         If .clase = Bard And Torneo.bardo = 0 Then
142             Call WriteConsoleMsg(Userindex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
144         If .clase = Assasin And Torneo.asesino = 0 Then
146             Call WriteConsoleMsg(Userindex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
   
148         If .clase = Druid And Torneo.druido = 0 Then
150             Call WriteConsoleMsg(Userindex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
152         If .clase = Paladin And Torneo.Paladin = 0 Then
154             Call WriteConsoleMsg(Userindex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
156         If .clase = Hunter And Torneo.cazador = 0 Then
158             Call WriteConsoleMsg(Userindex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
160         If .clase = Trabajador And Torneo.cazador = 0 Then
162             Call WriteConsoleMsg(Userindex, "Tu clase no participa de este evento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
   
164         If Torneo.Participantes = Torneo.cupos Then
166             Call WriteConsoleMsg(Userindex, "Los cupos ya estan llenos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
  
168         Call ParticiparTorneo(Userindex)

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

Public Sub HandleTurnCriminal(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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

ErrHandler:

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

Public Sub HandleResetFactions(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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

ErrHandler:

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

Public Sub HandleRemoveCharFromGuild(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
            
            GuildIndex = modGuilds.m_EcharMiembroDeClan(Userindex, UserName)
            
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(Userindex, "No pertenece a ningún clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(Userindex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Public Sub HandleRequestCharMail(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Request user mail
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                
                Call WriteConsoleMsg(Userindex, "Last email de " & UserName & ":" & mail, FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Public Sub HandleSystemMessage(ByVal Userindex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/29/06
    'Send a message to all the users
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim message As String

        message = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.name, "Mensaje de sistema:" & message)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(message))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Public Sub HandleSetMOTD(ByVal Userindex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 03/31/07
    'Set the MOTD
    'Modified by: Juan Martín Sotuyo Dodero (Maraxus)
    '   - Fixed a bug that prevented from properly setting the new number of lines.
    '   - Fixed a bug that caused the player to be kicked.
    '***************************************************
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
            
            Call WriteConsoleMsg(Userindex, "Se ha cambiado el MOTD con exito", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

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

Public Sub HandleChangeMOTD(ByVal Userindex As Integer)
        
        On Error GoTo HandleChangeMOTD_Err
        

        '***************************************************
        'Author: Juan Martín sotuyo Dodero (Maraxus)
        'Last Modification: 12/29/06
        'Change the MOTD
        '***************************************************
100     With UserList(Userindex)
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
        
118         Call WriteShowMOTDEditionForm(Userindex, auxiliaryString)

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

Public Sub HandlePing(ByVal Userindex As Integer)
        
        On Error GoTo HandlePing_Err
        

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/24/06
        'Show guilds messages
        '***************************************************
100     With UserList(Userindex)
            'Remove Packet ID
102         Call .incomingData.ReadByte

            Dim Time As Long
        
104         Time = .incomingData.ReadLong()
        
106         Call WritePong(Userindex, Time)

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

Public Sub WriteLoggedMessage(ByVal Userindex As Integer)

    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.logged)
    Exit Sub
ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteHora(ByVal Userindex As Integer)

    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageHora())
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal Userindex As Integer, ByVal CharIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteNadarToggle(ByVal Userindex As Integer, ByVal Puede As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.NadarToggle)
        Call .WriteBoolean(Puede)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If
    
End Sub

Public Sub WriteEquiteToggle(ByVal Userindex As Integer)
        
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.EquiteToggle)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteVelocidadToggle(ByVal Userindex As Integer)
        
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.VelocidadToggle)
        Call .WriteSingle(UserList(Userindex).Char.speeding)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteMacroTrabajoToggle(ByVal Userindex As Integer, ByVal Activar As Boolean)

    If Not Activar Then
        UserList(Userindex).flags.TargetObj = 0 ' Sacamos el targer del objeto
        UserList(Userindex).flags.UltimoMensaje = 0
        UserList(Userindex).Counters.Trabajando = 0
        UserList(Userindex).flags.UsandoMacro = False
       
    Else
        UserList(Userindex).flags.UsandoMacro = True

    End If

    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.MacroTrabajoToggle)
        Call .WriteBoolean(Activar)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Disconnect" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    'If UserList(UserIndex).flags.BattleModo = 0 Then
    '    Call SaveUser(UserIndex, CharPath & UCase$(UserList(UserIndex).Name) & ".chr")
    'Else
    '    Call WriteVar(CharPath & UserList(UserIndex).Name & ".chr", "Battle", "Puntos", UserList(UserIndex).flags.BattlePuntos)
    'End If
    
    'Call WriteVar(CuentasPath & UCase$(UserList(UserIndex).cuenta) & ".act", "INIT", "LOGEADA", 0)
    Call WritePersonajesDeCuenta(Userindex)

    Call WriteMostrarCuenta(Userindex)
    
    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Disconnect)

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BankEnd)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
    Call UserList(Userindex).outgoingData.WriteASCIIString(Npclist(UserList(Userindex).flags.TargetNPC).name)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BankInit)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowBlacksmithForm(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowBlacksmithForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowBlacksmithForm)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowCarpenterForm(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowCarpenterForm)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteShowAlquimiaForm(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowAlquimiaForm)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteShowSastreForm(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowCarpenterForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowSastreForm)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "NPCKillUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCKillUser(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NPCKillUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.NPCKillUser)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldUser(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockedWithShieldUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldUser)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockedWithShieldOther(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockedWithShieldOther" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BlockedWithShieldOther)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CharSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharSwing(ByVal Userindex As Integer, ByVal CharIndex As Integer, Optional ByVal FX As Boolean = True, Optional ByVal ShowText As Boolean = True)

    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharSwing(CharIndex, FX, ShowText))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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
        
110         PrepareMessageCharSwing = .ReadASCIIStringFixed(.Length)

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

Public Sub WriteSafeModeOn(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeModeOn" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.SafeModeOn)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "SafeModeOff" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeModeOff(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeModeOff" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.SafeModeOff)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "PartySafeOn" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySafeOn(ByVal Userindex As Integer)

    '***************************************************
    'Author: Rapsodius
    'Last Modification: 10/10/07
    'Writes the "PartySafeOn" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.PartySafeOn)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "PartySafeOff" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePartySafeOff(ByVal Userindex As Integer)

    '***************************************************
    'Author: Rapsodius
    'Last Modification: 10/10/07
    'Writes the "PartySafeOff" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.PartySafeOff)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteClanSeguro(ByVal Userindex As Integer, ByVal estado As Boolean)

    '***************************************************
    'Author: Rapsodius
    'Last Modification: 10/10/07
    'Writes the "PartySafeOff" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ClanSeguro)
    Call UserList(Userindex).outgoingData.WriteBoolean(estado)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCantUseWhileMeditating(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CantUseWhileMeditating" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.CantUseWhileMeditating)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(Userindex).Stats.MinSta)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(Userindex).Stats.MinMAN)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal Userindex As Integer)

    'Call SendData(SendTarget.ToDiosesYclan, UserIndex, PrepareMessageCharUpdateHP(UserIndex))
    Call SendData(SendTarget.ToDiosesYclan, UserList(Userindex).GuildIndex, PrepareMessageCharUpdateHP(Userindex))

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(Userindex).Stats.MinHp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateGold" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(Userindex).Stats.GLD)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateExp" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(Userindex).Stats.Exp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteChangeMap(ByVal Userindex As Integer, ByVal Map As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMap" message to the given user's outgoing data buffer
    '***************************************************

    Dim Version As Integer

    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(Map)
        Call .WriteInteger(Version)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PosUpdate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteByte(UserList(Userindex).Pos.X)
        Call .WriteByte(UserList(Userindex).Pos.Y)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteNPCHitUser(ByVal Userindex As Integer, ByVal Target As PartesCuerpo, ByVal damage As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NPCHitUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.NPCHitUser)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserHitNPC" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    damage The number of HP lost by the target creature.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserHitNPC(ByVal Userindex As Integer, ByVal damage As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserHitNPC" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserHitNPC)
        
        'It is a long to allow the "drake slayer" (matadracos) to kill the great red dragon of one blow.
        Call .WriteLong(damage)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserAttackedSwing" message to the given user's outgoing data buffer.
'
' @param    UserIndex       User to which the message is intended.
' @param    attackerIndex   The user index of the user that attacked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserAttackedSwing(ByVal Userindex As Integer, ByVal attackerIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserAttackedSwing" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserAttackedSwing)
        Call .WriteInteger(UserList(attackerIndex).Char.CharIndex)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteUserHittedByUser(ByVal Userindex As Integer, ByVal Target As PartesCuerpo, ByVal attackerChar As Integer, ByVal damage As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserHittedByUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserHittedByUser)
        Call .WriteInteger(attackerChar)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteUserHittedUser(ByVal Userindex As Integer, ByVal Target As PartesCuerpo, ByVal attackedChar As Integer, ByVal damage As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserHittedUser" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserHittedUser)
        Call .WriteInteger(attackedChar)
        Call .WriteByte(Target)
        Call .WriteInteger(damage)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteChatOverHead(ByVal Userindex As Integer, ByVal chat As String, ByVal CharIndex As Integer, ByVal Color As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatOverHead" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(chat, CharIndex, Color))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteEfectOverHead(ByVal Userindex As Integer, ByVal chat As String, ByVal CharIndex As Integer, Optional ByVal Color As Long = &HFF0000)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatOverHead" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageEfectOverHead(chat, CharIndex, Color))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteExpOverHead(ByVal Userindex As Integer, ByVal chat As String, ByVal CharIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatOverHead" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageExpOverHead(chat, CharIndex))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteOroOverHead(ByVal Userindex As Integer, ByVal chat As String, ByVal CharIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatOverHead" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageOroOverHead(chat, CharIndex))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteConsoleMsg(ByVal Userindex As Integer, ByVal chat As String, ByVal FontIndex As FontTypeNames)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(chat, FontIndex))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteRenderValueMsg(ByVal Userindex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal rValue As Double, ByVal rType As Byte)

    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateRenderValue(X, Y, rValue, rType))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteLocaleMsg(ByVal Userindex As Integer, ByVal Id As Integer, ByVal FontIndex As FontTypeNames, Optional ByVal strExtra As String = vbNullString)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageLocaleMsg(Id, strExtra, FontIndex))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteListaCorreo(ByVal Userindex As Integer, ByVal Actualizar As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageListaCorreo(Userindex, Actualizar))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal Userindex As Integer, ByVal chat As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildChat" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(chat))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMessageBox(ByVal Userindex As Integer, ByVal message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(message)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteMostrarCuenta(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.MostrarCuenta)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(Userindex)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(Userindex).Char.CharIndex)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteCharacterCreate(ByVal Userindex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal name As String, ByVal Status As Byte, ByVal privileges As Byte, ByVal ParticulaFx As Byte, ByVal Head_Aura As String, ByVal Arma_Aura As String, ByVal Body_Aura As String, ByVal Anillo_Aura As String, ByVal Otra_Aura As String, ByVal Escudo_Aura As String, ByVal speeding As Single, ByVal EsNPC As Boolean, ByVal donador As Byte, ByVal appear As Byte, ByVal group_index As Integer, ByVal clan_index As Integer, ByVal clan_nivel As Byte, ByVal UserMinHp As Long, ByVal UserMaxHp As Long, ByVal Simbolo As Byte, Optional ByVal Idle As Boolean = False)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterCreate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(Body, Head, Heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, helmet, name, Status, privileges, ParticulaFx, Head_Aura, Arma_Aura, Body_Aura, Anillo_Aura, Otra_Aura, Escudo_Aura, speeding, EsNPC, donador, appear, group_index, clan_index, clan_nivel, UserMinHp, UserMaxHp, Simbolo, Idle))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal Desvanecido As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterRemove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex, Desvanecido))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteCharacterMove(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterMove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteForceCharMove(ByVal Userindex, ByVal Direccion As eHeading)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/03/2009
    'Writes the "ForceCharMove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(Userindex)
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

Public Sub WriteCharacterChange(ByVal Userindex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, Optional ByVal Idle As Boolean = False)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterChange" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(Body, Head, Heading, CharIndex, weapon, shield, FX, FXLoops, helmet, Idle))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteObjectCreate(ByVal Userindex As Integer, ByVal ObjIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectCreate" message to the given user's outgoing data buffer
    '***************************************************

    'If ObjIndex = 251 Then
    ' Debug.Print "Crear la puerta"
    'End If
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(ObjIndex, X, Y))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteParticleFloorCreate(ByVal Userindex As Integer, ByVal Particula As Integer, ByVal ParticulaTime As Integer, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)

    On Error GoTo ErrHandler
  
    If Particula = 0 Then Exit Sub
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageParticleFXToFloor(X, Y, Particula, ParticulaTime))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteLightFloorCreate(ByVal Userindex As Integer, ByVal LuzColor As Long, ByVal Rango As Byte, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)

    On Error GoTo ErrHandler
     
    MapData(Map, X, Y).Luz.Color = LuzColor
    MapData(Map, X, Y).Luz.Rango = Rango

    If Rango = 0 Then Exit Sub
    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageLightFXToFloor(X, Y, LuzColor, Rango))
    Exit Sub
    
ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteFxPiso(ByVal Userindex As Integer, ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageFxPiso(GrhIndex, X, Y))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteObjectDelete(ByVal Userindex As Integer, ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectDelete" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteBlockPosition(ByVal Userindex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockPosition" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteByte(X)
        Call .WriteByte(Y)
        Call .WriteByte(Blocked)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WritePlayMidi(ByVal Userindex As Integer, ByVal midi As Byte, Optional ByVal loops As Integer = -1)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PlayMidi" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMidi(midi, loops))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WritePlayWave(ByVal Userindex As Integer, ByVal wave As Integer, ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal Userindex As Integer, ByRef guildList() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim Tmp As String

    Dim i   As Long
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AreaChanged" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteByte(UserList(Userindex).Pos.X)
        Call .WriteByte(UserList(Userindex).Pos.Y)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PauseToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "RainToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageRainToggle())
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteNubesToggle(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageNieblandoToggle(IntensidadDeNubes))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteTrofeoToggleOn(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageTrofeoToggleOn())
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteTrofeoToggleOff(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageTrofeoToggleOff())
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteCreateFX(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal FX As Integer, ByVal FXLoops As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateFX" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal Userindex As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
    '***************************************************
    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharUpdateHP(UserIndex))
    Call SendData(SendTarget.ToDiosesYclan, UserList(Userindex).GuildIndex, PrepareMessageCharUpdateHP(Userindex))

    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(Userindex).Stats.MaxHp)
        Call .WriteInteger(UserList(Userindex).Stats.MinHp)
        Call .WriteInteger(UserList(Userindex).Stats.MaxMAN)
        Call .WriteInteger(UserList(Userindex).Stats.MinMAN)
        Call .WriteInteger(UserList(Userindex).Stats.MaxSta)
        Call .WriteInteger(UserList(Userindex).Stats.MinSta)
        Call .WriteLong(UserList(Userindex).Stats.GLD)
        Call .WriteByte(UserList(Userindex).Stats.ELV)
        Call .WriteLong(UserList(Userindex).Stats.ELU)
        Call .WriteLong(UserList(Userindex).Stats.Exp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteUpdateUserKey(ByVal Userindex As Integer, ByVal slot As Integer, ByVal Llave As Integer)
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserKey)
        Call .WriteInteger(slot)
        Call .WriteInteger(Llave)
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

' Actualiza el indicador de daño mágico
Public Sub WriteUpdateDM(ByVal Userindex As Integer)
    On Error GoTo ErrHandler
    
    Dim Valor As Integer
    
    With UserList(Userindex).Invent
        ' % daño mágico del arma
        If .WeaponEqpObjIndex > 0 Then
            Valor = Valor + ObjData(.WeaponEqpObjIndex).MagicDamageBonus
        End If
        ' % daño mágico del anillo
        If .AnilloEqpObjIndex > 0 Then
            Valor = Valor + ObjData(.AnilloEqpObjIndex).MagicDamageBonus
        End If
    End With

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateDM)
        Call .WriteInteger(Valor)
    End With

    Exit Sub

ErrHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Resume
    End If
End Sub

' Actualiza el indicador de resistencia mágica
Public Sub WriteUpdateRM(ByVal Userindex As Integer)
    On Error GoTo ErrHandler
    
    Dim Valor As Integer
    
    With UserList(Userindex).Invent
        ' Resistencia mágica de la armadura
        If .ArmourEqpObjIndex > 0 Then
            Valor = Valor + ObjData(.ArmourEqpObjIndex).ResistenciaMagica
        End If
        
        ' Resistencia mágica del anillo
        If .AnilloEqpObjIndex > 0 Then
            Valor = Valor + ObjData(.AnilloEqpObjIndex).ResistenciaMagica
        End If
        
        ' Resistencia mágica del escudo
        If .EscudoEqpObjIndex > 0 Then
            Valor = Valor + ObjData(.EscudoEqpObjIndex).ResistenciaMagica
        End If
        
        ' Resistencia mágica del casco
        If .CascoEqpObjIndex > 0 Then
            Valor = Valor + ObjData(.CascoEqpObjIndex).ResistenciaMagica
        End If
    End With

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateRM)
        Call .WriteInteger(Valor)
    End With

    Exit Sub

ErrHandler:
    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        Resume
    End If
End Sub

''
' Writes the "WorkRequestTarget" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Skill The skill for which we request a target.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkRequestTarget(ByVal Userindex As Integer, ByVal Skill As eSkill)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WorkRequestTarget" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.WorkRequestTarget)
        Call .WriteByte(Skill)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

' Writes the "InventoryUnlockSlots" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInventoryUnlockSlots(ByVal Userindex As Integer)

    '***************************************************
    'Author: Ruthnar
    'Last Modification: 30/09/20
    'Writes the "WriteInventoryUnlockSlots" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.InventoryUnlockSlots)
        Call .WriteByte(UserList(Userindex).Stats.InventLevel)
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteIntervals(ByVal Userindex As Integer)

    On Error GoTo ErrHandler

    With UserList(Userindex)
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

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteChangeInventorySlot(ByVal Userindex As Integer, ByVal slot As Byte)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 3/12/09
    'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
    '3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
    '***************************************************

    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(slot)
        
        Dim ObjIndex As Integer
        
        ObjIndex = UserList(Userindex).Invent.Object(slot).ObjIndex
        
        Dim PodraUsarlo As Byte
    
        'Ladder
        If ObjIndex > 0 Then
            PodraUsarlo = PuedeUsarObjeto(Userindex, ObjIndex)
            'PodraUsarlo = IIf(SexoPuedeUsarItem(UserIndex, OBJIndex) = True And UserList(UserIndex).Stats.ELV >= ObjData(OBJIndex).MinELV And ClasePuedeUsarItem(UserIndex, OBJIndex) = True And CheckRazaUsaRopa(UserIndex, OBJIndex) = True, 1, 0)
            'Ladder
    
        End If
    
        Call .WriteInteger(ObjIndex)
        Call .WriteInteger(UserList(Userindex).Invent.Object(slot).Amount)
        Call .WriteBoolean(UserList(Userindex).Invent.Object(slot).Equipped)
        Call .WriteSingle(SalePrice(ObjIndex))
        Call .WriteByte(PodraUsarlo)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal Userindex As Integer, ByVal slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(slot)
        
        Dim ObjIndex As Integer

        Dim obData   As ObjData
        
        ObjIndex = UserList(Userindex).BancoInvent.Object(slot).ObjIndex
        
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)

        End If
        
        Dim PodraUsarlo As Byte
    
        'Ladder
        If ObjIndex > 0 Then
            PodraUsarlo = PuedeUsarObjeto(Userindex, ObjIndex)

            'PodraUsarlo = IIf(SexoPuedeUsarItem(UserIndex, OBJIndex) = True And UserList(UserIndex).Stats.ELV >= ObjData(OBJIndex).MinELV = True And ClasePuedeUsarItem(UserIndex, OBJIndex) = True And CheckRazaUsaRopa(UserIndex, OBJIndex) = True, 1, 0)
            'Ladder
        End If

        Call .WriteInteger(UserList(Userindex).BancoInvent.Object(slot).Amount)
        Call .WriteLong(obData.Valor)
        Call .WriteByte(PodraUsarlo)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal Userindex As Integer, ByVal slot As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(slot)
        Call .WriteInteger(UserList(Userindex).Stats.UserHechizos(slot))
        
        If UserList(Userindex).Stats.UserHechizos(slot) > 0 Then
            Call .WriteByte(UserList(Userindex).Stats.UserHechizos(slot))
        Else
            Call .WriteByte("255")

        End If

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Atributes" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.Atributes)
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Constitucion))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma))
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithWeapons(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
    'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ArmasHerrero()))
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithWeapons)
        
        For i = 1 To UBound(ArmasHerrero())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmasHerrero(i)).SkHerreria <= UserList(Userindex).Stats.UserSkills(eSkill.Herreria) Then
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

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BlacksmithArmors" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlacksmithArmors(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
    'Writes the "BlacksmithArmors" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ArmadurasHerrero()))
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.BlacksmithArmors)
        
        For i = 1 To UBound(ArmadurasHerrero())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ArmadurasHerrero(i)).SkHerreria <= Round(UserList(Userindex).Stats.UserSkills(eSkill.Herreria) / ModHerreria(UserList(Userindex).clase), 0) Then
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

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "CarpenterObjects" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCarpenterObjects(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CarpenterObjects" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i              As Long

    'Dim obj As ObjData
    Dim validIndexes() As Integer

    Dim Count          As Byte
    
    ReDim validIndexes(1 To UBound(ObjCarpintero()))
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.CarpenterObjects)
        
        For i = 1 To UBound(ObjCarpintero())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjCarpintero(i)).SkCarpinteria <= UserList(Userindex).Stats.UserSkills(eSkill.Carpinteria) Then
                If i = 1 Then Debug.Print UserList(Userindex).Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(UserList(Userindex).clase)
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

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteAlquimistaObjects(ByVal Userindex As Integer)

    '***************************************************
    On Error GoTo ErrHandler

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ObjAlquimista()))
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AlquimistaObj)
        
        For i = 1 To UBound(ObjAlquimista())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjAlquimista(i)).SkPociones <= UserList(Userindex).Stats.UserSkills(eSkill.Alquimia) \ ModAlquimia(UserList(Userindex).clase) Then
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

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteSastreObjects(ByVal Userindex As Integer)

    '***************************************************
    On Error GoTo ErrHandler

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ObjSastre()))
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.SastreObj)
        
        For i = 1 To UBound(ObjSastre())

            ' Can the user create this object? If so add it to the list....
            If ObjData(ObjSastre(i)).SkMAGOria <= UserList(Userindex).Stats.UserSkills(eSkill.Sastreria) Then

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

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RestOK" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.RestOK)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteErrorMsg(ByVal Userindex As Integer, ByVal message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ErrorMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(message))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Blind" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Blind)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Dumb" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Dumb)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteShowSignal(ByVal Userindex As Integer, ByVal ObjIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowSignal" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteInteger(ObjIndex)
        Call .WriteInteger(ObjData(ObjIndex).GrhSecundario)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteChangeNPCInventorySlot(ByVal Userindex As Integer, ByVal slot As Byte, ByRef obj As obj, ByVal price As Single)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Last Modified by: Nicolas Ezequiel Bouhid (NicoNZ)
    'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim ObjInfo As ObjData
    
    If obj.ObjIndex >= LBound(ObjData()) And obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(obj.ObjIndex)

    End If
    
    Dim PodraUsarlo As Byte
    
    'Ladder
    If obj.ObjIndex > 0 Then
        PodraUsarlo = PuedeUsarObjeto(Userindex, obj.ObjIndex)

        'PodraUsarlo = IIf(SexoPuedeUsarItem(UserIndex, obj.OBJIndex) = True And UserList(UserIndex).Stats.ELV >= ObjData(obj.OBJIndex).MinELV And ClasePuedeUsarItem(UserIndex, obj.OBJIndex) = True And CheckRazaUsaRopa(UserIndex, obj.OBJIndex) = True, 1, 0)
        'Ladder
    End If
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(slot)
        Call .WriteInteger(obj.ObjIndex)
        Call .WriteInteger(obj.Amount)
        Call .WriteSingle(price)
        Call .WriteByte(PodraUsarlo)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(Userindex).Stats.MaxAGU)
        Call .WriteByte(UserList(Userindex).Stats.MinAGU)
        Call .WriteByte(UserList(Userindex).Stats.MaxHam)
        Call .WriteByte(UserList(Userindex).Stats.MinHam)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteLight(ByVal Userindex As Integer, ByVal Map As Integer)

    On Error GoTo ErrHandler

    Dim light As String
 
    light = MapInfo(Map).base_light

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.light)
        Call .WriteASCIIString(light)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteEfectToScreen(ByVal Userindex As Integer, ByVal Color As Long, ByVal Time As Long, Optional ByVal Ignorar As Boolean = False)

    On Error GoTo ErrHandler
 
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.EfectToScreen)
        Call .WriteLong(Color)
        Call .WriteLong(Time)
        Call .WriteBoolean(Ignorar)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteFYA(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.FYA)
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(1))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(2))
        Call .WriteInteger(UserList(Userindex).flags.DuracionEfecto)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteCerrarleCliente(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.CerrarleCliente)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteOxigeno(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.Oxigeno)
        Call .WriteInteger(UserList(Userindex).Counters.Oxigeno)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteContadores(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.Contadores)
        Call .WriteInteger(UserList(Userindex).Counters.Invisibilidad)
        Call .WriteInteger(UserList(Userindex).Counters.ScrollExperiencia)
        Call .WriteInteger(UserList(Userindex).Counters.ScrollOro)

        If UserList(Userindex).flags.NecesitaOxigeno Then
            Call .WriteInteger(UserList(Userindex).Counters.Oxigeno)
        Else
            Call .WriteInteger(0)

        End If

        Call .WriteInteger(UserList(Userindex).flags.DuracionEfecto)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteBindKeys(ByVal Userindex As Integer)

    '***************************************************
    'Envia los macros al cliente!
    'Por Ladder
    '23/09/2014
    'Flor te amo!
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.BindKeys)
        Call .WriteByte(UserList(Userindex).ChatCombate)
        Call .WriteByte(UserList(Userindex).ChatGlobal)
        
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MiniStats" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        
        Call .WriteLong(UserList(Userindex).Faccion.CiudadanosMatados)
        Call .WriteLong(UserList(Userindex).Faccion.CriminalesMatados)
        Call .WriteByte(UserList(Userindex).Faccion.Status)
        
        'TODO : Este valor es calculable, no debería NI EXISTIR, ya sea en el servidor ni en el cliente!!!
        'Call .WriteLong(UserList(UserIndex).Stats.UsuariosMatados)
        
        Call .WriteInteger(UserList(Userindex).Stats.NPCsMuertos)
        
        Call .WriteByte(UserList(Userindex).clase)
        Call .WriteLong(UserList(Userindex).Counters.Pena)
        
        'Ladder 31/07/08  Envio mas estadisticas :P
        Call .WriteLong(UserList(Userindex).flags.VecesQueMoriste)
        Call .WriteByte(UserList(Userindex).genero)
        Call .WriteByte(UserList(Userindex).raza)
        
        Call .WriteByte(UserList(Userindex).donador.activo)
        Call .WriteLong(CreditosDonadorCheck(UserList(Userindex).Cuenta))
        'ARREGLANDO
        
        Call .WriteInteger(DiasDonadorCheck(UserList(Userindex).Cuenta))
        
        Call .WriteLong(UserList(Userindex).flags.BattlePuntos)
                
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "LevelUp" message to the given user's outgoing data buffer.
'
' @param    skillPoints The number of free skill points the player has.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLevelUp(ByVal Userindex As Integer, ByVal skillPoints As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LevelUp" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.LevelUp)
        Call .WriteInteger(skillPoints)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data buffer.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal Userindex As Integer, ByVal title As String, ByVal message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AddForumMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AddForumMsg)
        Call .WriteASCIIString(title)
        Call .WriteASCIIString(message)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowForumForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowForumForm)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteSetInvisible(ByVal Userindex As Integer, ByVal CharIndex As Integer, ByVal invisible As Boolean)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetInvisible" message to the given user's outgoing data buffer
    '***************************************************

    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteDiceRoll(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DiceRoll" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.DiceRoll)
        ' TODO: SACAR ESTE PAQUETE USAR EL DE ATRIBUTOS
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Constitucion))
        Call .WriteByte(UserList(Userindex).Stats.UserAtributos(eAtributos.Carisma))
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MeditateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlindNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DumbNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SendSkills" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i As Long
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.SendSkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(UserList(Userindex).Stats.UserSkills(i))
        Next i

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal Userindex As Integer, ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long

    Dim str As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To Npclist(NpcIndex).NroCriaturas
            str = str & Npclist(NpcIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then str = Left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteGuildNews(ByVal Userindex As Integer, ByVal guildNews As String, ByRef guildList() As String, ByRef MemberList() As String, ByVal ClanNivel As Byte, ByVal ExpAcu As Integer, ByVal ExpNe As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildNews" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
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

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal Userindex As Integer, ByVal details As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OfferDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i As Long
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)
        
        Call .WriteASCIIString(details)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal Userindex As Integer, ByRef guilds() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)
        
        ' Prepare guild's list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal Userindex As Integer, ByRef guilds() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)
                
        ' Prepare guilds' list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteCharacterInfo(ByVal Userindex As Integer, ByVal CharName As String, ByVal race As eRaza, ByVal Class As eClass, ByVal gender As eGenero, ByVal level As Byte, ByVal gold As Long, ByVal bank As Long, ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
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

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteGuildLeaderInfo(ByVal Userindex As Integer, ByRef guildList() As String, ByRef MemberList() As String, ByVal guildNews As String, ByRef joinRequests() As String, ByVal NivelDeClan As Byte, ByVal ExpActual As Integer, ByVal ExpNecesaria As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
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

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteGuildDetails(ByVal Userindex As Integer, ByVal GuildName As String, ByVal founder As String, ByVal foundationDate As String, ByVal leader As String, ByVal memberCount As Integer, ByVal alignment As String, ByVal guildDesc As String, ByVal NivelDeClan As Byte, ByVal ExpActual As Integer, ByVal ExpNecesaria As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i    As Long

    Dim temp As String
    
    With UserList(Userindex).outgoingData
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

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/12/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    'Writes the "ParalizeOK" message to the given user's outgoing data buffer
    'And updates user position
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ParalizeOK)

    
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteInmovilizaOK(ByVal Userindex As Integer)

    '***************************************************
    'Inmovilizar
    'Por Ladder
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.InmovilizadoOK)
    '  Call WritePosUpdate(UserIndex)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal Userindex As Integer, ByVal details As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        
        Call .WriteASCIIString(details)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteChangeUserTradeSlot(ByVal Userindex As Integer, ByVal ObjIndex As Integer, ByVal Amount As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
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

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal Userindex As Integer, ByRef npcNames() As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & i & SEPARATOR
            
        Next i
     
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMOTDEditionForm(ByVal Userindex As Integer, ByVal currentMOTD As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMOTDEditionForm)
        
        Call .WriteASCIIString(currentMOTD)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowGMPanelForm)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteShowFundarClanForm(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.ShowFundarClanForm)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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

Public Sub WriteUserNameList(ByVal Userindex As Integer, ByRef userNamesList() As String, ByVal cant As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06 NIGO:
    'Writes the "UserNameList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal Userindex As Integer, ByVal Time As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Pong" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.Pong)
    Call UserList(Userindex).outgoingData.WriteLong(Time)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal Userindex As Integer)
        
        On Error GoTo FlushBuffer_Err
        

        '***************************************************
        'Sends all data existing in the buffer
        '***************************************************
    
100     With UserList(Userindex).outgoingData

102         If .Length = 0 Then Exit Sub
        
            ' Tratamos de enviar los datos.
104         Dim ret As Long: ret = WsApiEnviar(Userindex, .ReadASCIIStringFixed(.Length))
    
            ' Si recibimos un error como respuesta de la API, cerramos el socket.
106         If ret <> 0 And ret <> WSAEWOULDBLOCK Then
                ' Close the socket avoiding any critical error
108             Call CloseSocketSL(Userindex)
110             Call Cerrar_Usuario(Userindex)
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
        
108         PrepareMessageSetInvisible = .ReadASCIIStringFixed(.Length)

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
        
108         PrepareMessageSetEscribiendo = .ReadASCIIStringFixed(.Length)

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
        
122         PrepareMessageChatOverHead = .ReadASCIIStringFixed(.Length)

        End With

        
        Exit Function

PrepareMessageChatOverHead_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageChatOverHead", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageEfectOverHead(ByVal chat As String, ByVal CharIndex As Integer, Optional ByVal Color As Long = vbRed) As String
        
        On Error GoTo PrepareMessageEfectOverHead_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ChatOverHead" message and returns it.
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.EfectOverHEad)
104         Call .WriteASCIIString(chat)
106         Call .WriteInteger(CharIndex)
108         Call .WriteLong(Color)
110         PrepareMessageEfectOverHead = .ReadASCIIStringFixed(.Length)

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
108         PrepareMessageExpOverHead = .ReadASCIIStringFixed(.Length)

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
108         PrepareMessageOroOverHead = .ReadASCIIStringFixed(.Length)

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
        
108         PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.Length)

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
        
110         PrepareMessageLocaleMsg = .ReadASCIIStringFixed(.Length)

        End With

        
        Exit Function

PrepareMessageLocaleMsg_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageLocaleMsg", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageListaCorreo(ByVal Userindex As Integer, ByVal Actualizar As Boolean) As String
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ConsoleMsg" message and returns it.
        '***************************************************
        
        On Error GoTo PrepareMessageListaCorreo_Err
        

        Dim cant As Byte

        Dim i    As Byte

100     cant = UserList(Userindex).Correo.CantCorreo
102     UserList(Userindex).Correo.NoLeidos = 0

104     With auxiliarBuffer
106         Call .WriteByte(ServerPacketID.ListaCorreo)
108         Call .WriteByte(cant)

110         If cant > 0 Then

112             For i = 1 To cant
114                 Call .WriteASCIIString(UserList(Userindex).Correo.Mensaje(i).Remitente)
116                 Call .WriteASCIIString(UserList(Userindex).Correo.Mensaje(i).Mensaje)
118                 Call .WriteByte(UserList(Userindex).Correo.Mensaje(i).ItemCount)
120                 Call .WriteASCIIString(UserList(Userindex).Correo.Mensaje(i).Item)

122                 Call .WriteByte(UserList(Userindex).Correo.Mensaje(i).Leido)
124                 Call .WriteASCIIString(UserList(Userindex).Correo.Mensaje(i).Fecha)
                    'Call ReadMessageCorreo(UserIndex, i)
126             Next i

            End If

128         Call .WriteBoolean(Actualizar)
        
130         PrepareMessageListaCorreo = .ReadASCIIStringFixed(.Length)

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
        
110         PrepareMessageCreateFX = .ReadASCIIStringFixed(.Length)

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
        
108         PrepareMessageMeditateToggle = .ReadASCIIStringFixed(.Length)
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
        
112         PrepareMessageParticleFX = .ReadASCIIStringFixed(.Length)

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
        
118         PrepareMessageParticleFXWithDestino = .ReadASCIIStringFixed(.Length)

        End With

        
        Exit Function

PrepareMessageParticleFXWithDestino_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageParticleFXWithDestino", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageParticleFXWithDestinoXY(ByVal Emisor As Integer, ByVal ParticulaViaje As Integer, ByVal ParticulaFinal As Integer, ByVal Time As Long, ByVal wav As Integer, ByVal FX As Integer, ByVal X As Byte, ByVal Y As Byte) As String
        
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
116         Call .WriteByte(X)
118         Call .WriteByte(Y)
        
120         PrepareMessageParticleFXWithDestinoXY = .ReadASCIIStringFixed(.Length)

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
112         PrepareMessageAuraToChar = .ReadASCIIStringFixed(.Length)

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
102         Call .WriteByte(ServerPacketID.SpeedToChar)
104         Call .WriteInteger(CharIndex)
106         Call .WriteSingle(speeding)
108         PrepareMessageSpeedingACT = .ReadASCIIStringFixed(.Length)

        End With

        
        Exit Function

PrepareMessageSpeedingACT_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageSpeedingACT", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageParticleFXToFloor(ByVal X As Byte, ByVal Y As Byte, ByVal Particula As Integer, ByVal Time As Long) As String
        
        On Error GoTo PrepareMessageParticleFXToFloor_Err
        

        '***************************************************
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ParticleFXToFloor)
104         Call .WriteByte(X)
106         Call .WriteByte(Y)
108         Call .WriteInteger(Particula)
110         Call .WriteLong(Time)
112         PrepareMessageParticleFXToFloor = .ReadASCIIStringFixed(.Length)

        End With

        
        Exit Function

PrepareMessageParticleFXToFloor_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageParticleFXToFloor", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageLightFXToFloor(ByVal X As Byte, ByVal Y As Byte, ByVal LuzColor As Long, ByVal Rango As Byte) As String
        
        On Error GoTo PrepareMessageLightFXToFloor_Err
        

        '***************************************************
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.LightToFloor)
104         Call .WriteByte(X)
106         Call .WriteByte(Y)
108         Call .WriteLong(LuzColor)
110         Call .WriteByte(Rango)
112         PrepareMessageLightFXToFloor = .ReadASCIIStringFixed(.Length)

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

Public Function PrepareMessagePlayWave(ByVal wave As Integer, ByVal X As Byte, ByVal Y As Byte) As String
        
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
106         Call .WriteByte(X)
108         Call .WriteByte(Y)
        
110         PrepareMessagePlayWave = .ReadASCIIStringFixed(.Length)

        End With

        
        Exit Function

PrepareMessagePlayWave_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessagePlayWave", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageUbicacionLlamada(ByVal Mapa As Integer, ByVal X As Byte, ByVal Y As Byte) As String
        
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
106         Call .WriteByte(X)
108         Call .WriteByte(Y)
        
110         PrepareMessageUbicacionLlamada = .ReadASCIIStringFixed(.Length)

        End With

        
        Exit Function

PrepareMessageUbicacionLlamada_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageUbicacionLlamada", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageCharUpdateHP(ByVal Userindex As Integer) As String
        
        On Error GoTo PrepareMessageCharUpdateHP_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 08/08/07
        'Last Modified by: Rapsodius
        'Added X and Y positions for 3D Sounds
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.CharUpdateHP)
104         Call .WriteInteger(UserList(Userindex).Char.CharIndex)
106         Call .WriteInteger(UserList(Userindex).Stats.MinHp)
108         Call .WriteInteger(UserList(Userindex).Stats.MaxHp)
        
110         PrepareMessageCharUpdateHP = .ReadASCIIStringFixed(.Length)

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
        
106         PrepareMessageArmaMov = .ReadASCIIStringFixed(.Length)

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
        
106         PrepareMessageEscudoMov = .ReadASCIIStringFixed(.Length)

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
110         PrepareMessageEfectToScreen = .ReadASCIIStringFixed(.Length)

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
        
106         PrepareMessageGuildChat = .ReadASCIIStringFixed(.Length)

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
        
106         PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.Length)

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
102         Call .WriteByte(ServerPacketID.PlayMIDI)
104         Call .WriteByte(midi)
106         Call .WriteInteger(loops)
        
108         PrepareMessagePlayMidi = .ReadASCIIStringFixed(.Length)

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
        
106         PrepareMessageOnlineUser = .ReadASCIIStringFixed(.Length)

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
104         PrepareMessagePauseToggle = .ReadASCIIStringFixed(.Length)

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
        
104         PrepareMessageRainToggle = .ReadASCIIStringFixed(.Length)

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
        
104         PrepareMessageTrofeoToggleOn = .ReadASCIIStringFixed(.Length)

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
102         Call .WriteByte(ServerPacketID.TrofeoToggleOff)
        
104         PrepareMessageTrofeoToggleOff = .ReadASCIIStringFixed(.Length)

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
104         Call .WriteLong(((timeGetTime And &H7FFFFFFF) - HoraMundo) Mod DuracionDia)
106         Call .WriteLong(DuracionDia)
        
108         PrepareMessageHora = .ReadASCIIStringFixed(.Length)

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

Public Function PrepareMessageObjectDelete(ByVal X As Byte, ByVal Y As Byte) As String
        
        On Error GoTo PrepareMessageObjectDelete_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ObjectDelete" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ObjectDelete)
104         Call .WriteByte(X)
106         Call .WriteByte(Y)
        
108         PrepareMessageObjectDelete = .ReadASCIIStringFixed(.Length)

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

Public Function PrepareMessageBlockPosition(ByVal X As Byte, ByVal Y As Byte, ByVal Blocked As Byte) As String
        
        On Error GoTo PrepareMessageBlockPosition_Err
        

        '***************************************************
        'Author: Fredy Horacio Treboux (liquid)
        'Last Modification: 01/08/07
        'Prepares the "BlockPosition" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.BlockPosition)
104         Call .WriteByte(X)
106         Call .WriteByte(Y)
108         Call .WriteByte(Blocked)
        
110         PrepareMessageBlockPosition = .ReadASCIIStringFixed(.Length)

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
Public Function PrepareMessageObjectCreate(ByVal ObjIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String
        
        On Error GoTo PrepareMessageObjectCreate_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'prepares the "ObjectCreate" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ObjectCreate)
104         Call .WriteByte(X)
106         Call .WriteByte(Y)
108         Call .WriteInteger(ObjIndex)
        
110         PrepareMessageObjectCreate = .ReadASCIIStringFixed(.Length)

        End With

        
        Exit Function

PrepareMessageObjectCreate_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageObjectCreate", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageFxPiso(ByVal GrhIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String
        
        On Error GoTo PrepareMessageFxPiso_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'prepares the "ObjectCreate" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.fxpiso)
104         Call .WriteByte(X)
106         Call .WriteByte(Y)
108         Call .WriteInteger(GrhIndex)
        
110         PrepareMessageFxPiso = .ReadASCIIStringFixed(.Length)

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
        
108         PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.Length)

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
        
106         PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.Length)

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

Public Function PrepareMessageCharacterCreate(ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal name As String, ByVal Status As Byte, ByVal privileges As Byte, ByVal ParticulaFx As Byte, ByVal Head_Aura As String, ByVal Arma_Aura As String, ByVal Body_Aura As String, ByVal Anillo_Aura As String, ByVal Otra_Aura As String, ByVal Escudo_Aura As String, ByVal speeding As Single, ByVal EsNPC As Boolean, ByVal donador As Byte, ByVal appear As Byte, ByVal group_index As Integer, ByVal clan_index As Integer, ByVal clan_nivel As Byte, ByVal UserMinHp As Long, ByVal UserMaxHp As Long, ByVal Simbolo As Byte, ByVal Idle As Boolean) As String
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
110         Call .WriteByte(Heading)
112         Call .WriteByte(X)
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
139         Call .WriteASCIIString(Anillo_Aura)
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
163         Call .WriteBoolean(Idle)

164         PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.Length)

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

Public Function PrepareMessageCharacterChange(ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading, ByVal CharIndex As Integer, ByVal weapon As Integer, ByVal shield As Integer, ByVal FX As Integer, ByVal FXLoops As Integer, ByVal helmet As Integer, ByVal Idle As Boolean) As String
        
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
110         Call .WriteByte(Heading)
112         Call .WriteInteger(weapon)
114         Call .WriteInteger(shield)
116         Call .WriteInteger(helmet)
118         Call .WriteInteger(FX)
120         Call .WriteInteger(FXLoops)
121         Call .WriteBoolean(Idle)
        
122         PrepareMessageCharacterChange = .ReadASCIIStringFixed(.Length)

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

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, ByVal X As Byte, ByVal Y As Byte) As String
        
        On Error GoTo PrepareMessageCharacterMove_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "CharacterMove" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.CharacterMove)
104         Call .WriteInteger(CharIndex)
106         Call .WriteByte(X)
108         Call .WriteByte(Y)
        
110         PrepareMessageCharacterMove = .ReadASCIIStringFixed(.Length)

        End With

        
        Exit Function

PrepareMessageCharacterMove_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCharacterMove", Erl)
        Resume Next
        
End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String

    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/03/2009
    'Prepares the "ForceCharMove" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ForceCharMove)
        Call .WriteByte(Direccion)
        
        PrepareMessageForceCharMove = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal Userindex As Integer, Status As Byte, Tag As String) As String
        
        On Error GoTo PrepareMessageUpdateTagAndStatus_Err
        

        '***************************************************
        'Author: Alejandro Salvo (Salvito)
        'Last Modification: 04/07/07
        'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
        'Prepares the "UpdateTagAndStatus" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        
104         Call .WriteInteger(UserList(Userindex).Char.CharIndex)
106         Call .WriteByte(Status)
108         Call .WriteASCIIString(Tag)
110         Call .WriteInteger(UserList(Userindex).Grupo.Lider)
        
112         PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.Length)

        End With

        
        Exit Function

PrepareMessageUpdateTagAndStatus_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageUpdateTagAndStatus", Erl)
        Resume Next
        
End Function

Public Sub WriteUpdateNPCSimbolo(ByVal Userindex As Integer, ByVal NpcIndex As Integer, ByVal Simbolo As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateNPCSimbolo)
        Call .WriteInteger(Npclist(NpcIndex).Char.CharIndex)
        Call .WriteByte(Simbolo)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal message As String) As String
        
        On Error GoTo PrepareMessageErrorMsg_Err
        

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Prepares the "ErrorMsg" message and returns it
        '***************************************************
100     With auxiliarBuffer
102         Call .WriteByte(ServerPacketID.ErrorMsg)
104         Call .WriteASCIIString(message)
        
106         PrepareMessageErrorMsg = .ReadASCIIStringFixed(.Length)

        End With

        
        Exit Function

PrepareMessageErrorMsg_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageErrorMsg", Erl)
        Resume Next
        
End Function

Private Sub HandleQuestionGM(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim Consulta       As String

        Dim TipoDeConsulta As String

        Consulta = buffer.ReadASCIIString()
        TipoDeConsulta = buffer.ReadASCIIString()

        If UserList(Userindex).donador.activo = 1 Then
            Call Ayuda.Push(.name, Consulta, TipoDeConsulta & "-Prioritario")
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Se ha recibido un nuevo mensaje de soporte de " & UserList(Userindex).name & "(Prioritario).", FontTypeNames.FONTTYPE_SERVER))
            
        Else
            Call Ayuda.Push(.name, Consulta, TipoDeConsulta)
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Se ha recibido un nuevo mensaje de soporte de " & UserList(Userindex).name & ".", FontTypeNames.FONTTYPE_SERVER))

        End If

        Call WriteConsoleMsg(Userindex, "Tu mensaje fue recibido por el equipo de soporte.", FontTypeNames.FONTTYPE_INFOIAO)
        'Call WriteConsoleMsg(UserIndex, "Tu mensaje fue recibido por el equipo de soporte.", FontTypeNames.FONTTYPE_INFOIAO)
        
        Call LogConsulta(.name & "(" & TipoDeConsulta & ") " & Consulta)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleOfertaInicial(ByVal Userindex As Integer)
        
        On Error GoTo HandleOfertaInicial_Err
        

        'Author: Pablo Mercavides
100     If UserList(Userindex).incomingData.Length < 6 Then
102         Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
            Exit Sub

        End If
    
104     With UserList(Userindex)
            'Remove packet ID
106         Call .incomingData.ReadInteger

            Dim Oferta As Long

108         Oferta = .incomingData.ReadLong()
        
110         If UserList(Userindex).flags.Muerto = 1 Then
112             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
                
                Exit Sub

            End If

114         If .flags.TargetNPC < 1 Then
116             Call WriteConsoleMsg(Userindex, "Primero tenés que hacer click sobre el subastador.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

118         If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Subastador Then
120             Call WriteConsoleMsg(Userindex, "Primero tenés que hacer click sobre el subastador.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
122         If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 2 Then
124             Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos del subastador.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
126         If .flags.Subastando = False Then
128             Call WriteChatOverHead(Userindex, "Ollí amigo, tu no podés decirme cual es la oferta inicial.", Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, vbWhite)
                Exit Sub

            End If
        
130         If Subasta.HaySubastaActiva = False And .flags.Subastando = False Then
132             Call WriteConsoleMsg(Userindex, "No hay ninguna subasta en curso.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
134         If .flags.Subastando = True Then
136             UserList(Userindex).Counters.TiempoParaSubastar = 0
138             Subasta.OfertaInicial = Oferta
140             Subasta.MejorOferta = 0
142             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " está subastando: " & ObjData(Subasta.ObjSubastado).name & " (Cantidad: " & Subasta.ObjSubastadoCantidad & " ) - con un precio inicial de " & Subasta.OfertaInicial & " monedas. Escribe /OFERTAR (cantidad) para participar.", FontTypeNames.FONTTYPE_SUBASTA))
144             .flags.Subastando = False
146             Subasta.HaySubastaActiva = True
148             Subasta.Subastador = .name
150             Subasta.MinutosDeSubasta = 5
152             Subasta.TiempoRestanteSubasta = 300
154             Call LogearEventoDeSubasta("#################################################################################################################################################################################################")
156             Call LogearEventoDeSubasta("El dia: " & Date & " a las " & Time)
158             Call LogearEventoDeSubasta(.name & ": Esta subastando el item numero " & Subasta.ObjSubastado & " con una cantidad de " & Subasta.ObjSubastadoCantidad & " y con un precio inicial de " & Subasta.OfertaInicial & " monedas.")
160             frmMain.SubastaTimer.Enabled = True
162             Call WarpUserChar(Userindex, 14, 27, 64, True)

                'lalala toda la bola de los timerrr
            End If

        End With

        
        Exit Sub

HandleOfertaInicial_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleOfertaInicial", Erl)
        Resume Next
        
End Sub

Private Sub HandleOfertaDeSubasta(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Dim Oferta   As Long

        Dim ExOferta As Long
        
        Oferta = buffer.ReadLong()
        
        If Subasta.HaySubastaActiva = False Then
            Call WriteConsoleMsg(Userindex, "No hay ninguna subasta en curso.", FontTypeNames.FONTTYPE_INFOIAO)
            Call .incomingData.CopyBuffer(buffer)
            Exit Sub

        End If
        
        If UserList(Userindex).flags.BattleModo = 1 Then
            Call WriteConsoleMsg(Userindex, "Subastador > íComo vas a ofertar con dinero que no es tuyo? Bríbon.", FontTypeNames.FONTTYPE_INFOIAO)
            Call .incomingData.CopyBuffer(buffer)
            Exit Sub

        End If
        
        If Oferta < Subasta.MejorOferta + 100 Then
            Call WriteConsoleMsg(Userindex, "Debe haber almenos una diferencia de 100 monedas a la ultima oferta!", FontTypeNames.FONTTYPE_INFOIAO)
            Call .incomingData.CopyBuffer(buffer)
            Exit Sub

        End If
        
        If .name = Subasta.Subastador Then
            Call WriteConsoleMsg(Userindex, "No podés auto ofertar en tus subastas. La proxima vez iras a la carcel...", FontTypeNames.FONTTYPE_INFOIAO)
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
            Call WriteUpdateGold(Userindex)
            
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
            Call WriteConsoleMsg(Userindex, "No posees esa cantidad de oro.", FontTypeNames.FONTTYPE_INFOIAO)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleGlobalMessage(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadByte
        
        Dim chat As String

        chat = buffer.ReadASCIIString()

        If .flags.Silenciado = 1 Then
            Call WriteLocaleMsg(Userindex, "110", FontTypeNames.FONTTYPE_VENENO, .flags.MinutosRestantes)
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
                Call WriteConsoleMsg(Userindex, "El global se encuentra Desactivado.", FontTypeNames.FONTTYPE_GLOBAL)

            End If

        End If
    
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub HandleGlobalOnOff(ByVal Userindex As Integer)
        
        On Error GoTo HandleGlobalOnOff_Err
        

        'Author: Pablo Mercavides
100     With UserList(Userindex)
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

Private Sub HandleCrearCuenta(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 18 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim CuentaEmail    As String

    Dim CuentaPassword As String
    
    CuentaEmail = buffer.ReadASCIIString()
    CuentaPassword = buffer.ReadASCIIString()
  
    If Not CheckMailString(CuentaEmail) Then
        Call WriteErrorMsg(Userindex, "Email inválido.")
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If

    If Not CuentaExiste(CuentaEmail) Then

        Call SaveNewAccount(Userindex, CuentaEmail, SDesencriptar(CuentaPassword))
    
        Call EnviarCorreo(CuentaEmail)
        Call WriteShowFrmLogear(Userindex)
        Call WriteShowMessageBox(Userindex, "Cuenta creada. Se ha enviado un código de validación a su email, debe activar la cuenta antes de poder usarla. Recuerde revisar SPAM en caso de no encontrar el mail.")
        
        Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    Else
        Call WriteShowMessageBox(Userindex, "El email ya está en uso.")
        
        Call CloseSocket(Userindex)

    End If
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleValidarCuenta(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 7 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim CuentaEmail    As String

    Dim ValidacionCode As String
    
    CuentaEmail = buffer.ReadASCIIString()
    ValidacionCode = buffer.ReadASCIIString()

    If Not CheckMailString(CuentaEmail) Then
        Call WriteShowFrmLogear(Userindex)
        Call WriteShowMessageBox(Userindex, "Email inválido.")
        
        Call CloseSocket(Userindex)
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

                Call WriteShowFrmLogear(Userindex)
                Call WriteShowMessageBox(Userindex, "Cuenta activada con éxito, ya puede ingresar.")
            Else
                Call WriteShowFrmLogear(Userindex)
                Call WriteShowMessageBox(Userindex, "¡Código de activación inválido!")

            End If

        Else
            Call WriteShowFrmLogear(Userindex)
            Call WriteShowMessageBox(Userindex, "La cuenta ya ha sido validada anteriormente.")

        End If

    Else
        Call WriteShowFrmLogear(Userindex)
        Call WriteShowMessageBox(Userindex, "La cuenta no existe.")

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleReValidarCuenta(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
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
        Call WriteShowFrmLogear(Userindex)
        Call WriteShowMessageBox(Userindex, "Nombre invalido.")
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If
    
    'If Useremail <> ObtenerEmail(UserCuenta) Then
    Call WriteShowFrmLogear(Userindex)
    Call WriteShowMessageBox(Userindex, "El email introducido no coincide con el email registrador.")
    
    Call CloseSocket(Userindex)
    Exit Sub
    'End If
    
    If CuentaExiste(UserCuenta) Then
        If ObtenerValidacion(UserCuenta) = 0 Then
            'Call EnviarCorreo(UserCuenta, ObtenerEmail(UserCuenta))
            Call WriteShowFrmLogear(Userindex)
            Call WriteShowMessageBox(Userindex, "Se ha enviado el mail de validación a la dirección designada cuando se creo la cuenta.")
                
        Else
            Call WriteShowFrmLogear(Userindex)
            Call WriteShowMessageBox(Userindex, "La cuenta ya ha sido validada anteriormente.")

        End If

    Else
        Call WriteShowFrmLogear(Userindex)
        Call WriteShowMessageBox(Userindex, "La cuenta no existe.")

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleIngresarConCuenta(ByVal Userindex As Integer)

    Dim Version As String

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 14 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
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
        Call WriteShowMessageBox(Userindex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
        
        Call CloseSocket(Userindex)
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
                    Call WriteShowMessageBox(Userindex, "El servidor se encuentra habilitado solo para administradores. ¡Te esperamos pronto!")
                    Call FlushBuffer(Userindex)
                    Call CloseSocket(Userindex)
                    Exit Sub
            End Select
    End If
    
    
    
    If EntrarCuenta(Userindex, CuentaEmail, CuentaPassword, MacAddress, HDserial) Then
        Call WritePersonajesDeCuenta(Userindex)
        Call WriteMostrarCuenta(Userindex)
    Else
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBorrarPJ(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 15 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
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
        Call WriteShowMessageBox(Userindex, "Esta versión del juego es obsoleta, la versión correcta es la " & ULTIMAVERSION & ". Ejecute el launcher por favor.")
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If

    MacAddress = buffer.ReadASCIIString()
    HDserial = buffer.ReadLong()
    
    If Not EntrarCuenta(Userindex, CuentaEmail, CuentaPassword, MacAddress, HDserial) Then
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If
    
    If Not AsciiValidos(UserDelete) Then
        Call WriteShowMessageBox(Userindex, "Nombre inválido.")
        
        Call CloseSocket(Userindex)
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
    
    Call WritePersonajesDeCuenta(Userindex)
  
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBorrandoCuenta(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 7 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
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
            Call WriteShowFrmLogear(Userindex)
            Call WriteShowMessageBox(Userindex, "Cuenta invalida.")
            
            Call CloseSocket(Userindex)
            Exit Sub

        End If
        
        If UserMail <> ObtenerEmail(AccountDelete) Then
            Call WriteShowFrmLogear(Userindex)
            Call WriteShowMessageBox(Userindex, "El email introducido no coincide con el email registrador.")
            
            Call CloseSocket(Userindex)
            Exit Sub

        End If
        
        If True Then ' Desactivado
            Call WriteShowFrmLogear(Userindex)
            Call WriteShowMessageBox(Userindex, "La contraseña introducida no es correcta.")
            
            Call CloseSocket(Userindex)
            Exit Sub

        End If

        Call BorrarCuenta(AccountDelete)
        Call WriteShowFrmLogear(Userindex)
        Call WriteShowMessageBox(Userindex, "La cuenta ha sido borrada.")
        
    Else
        Call WriteShowFrmLogear(Userindex)
        Call WriteShowMessageBox(Userindex, "La cuenta ingresada no existe.")
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleRecuperandoContraseña(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo ErrHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(UserList(Userindex).incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte

    Dim AcountDelete As String

    Dim UserMail     As String
    
    AcountDelete = buffer.ReadASCIIString()
    UserMail = buffer.ReadASCIIString()
    
    If FileExist(CuentasPath & UCase$(AcountDelete) & ".act", vbNormal) Then
    
        If Not AsciiValidos(AcountDelete) Then
            Call WriteShowFrmLogear(Userindex)
            Call WriteShowMessageBox(Userindex, "Cuenta invalida.")
            
            
            Call CloseSocket(Userindex)
            Exit Sub

        End If

        If UserMail <> ObtenerEmail(AcountDelete) Then
            Call WriteShowFrmLogear(Userindex)
            Call WriteShowMessageBox(Userindex, "El email introducido no coincide con el email registrador.")
            
            Call CloseSocket(Userindex)
            Exit Sub

        End If
        
        If EnviarCorreoRecuperacion(AcountDelete, ObtenerEmail(AcountDelete)) Then
            Call WriteShowFrmLogear(Userindex)
            Call WriteShowMessageBox(Userindex, "La contraseña de la cuenta a sido enviada por email a la direccion registrada.")
        Else
            Call WriteShowFrmLogear(Userindex)
            Call WriteShowMessageBox(Userindex, "Se ha provocado un error al recuperar la clave, reintente mas tarde.")

        End If

    Else
        Call WriteShowFrmLogear(Userindex)
        Call WriteShowMessageBox(Userindex, "La cuenta ingresada no existe.")
        
        Call CloseSocket(Userindex)
        Exit Sub

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(Userindex).incomingData.CopyBuffer(buffer)
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub WritePersonajesDeCuenta(ByVal Userindex As Integer)
    'Author: Pablo Mercavides
    
    Dim UserCuenta                     As String

    Dim CantPersonajes                 As Byte

    Dim Personaje(1 To MAX_PERSONAJES) As PersonajeCuenta

    Dim donador                        As Boolean

    Dim i                              As Byte
    
    UserCuenta = UserList(Userindex).Cuenta
    
    donador = DonadorCheck(UserCuenta)

    If Database_Enabled Then
        CantPersonajes = GetPersonajesCuentaDatabase(UserList(Userindex).AccountID, Personaje)
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
    
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
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

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleCuentaRegresiva(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandlePossUser(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String
        
        UserName = buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.user Then
        
            If Database_Enabled Then
                If Not SetPositionDatabase(UserName, UserList(Userindex).Pos.Map, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y) Then
                    Call WriteConsoleMsg(Userindex, "El usuario " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call WriteVar(CharPath & UCase$(UserName) & ".chr", "INIT", "Position", UserList(Userindex).Pos.Map & "-" & UserList(Userindex).Pos.X & "-" & UserList(Userindex).Pos.Y)
            End If

            Call WriteConsoleMsg(Userindex, "Servidor> Acción realizada con exito! La nueva posicion de " & UserName & "es: " & UserList(Userindex).Pos.Map & "-" & UserList(Userindex).Pos.X & "-" & UserList(Userindex).Pos.Y & "...", FontTypeNames.FONTTYPE_INFO)

            ' Call SendData(UserIndex, UserIndex, PrepareMessageConsoleMsg("Acciín realizada con exito! La nueva posicion de " & UserName & "es: " & UserList(UserIndex).Pos.Map & "-" & UserList(UserIndex).Pos.X & "-" & UserList(UserIndex).Pos.y & "...", FontTypeNames.FONTTYPE_SERVER))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleDuelo(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                    Call WriteConsoleMsg(Userindex, "Duelos> Primero haz click sobre el personaje.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                
                Case Is < 0
                    Call WriteConsoleMsg(Userindex, "Duelos> ¡El persona se encuentra offline!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                
            End Select

            If MapaOcupado Then
                Call WriteConsoleMsg(Userindex, "Duelos> El mapa de duelos esta ocupado, intentalo mas tarde.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            .flags.RetoA = UserList(UserRetado).name
            UserList(UserRetado).flags.SolicitudPendienteDe = .name
        
            Call WriteConsoleMsg(UserRetado, "Duelos> Has sido retado a duelo por " & .name & " si quieres aceptar el duelo escribe /DUELO.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(Userindex, "Duelos> La solicitud a sido enviada al usuario, ahora debes esperar la respuesta de " & UserList(UserRetado).name & ".", FontTypeNames.FONTTYPE_INFO)
               
        Else

           Exit Sub

        End If

        Call SendData(Userindex, 0, PrepareMessageConsoleMsg("Duelo comenzado!", FontTypeNames.FONTTYPE_SERVER))

    End With
    
    Exit Sub
    
ErrHandler:

    Dim Error As Long: Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub WriteGoliathInit(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.Goliath)
        Call .WriteLong(UserList(Userindex).Stats.Banco)
        Call .WriteByte(UserList(Userindex).BancoInvent.NroItems)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteShowFrmLogear(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowFrmLogear)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteShowFrmMapa(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowFrmMapa)
        
        If UserList(Userindex).donador.activo = 1 Then
            Call .WriteInteger(ExpMult * UserList(Userindex).flags.ScrollExp * 1.1)
        Else
            Call .WriteInteger(ExpMult * UserList(Userindex).flags.ScrollExp)

        End If

        Call .WriteInteger(OroMult * UserList(Userindex).flags.ScrollOro)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleNieveToggle(ByVal Userindex As Integer)
        
        On Error GoTo HandleNieveToggle_Err
        

        'Author: Pablo Mercavides
100     With UserList(Userindex)
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

Private Sub HandleNieblaToggle(ByVal Userindex As Integer)
        
        On Error GoTo HandleNieblaToggle_Err
        

        'Author: Pablo Mercavides
100     With UserList(Userindex)
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

Private Sub HandleTransFerGold(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 8 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)
        
        tUser = NameIndex(UserName)

        If tUser <= 0 Then

            If Database_Enabled Then
                If Not AddOroBancoDatabase(UserName, Cantidad) Then
                    Call WriteChatOverHead(Userindex, "El usuario no existe.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
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
            UserList(Userindex).Stats.Banco = UserList(Userindex).Stats.Banco - val(Cantidad) 'Quitamos el oro al usuario
        Else
            UserList(Userindex).Stats.Banco = UserList(Userindex).Stats.Banco - val(Cantidad) 'Quitamos el oro al usuario
            UserList(tUser).Stats.Banco = UserList(tUser).Stats.Banco + val(Cantidad) 'Se lo damos al otro.
        End If

        Call WriteChatOverHead(Userindex, "¡El envio se ha realizado con exito! Gracias por utilizar los servicios de Finanzas Goliath", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave("173", UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
        
        Call WriteUpdateGold(Userindex)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleMoveItem(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
            Call WriteConsoleMsg(Userindex, "Slot bloqueado.", FontTypeNames.FONTTYPE_INFOIAO)
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
            
            Call UpdateUserInv(False, Userindex, SlotViejo)
            Call UpdateUserInv(False, Userindex, SlotNuevo)

        End If
    
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBovedaMoveItem(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
        
        Objeto.ObjIndex = UserList(Userindex).BancoInvent.Object(SlotViejo).ObjIndex
        Objeto.Amount = UserList(Userindex).BancoInvent.Object(SlotViejo).Amount
        
        UserList(Userindex).BancoInvent.Object(SlotViejo).ObjIndex = UserList(Userindex).BancoInvent.Object(SlotNuevo).ObjIndex
        UserList(Userindex).BancoInvent.Object(SlotViejo).Amount = UserList(Userindex).BancoInvent.Object(SlotNuevo).Amount
         
        UserList(Userindex).BancoInvent.Object(SlotNuevo).ObjIndex = Objeto.ObjIndex
        UserList(Userindex).BancoInvent.Object(SlotNuevo).Amount = Objeto.Amount
    
        'Actualizamos el banco
        Call UpdateBanUserInv(False, Userindex, SlotViejo)
        Call UpdateBanUserInv(False, Userindex, SlotNuevo)
        

    End With
    
    Exit Sub
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleQuieroFundarClan(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim refError As String
        
        If UserList(Userindex).GuildIndex > 0 Then
            refError = "Ya perteneces a un clan, no podés fundar otro."
        Else

            If UserList(Userindex).Stats.ELV < 45 Or UserList(Userindex).Stats.UserSkills(eSkill.Liderazgo) < 80 Then
                refError = "Para fundar un clan debes ser nivel 45, tener 80 en liderazgo y tener en tu inventario las 4 gemas: Gema Azul(1), Gema Naranja(1), Gema Gris(1), Gema Roja(1)."
            Else

                If Not TieneObjetos(407, 1, Userindex) Then
                    refError = "Para fundar un clan debes ser nivel 45, tener 80 en liderazgo y tener en tu inventario las 4 gemas: Gema Azul(1), Gema Naranja(1), Gema Gris(1), Gema Roja(1)."
                Else

                    If Not TieneObjetos(408, 1, Userindex) Then
                        refError = "Para fundar un clan debes ser nivel 45, tener 80 en liderazgo y tener en tu inventario las 4 gemas: Gema Azul(1), Gema Naranja(1), Gema Gris(1), Gema Roja(1)."
                    Else

                        If Not TieneObjetos(409, 1, Userindex) Then
                            refError = "Para fundar un clan debes ser nivel 45, tener 80 en liderazgo y tener en tu inventario las 4 gemas: Gema Azul(1), Gema Naranja(1), Gema Gris(1), Gema Roja(1)."
                        Else

                            If Not TieneObjetos(411, 1, Userindex) Then
                                refError = "Para fundar un clan debes ser nivel 45, tener 80 en liderazgo y tener en tu inventario las 4 gemas: Gema Azul(1), Gema Naranja(1), Gema Gris(1), Gema Roja(1)."
                            Else

                                If UserList(Userindex).flags.BattleModo = 1 Then
                                    refError = "No podés fundar un clan ací."
                                Else
                                    refError = "Servidor> íComenzamos a fundar el clan! Ingresa todos los datos solicitados."
                                    Call WriteShowFundarClanForm(Userindex)
                                    
                                End If

                            End If

                        End If

                    End If

                End If

            End If

        End If
                    
        Call WriteConsoleMsg(Userindex, refError, FontTypeNames.FONTTYPE_INFOIAO)
    
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleLlamadadeClan(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Clan> [" & .name & "] solicita apoyo de su clan en " & DarNameMapa(.Pos.Map) & " (" & .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y & "). Puedes ver su ubicaciín en el mapa del mundo.", FontTypeNames.FONTTYPE_GUILD))
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave("43", NO_3D_SOUND, NO_3D_SOUND))
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageUbicacionLlamada(.Pos.Map, .Pos.X, .Pos.Y))
            Else
                Call WriteConsoleMsg(Userindex, "Servidor> El nivel de tu clan debe ser 2 para utilizar esta opciín.", FontTypeNames.FONTTYPE_INFOIAO)

            End If

        Else
            Call WriteConsoleMsg(Userindex, "Servidor> No Perteneces a ningun clan.", FontTypeNames.FONTTYPE_INFOIAO)

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

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
        
106         PrepareMessageNieblandoToggle = .ReadASCIIStringFixed(.Length)

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
        
104         PrepareMessageNevarToggle = .ReadASCIIStringFixed(.Length)

        End With

        
        Exit Function

PrepareMessageNevarToggle_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageNevarToggle", Erl)
        Resume Next
        
End Function

Private Sub HandleGenio(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(Userindex)
        
        'Remove packet ID
        Call .incomingData.ReadInteger
        
        'Si no es GM, no pasara nada :P
        If (.flags.Privilegios And PlayerType.user) <> 0 Then Exit Sub
        
        Dim i As Byte
        For i = 1 To NUMSKILLS
            .Stats.UserSkills(i) = 100
        Next i
        
        Call WriteConsoleMsg(Userindex, "Tus skills fueron editados.", FontTypeNames.FONTTYPE_INFOIAO)

    End With

End Sub

Private Sub HandleCasamiento(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "Primero haz click sobre un sacerdote.", FontTypeNames.FONTTYPE_INFO)
            Else

                If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
                    Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "El sacerdote no puede casarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                Else
            
                    If tUser = Userindex Then
                        Call WriteConsoleMsg(Userindex, "No podés casarte contigo mismo.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If tUser <= 0 Then
                            Call WriteConsoleMsg(Userindex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
                        Else

                            If UserList(tUser).flags.Candidato = Userindex Then
                                UserList(tUser).flags.Casado = 1
                                UserList(tUser).flags.Pareja = UserList(Userindex).name
                                UserList(Userindex).flags.Casado = 1
                                UserList(Userindex).flags.Pareja = UserList(tUser).name
                                Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(FXSound.Casamiento_sound, NO_3D_SOUND, NO_3D_SOUND))
                                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El sacerdote de " & DarNameMapa(.Pos.Map) & " celebra el casamiento entre " & UserList(Userindex).name & " y " & UserList(tUser).name & ".", FontTypeNames.FONTTYPE_WARNING))
                                Call WriteChatOverHead(Userindex, "Los declaro unidos en legal matrimonio íFelicidades!", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                                Call WriteChatOverHead(tUser, "Los declaro unidos en legal matrimonio íFelicidades!", Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, vbWhite)
                                
                            Else
                                Call WriteChatOverHead(Userindex, "La solicitud de casamiento a sido enviada a " & UserName & ".", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
                                Call WriteConsoleMsg(tUser, .name & " desea casarse contigo, para permitirlo haz click en el sacerdote y escribe /PROPONER " & .name & ".", FontTypeNames.FONTTYPE_TALK)
                                UserList(Userindex).flags.Candidato = tUser

                            End If

                        End If

                    End If

                End If

            End If

        Else
            Call WriteConsoleMsg(Userindex, "Primero haz click sobre el sacerdote.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleEnviarCodigo(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim Codigo As String

        Codigo = buffer.ReadASCIIString()

        Call CheckearCodigo(Userindex, Codigo)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCrearTorneo(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 26 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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

        Dim X           As Byte

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
        X = buffer.ReadByte
        Y = buffer.ReadByte
        nombre = buffer.ReadASCIIString
        reglas = buffer.ReadASCIIString
  
        If EsGM(Userindex) Then
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
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleComenzarTorneo(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        'NivelMinimo = buffer.ReadByte
  
        If EsGM(Userindex) Then

            Call ComenzarTorneoOk

        End If
            
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCancelarTorneo(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
  
        If EsGM(Userindex) Then
            Call ResetearTorneo

        End If
            
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBusquedaTesoro(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim Tipo As Byte

        Tipo = buffer.ReadByte()
        
        Dim Mapa As Byte
  
        If EsGM(Userindex) Then
    
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
                        Pos.X = 50
                        Call SpawnNpc(RandomNumber(592, 593), Pos, True, False, True)
                
                End Select
            
            Else
            
                If BusquedaTesoroActiva = True Then
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Todavia nadie fue capaz de encontar el tesoro, recorda que se encuentra en " & DarNameMapa(TesoroNumMapa) & "(" & TesoroNumMapa & "). íQuien sera el valiente que lo encuentre?", FontTypeNames.FONTTYPE_TALK))
                    Call WriteConsoleMsg(Userindex, "Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: " & TesoroNumMapa & "-" & TesoroX & "-" & TesoroY, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Ningun valiente fue capaz de encontrar el item misterioso, recorda que se encuentra en " & DarNameMapa(RegaloNumMapa) & "(" & RegaloNumMapa & "). íTení cuidado!", FontTypeNames.FONTTYPE_TALK))
                    Call WriteConsoleMsg(Userindex, "Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: " & RegaloNumMapa & "-" & RegaloX & "-" & RegaloY, FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
            
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleDropItem(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim Item         As Byte

        Dim X            As Byte

        Dim Y            As Byte

        Dim Depositado   As Byte

        Dim DropCantidad As Integer

        Item = buffer.ReadByte()
        X = buffer.ReadByte()
        Y = buffer.ReadByte()
        DropCantidad = buffer.ReadInteger()
        Depositado = 0

        If UserList(Userindex).flags.Muerto = 1 Then
            ' Call WriteConsoleMsg(UserIndex, "Estas muerto!.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
        Else
    
            If (MapData(UserList(Userindex).Pos.Map, X, Y).Blocked And eBlock.ALL_SIDES) = eBlock.ALL_SIDES Or MapData(UserList(Userindex).Pos.Map, X, Y).TileExit.Map > 0 Or MapData(UserList(Userindex).Pos.Map, X, Y).NpcIndex > 0 Or (MapData(UserList(Userindex).Pos.Map, X, Y).Blocked And FLAG_AGUA) <> 0 Then
            
                'Call WriteConsoleMsg(UserIndex, "Area invalida para tirar el item.", FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(Userindex, "262", FontTypeNames.FONTTYPE_INFO)
            Else
            
                If UserList(Userindex).flags.BattleModo = 1 Then
                    Call WriteConsoleMsg(Userindex, "No podes tirar items en este mapa.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If ObjData(.Invent.Object(Item).ObjIndex).Destruye = 1 Then
                        Call WriteConsoleMsg(Userindex, "Acciín no disponible.", FontTypeNames.FONTTYPE_INFO)
                    Else
                
                        If ObjData(.Invent.Object(Item).ObjIndex).Instransferible = 1 Then
                            Call WriteConsoleMsg(Userindex, "Acciín no disponible.", FontTypeNames.FONTTYPE_INFO)
                        Else
            
                            If ObjData(.Invent.Object(Item).ObjIndex).Newbie = 1 Then
                                Call WriteConsoleMsg(Userindex, "No se pueden tirar los objetos Newbies.", FontTypeNames.FONTTYPE_INFO)
                            Else

                                If ObjData(.Invent.Object(Item).ObjIndex).Intirable = 1 Then
                                    Call WriteConsoleMsg(Userindex, "Este objeto es imposible de tirar.", FontTypeNames.FONTTYPE_INFO)
                                Else
                    
                                    If ObjData(.Invent.Object(Item).ObjIndex).OBJType = eOBJType.otBarcos And UserList(Userindex).flags.Navegando Then
                                        Call WriteConsoleMsg(Userindex, "Para tirar la barca deberias estar en tierra firme.", FontTypeNames.FONTTYPE_INFO)
        
                                    Else
                                            
                                        If ObjData(.Invent.Object(Item).ObjIndex).OBJType = eOBJType.otMonturas And UserList(Userindex).flags.Montado Then
                                            Call WriteConsoleMsg(Userindex, "Para tirar tu montura deberias descender de ella.", FontTypeNames.FONTTYPE_INFO)
        
                                        Else
                
                                            Call DropObj(Userindex, Item, DropCantidad, UserList(Userindex).Pos.Map, X, Y)

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
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleFlagTrabajar(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        UserList(Userindex).Counters.Trabajando = 0
        UserList(Userindex).flags.UsandoMacro = False
        UserList(Userindex).flags.TargetObj = 0 ' Sacamos el targer del objeto
        UserList(Userindex).flags.UltimoMensaje = 0
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleEscribiendo(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        If .flags.Escribiendo = False Then
            .flags.Escribiendo = True
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetEscribiendo(.Char.CharIndex, True))
        Else
            .flags.Escribiendo = False
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetEscribiendo(.Char.CharIndex, False))

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleRequestFamiliar(ByVal Userindex As Integer)
        'Author: Pablo Mercavides
        'Remove packet ID
        
        On Error GoTo HandleRequestFamiliar_Err
        
100     Call UserList(Userindex).incomingData.ReadInteger

102     Call WriteFamiliar(Userindex)

        
        Exit Sub

HandleRequestFamiliar_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleRequestFamiliar", Erl)
        Resume Next
        
End Sub

Public Sub WriteFamiliar(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.Familiar)
        Call .WriteByte(UserList(Userindex).Familiar.Existe)
        Call .WriteByte(UserList(Userindex).Familiar.Muerto)
        Call .WriteASCIIString(UserList(Userindex).Familiar.nombre)
        Call .WriteLong(UserList(Userindex).Familiar.Exp)
        Call .WriteLong(UserList(Userindex).Familiar.ELU)
        Call .WriteByte(UserList(Userindex).Familiar.nivel)
        Call .WriteInteger(UserList(Userindex).Familiar.MinHp)
        Call .WriteInteger(UserList(Userindex).Familiar.MaxHp)
        Call .WriteInteger(UserList(Userindex).Familiar.MinHIT)
        Call .WriteInteger(UserList(Userindex).Familiar.MaxHit)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
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
        
110         PrepareMessageBarFx = .ReadASCIIStringFixed(.Length)

        End With

        
        Exit Function

PrepareMessageBarFx_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageBarFx", Erl)
        Resume Next
        
End Function

Private Sub HandleCompletarAccion(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim Accion As Byte

        Accion = buffer.ReadByte()
        
        If .Accion.AccionPendiente = True Then
            If .Accion.TipoAccion = Accion Then
                Call CompletarAccionFin(Userindex)
            Else
                Call WriteConsoleMsg(Userindex, "Servidor> La acciín que solicitas no se corresponde.", FontTypeNames.FONTTYPE_SERVER)

            End If

        Else
            Call WriteConsoleMsg(Userindex, "Servidor> Tu no tenias ninguna acciín pendiente. ", FontTypeNames.FONTTYPE_SERVER)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleReclamarRecompensa(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Dim Index  As Byte

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Index = buffer.ReadByte()
        
        Call EntregarRecompensas(Userindex, Index)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleTraerRecompensas(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Call EnviarRecompensaStat(Userindex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub WriteRecompensas(ByVal Userindex As Integer)
        
        On Error GoTo WriteRecompensas_Err
        

        '***************************************************
        'Envia las recompensas al cliente!
        'Por Ladder
        '22/04/2015
        'Flor te amo!
        '***************************************************

100     With UserList(Userindex).outgoingData
    
            Dim a, b, c As Byte
 
102         b = UserList(Userindex).UserLogros + 1
104         a = UserList(Userindex).NPcLogros + 1
106         c = UserList(Userindex).LevelLogros + 1
        
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
        
134         Call .WriteInteger(UserList(Userindex).Stats.NPCsMuertos)
        
136         If UserList(Userindex).Stats.NPCsMuertos >= NPcLogros(a).cant Then
138             Call .WriteBoolean(True)
            Else
140             Call .WriteBoolean(False)

            End If
        
            'Logros User
142         Call .WriteASCIIString(UserLogros(b).nombre)
144         Call .WriteASCIIString(UserLogros(b).Desc)
146         Call .WriteInteger(UserLogros(b).cant)
148         Call .WriteInteger(UserLogros(b).TipoRecompensa)
150         Call .WriteInteger(UserList(Userindex).Stats.UsuariosMatados)

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

168         If UserList(Userindex).Stats.UsuariosMatados >= UserLogros(b).cant Then
170             Call .WriteBoolean(True)
            Else
172             Call .WriteBoolean(False)

            End If

            'Nivel User
174         Call .WriteASCIIString(LevelLogros(c).nombre)
176         Call .WriteASCIIString(LevelLogros(c).Desc)
178         Call .WriteInteger(LevelLogros(c).cant)
180         Call .WriteInteger(LevelLogros(c).TipoRecompensa)
182         Call .WriteByte(UserList(Userindex).Stats.ELV)

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

200         If UserList(Userindex).Stats.ELV >= LevelLogros(c).cant Then
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

Private Sub HandleCorreo(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Call WriteListaCorreo(Userindex, False)
        '    Call EnviarRecompensaStat(UserIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleSendCorreo(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 7 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
            AddCorreo Userindex, Nick, msg, 0, FinalCount

        End If
        
        Dim ObjArray As String
        
        If UserList(Userindex).flags.BattleModo = 0 Then

            For i = 1 To ItemCount
                ObjIndex = UserList(Userindex).Invent.Object(Itemlista(i).ObjIndex).ObjIndex
                
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

                                If ObjData(ObjIndex).OBJType = eOBJType.otMonturas And UserList(Userindex).flags.Montado Then
                                    HuboError = True
                                    '  Call WriteConsoleMsg(UserIndex, "Para transferir tu montura deberias descender de ella.", FontTypeNames.FONTTYPE_INFO)
                                Else
                                
                                    Call QuitarUserInvItem(Userindex, Itemlista(i).ObjIndex, Itemlista(i).Amount)
                                    Call UpdateUserInv(False, Userindex, Itemlista(i).ObjIndex)
                                    FinalCount = FinalCount + 1
                                    ObjArray = ObjArray & ObjIndex & "-" & Itemlista(i).Amount & "@"

                                End If

                            End If

                        End If

                    End If

                End If

            Next i
                
            IndexReceptor = NameIndex(Nick)
            AddCorreo Userindex, Nick, msg, ObjArray, FinalCount
    
            If HuboError Then
                Call WriteConsoleMsg(Userindex, "Hubo objetos que no se pudieron enviar.", FontTypeNames.FONTTYPE_INFO)

            End If
            
        Else
            Call WriteConsoleMsg(Userindex, "No podes usar el correo desde el battle.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
    Exit Sub
ErrHandler:
    LogError "Error HandleSendCorreo"

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleRetirarItemCorreo(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Dim MsgIndex As Integer

        MsgIndex = buffer.ReadInteger()
        
        Call ExtractItemCorreo(Userindex, MsgIndex)
        
        Call .incomingData.CopyBuffer(buffer)

    End With
    
    Exit Sub
    
ErrHandler:
    LogError "Error handleRetirarItem"

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBorrarCorreo(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Dim MsgIndex As Integer

        MsgIndex = buffer.ReadInteger()
        
        Call BorrarCorreoMail(Userindex, MsgIndex)
        
        Call .incomingData.CopyBuffer(buffer)

    End With
    
    Exit Sub
    
ErrHandler:

    LogError "Error BorrarCorreo"

    Dim Error As Long

    Error = Err.Number
    
    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleInvitarGrupo(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If


    With UserList(Userindex)
        
        'Remove packet ID
        Call .incomingData.ReadInteger
        
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
            
        Else
            
            If .Grupo.CantidadMiembros <= UBound(.Grupo.Miembros) Then
                Call WriteWorkRequestTarget(Userindex, eSkill.Grupo)
            Else
                Call WriteConsoleMsg(Userindex, "¡No podés invitar a más personas!", FontTypeNames.FONTTYPE_INFO)
            End If

        End If

    End With


End Sub

Private Sub HandleMarcaDeClan(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!.", FontTypeNames.FONTTYPE_INFO)
            Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
       
        Call WriteWorkRequestTarget(Userindex, eSkill.MarcaDeClan)
                
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleMarcaDeGM(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    With UserList(Userindex)
    
        If .incomingData.Length < 2 Then
            Err.raise .incomingData.NotEnoughDataErrCode
            Exit Sub
        End If

        Call .incomingData.ReadInteger
          
        Call WriteWorkRequestTarget(Userindex, eSkill.MarcaDeGM)

    End With

End Sub

Public Sub WritePreguntaBox(ByVal Userindex As Integer, ByVal message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.ShowPregunta)
        Call .WriteASCIIString(message)

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleResponderPregunta(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
        
            Select Case UserList(Userindex).flags.pregunta

                Case 1
                    Log = "Repuesta Afirmativa 1"

                    'Call WriteConsoleMsg(UserIndex, "El usuario desea unirse al grupo.", FontTypeNames.FONTTYPE_SUBASTA)
                    ' UserList(UserIndex).Grupo.PropuestaDe = 0
                    If UserList(Userindex).Grupo.PropuestaDe <> 0 Then
                
                        If UserList(UserList(Userindex).Grupo.PropuestaDe).Grupo.Lider <> UserList(Userindex).Grupo.PropuestaDe Then
                            Call WriteConsoleMsg(Userindex, "íEl lider del grupo a cambiado, imposible unirse!", FontTypeNames.FONTTYPE_INFOIAO)
                        Else
                        
                            Log = "Repuesta Afirmativa 1-1 "
                        
                            If UserList(UserList(Userindex).Grupo.PropuestaDe).Grupo.Lider = 0 Then
                                Call WriteConsoleMsg(Userindex, "íEl grupo ya no existe!", FontTypeNames.FONTTYPE_INFOIAO)
                            Else
                            
                                Log = "Repuesta Afirmativa 1-2 "
                            
                                If UserList(UserList(Userindex).Grupo.PropuestaDe).Grupo.CantidadMiembros = 1 Then
                                    Call WriteLocaleMsg(UserList(Userindex).Grupo.PropuestaDe, "36", FontTypeNames.FONTTYPE_INFOIAO)
                                    'Call WriteConsoleMsg(UserList(UserIndex).Grupo.PropuestaDe, "íEl grupo a sido creado!", FontTypeNames.FONTTYPE_INFOIAO)
                                    UserList(UserList(Userindex).Grupo.PropuestaDe).Grupo.EnGrupo = True
                                    Log = "Repuesta Afirmativa 1-3 "

                                End If
                                
                                Log = "Repuesta Afirmativa 1-4"
                                UserList(UserList(Userindex).Grupo.PropuestaDe).Grupo.CantidadMiembros = UserList(UserList(Userindex).Grupo.PropuestaDe).Grupo.CantidadMiembros + 1
                                UserList(UserList(Userindex).Grupo.PropuestaDe).Grupo.Miembros(UserList(UserList(Userindex).Grupo.PropuestaDe).Grupo.CantidadMiembros) = Userindex
                                UserList(Userindex).Grupo.EnGrupo = True
                                
                                Dim Index As Byte
                                
                                Log = "Repuesta Afirmativa 1-5 "
                                
                                For Index = 2 To UserList(UserList(Userindex).Grupo.PropuestaDe).Grupo.CantidadMiembros - 1
                                    Call WriteLocaleMsg(UserList(UserList(Userindex).Grupo.PropuestaDe).Grupo.Miembros(Index), "40", FontTypeNames.FONTTYPE_INFOIAO, UserList(Userindex).name)
                                
                                Next Index
                                
                                Log = "Repuesta Afirmativa 1-6 "
                                'Call WriteConsoleMsg(UserList(UserIndex).Grupo.PropuestaDe, "í" & UserList(UserIndex).name & " a sido aíadido al grupo!", FontTypeNames.FONTTYPE_INFOIAO)
                                Call WriteLocaleMsg(UserList(Userindex).Grupo.PropuestaDe, "40", FontTypeNames.FONTTYPE_INFOIAO, UserList(Userindex).name)
                                
                                Call WriteConsoleMsg(Userindex, "¡Has sido aíadido al grupo!", FontTypeNames.FONTTYPE_INFOIAO)
                                
                                Log = "Repuesta Afirmativa 1-7 "
                                
                                Call RefreshCharStatus(UserList(Userindex).Grupo.PropuestaDe)
                                Call RefreshCharStatus(Userindex)
                                 
                                Log = "Repuesta Afirmativa 1-8"

                            End If

                        End If

                    Else
                    
                        Call WriteConsoleMsg(Userindex, "Servidor> Solicitud de grupo invalida, reintente...", FontTypeNames.FONTTYPE_SERVER)
                    
                    End If

                    'unirlo
                Case 2
                    Log = "Repuesta Afirmativa 2"
                    UserList(Userindex).Faccion.Status = 1
                    Call WriteConsoleMsg(Userindex, "íAhora sos un ciudadano!", FontTypeNames.FONTTYPE_INFOIAO)
                    Call RefreshCharStatus(Userindex)
                    
                Case 3
                    Log = "Repuesta Afirmativa 3"
                    
                    UserList(Userindex).Hogar = UserList(Userindex).PosibleHogar

                    Select Case UserList(Userindex).Hogar

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
                    
                    If UserList(Userindex).flags.TargetNPC <> 0 Then
                    
                        Call WriteChatOverHead(Userindex, "íGracias " & UserList(Userindex).name & "! Ahora perteneces a la ciudad de " & DeDonde & ".", Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, vbWhite)
                    Else
                        Call WriteConsoleMsg(Userindex, "íGracias " & UserList(Userindex).name & "! Ahora perteneces a la ciudad de " & DeDonde & ".", FontTypeNames.FONTTYPE_INFOIAO)

                    End If
                    
                Case 4
                    Log = "Repuesta Afirmativa 4"
                
                    If UserList(Userindex).flags.TargetUser <> 0 Then
                
                        If UserList(UserList(Userindex).flags.TargetUser).flags.BattleModo = 1 Then
                            Call WriteConsoleMsg(Userindex, "No podes usar el sistema de comercio cuando el otro personaje esta en el battle.", FontTypeNames.FONTTYPE_EXP)
                        
                        Else
                    
                            UserList(Userindex).ComUsu.DestUsu = UserList(Userindex).flags.TargetUser
                            UserList(Userindex).ComUsu.DestNick = UserList(UserList(Userindex).flags.TargetUser).name
                            UserList(Userindex).ComUsu.cant = 0
                            UserList(Userindex).ComUsu.Objeto = 0
                            UserList(Userindex).ComUsu.Acepto = False
                    
                            'Rutina para comerciar con otro usuario
                            Call IniciarComercioConUsuario(Userindex, UserList(Userindex).flags.TargetUser)

                        End If

                    Else
                        Call WriteConsoleMsg(Userindex, "Servidor> Solicitud de comercio invalida, reintente...", FontTypeNames.FONTTYPE_SERVER)
                
                    End If
                
                Case 5
                    Log = "Repuesta Afirmativa 5"
                
                    If UCase$(MapInfo(UserList(Userindex).Pos.Map).restrict_mode) = "NEWBIE" Then
                        Call WarpToLegalPos(Userindex, 140, 53, 58)
                    
                        If UserList(Userindex).donador.activo = 0 Then ' Donador no espera tiempo
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Resucitar, 400, False))
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 400, Accion_Barra.Resucitar))
                        Else
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Resucitar, 10, False))
                            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 10, Accion_Barra.Resucitar))

                        End If
                    
                        UserList(Userindex).Accion.AccionPendiente = True
                        UserList(Userindex).Accion.Particula = ParticulasIndex.Resucitar
                        UserList(Userindex).Accion.TipoAccion = Accion_Barra.Resucitar
    
                        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave("104", UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
                        'Call WriteConsoleMsg(UserIndex, "El Cura lanza unas palabras al aire. Comienzas a sentir como tu cuerpo se vuelve a formar...", FontTypeNames.FONTTYPE_INFO)
                        Call WriteLocaleMsg(Userindex, "82", FontTypeNames.FONTTYPE_INFOIAO)
                    Else
                        Call WriteConsoleMsg(Userindex, "Ya no te encuentras en un mapa newbie.", FontTypeNames.FONTTYPE_INFOIAO)

                    End If
                
                Case Else
                    Call WriteConsoleMsg(Userindex, "No tienes preguntas pendientes.", FontTypeNames.FONTTYPE_INFOIAO)
                    
            End Select
        
        Else
            Log = "Repuesta negativa"
        
            Select Case UserList(Userindex).flags.pregunta

                Case 1
                    Log = "Repuesta negativa 1"

                    If UserList(Userindex).Grupo.PropuestaDe <> 0 Then
                        Call WriteConsoleMsg(UserList(Userindex).Grupo.PropuestaDe, "El usuario no esta interesado en formar parte del grupo.", FontTypeNames.FONTTYPE_INFOIAO)

                    End If

                    UserList(Userindex).Grupo.PropuestaDe = 0
                    Call WriteConsoleMsg(Userindex, "Has rechazado la propuesta.", FontTypeNames.FONTTYPE_INFOIAO)
                
                Case 2
                    Log = "Repuesta negativa 2"
                    UserList(Userindex).Faccion.Status = 0
                    Call WriteConsoleMsg(Userindex, "¡Continuas siendo neutral!", FontTypeNames.FONTTYPE_INFOIAO)
                    Call RefreshCharStatus(Userindex)

                Case 3
                    Log = "Repuesta negativa 3"
                    
                    Select Case UserList(Userindex).PosibleHogar

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
                    
                    If UserList(Userindex).flags.TargetNPC <> 0 Then
                        Call WriteChatOverHead(Userindex, "¡No hay problema " & UserList(Userindex).name & "! Sos bienvenido en " & DeDonde & " cuando gustes.", Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, vbWhite)

                    End If

                    UserList(Userindex).PosibleHogar = UserList(Userindex).Hogar
                    
                Case 4
                    Log = "Repuesta negativa 4"
                    
                    If UserList(Userindex).flags.TargetUser <> 0 Then
                        Call WriteConsoleMsg(UserList(Userindex).flags.TargetUser, "El usuario no desea comerciar en este momento.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Case 5
                    Log = "Repuesta negativa 5"
                    'No hago nada. dijo que no lo resucite
                        
                Case Else
                    Call WriteConsoleMsg(Userindex, "No tienes preguntas pendientes.", FontTypeNames.FONTTYPE_INFOIAO)

            End Select
            
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
    Exit Sub
    
ErrHandler:

    LogError "Error ResponderPregunta " & Log

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleRequestGrupo(ByVal Userindex As Integer)

    On Error GoTo hErr

    'Author: Pablo Mercavides
    'Remove packet ID
    Call UserList(Userindex).incomingData.ReadInteger

    Call WriteDatosGrupo(Userindex)
    
    Exit Sub
    
hErr:
    LogError "Error HandleRequestGrupo"

End Sub

Public Sub WriteDatosGrupo(ByVal Userindex As Integer)

    Dim i As Byte

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.DatosGrupo)
        Call .WriteBoolean(UserList(Userindex).Grupo.EnGrupo)
        
        If UserList(Userindex).Grupo.EnGrupo = True Then
            Call .WriteByte(UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros)
            'Call .WriteByte(UserList(UserList(UserIndex).Grupo.Lider).name)
   
            If UserList(Userindex).Grupo.Lider = Userindex Then
             
                For i = 1 To UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros

                    If i = 1 Then
                        Call .WriteASCIIString(UserList(UserList(Userindex).Grupo.Miembros(i)).name & "(Lider)")
                    Else
                        Call .WriteASCIIString(UserList(UserList(Userindex).Grupo.Miembros(i)).name)

                    End If

                Next i

            Else
          
                For i = 1 To UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros
                
                    If i = 1 Then
                        Call .WriteASCIIString(UserList(UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(i)).name & "(Lider)")
                    Else
                        Call .WriteASCIIString(UserList(UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(i)).name)

                    End If

                Next i

            End If

        End If
   
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleAbandonarGrupo(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        If UserList(Userindex).Grupo.Lider = Userindex Then
            
            Call FinalizarGrupo(Userindex)

            Dim i As Byte
            
            For i = 2 To UserList(Userindex).Grupo.CantidadMiembros
                Call WriteUbicacion(Userindex, i, 0)
            Next i

            UserList(Userindex).Grupo.CantidadMiembros = 0
            UserList(Userindex).Grupo.EnGrupo = False
            UserList(Userindex).Grupo.Lider = 0
            UserList(Userindex).Grupo.PropuestaDe = 0
            Call WriteConsoleMsg(Userindex, "Has disuelto el grupo.", FontTypeNames.FONTTYPE_INFOIAO)
            Call RefreshCharStatus(Userindex)
        Else
            Call SalirDeGrupo(Userindex)

        End If
                
        Call .incomingData.CopyBuffer(buffer)

    End With
    
    Exit Sub
ErrHandler:

    LogError "Error HandleAbandonarGrupo"

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub WriteUbicacion(ByVal Userindex As Integer, ByVal Miembro As Byte, ByVal GPS As Integer)

    Dim i   As Byte

    Dim X   As Byte

    Dim Y   As Byte

    Dim Map As Integer

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
    
        Call .WriteByte(ServerPacketID.ubicacion)
        Call .WriteByte(Miembro)

        If GPS > 0 Then
        
            Call .WriteByte(UserList(GPS).Pos.X)
            Call .WriteByte(UserList(GPS).Pos.Y)
            Call .WriteInteger(UserList(GPS).Pos.Map)
        Else
            Call .WriteByte(0)
            Call .WriteByte(0)
            Call .WriteInteger(0)

        End If
   
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleHecharDeGrupo(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim Indice As Byte

        Indice = buffer.ReadByte()
        
        Call HecharMiembro(Userindex, Indice)
                
        Call .incomingData.CopyBuffer(buffer)

    End With
    
    Exit Sub
ErrHandler:
    LogError "Error HandleHecharDeGrupo"

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleMacroPos(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        UserList(Userindex).ChatCombate = buffer.ReadByte()
        UserList(Userindex).ChatGlobal = buffer.ReadByte()
                
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub WriteCorreoPicOn(ByVal Userindex As Integer)

    '***************************************************
    '***************************************************
    On Error GoTo ErrHandler

    Call UserList(Userindex).outgoingData.WriteByte(ServerPacketID.CorreoPicOn)
    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleSubastaInfo(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        If Subasta.HaySubastaActiva Then

            Call WriteConsoleMsg(Userindex, "Subastador: " & Subasta.Subastador, FontTypeNames.FONTTYPE_SUBASTA)
            Call WriteConsoleMsg(Userindex, "Objeto: " & ObjData(Subasta.ObjSubastado).name & " (" & Subasta.ObjSubastadoCantidad & ")", FontTypeNames.FONTTYPE_SUBASTA)

            If Subasta.HuboOferta Then
                Call WriteConsoleMsg(Userindex, "Mejor oferta: " & Subasta.MejorOferta & " monedas de oro por " & Subasta.Comprador & ".", FontTypeNames.FONTTYPE_SUBASTA)
                Call WriteConsoleMsg(Userindex, "Podes realizar una oferta escribiendo /OFERTAR " & Subasta.MejorOferta + 100, FontTypeNames.FONTTYPE_SUBASTA)
            Else
                Call WriteConsoleMsg(Userindex, "Oferta inicial: " & Subasta.OfertaInicial & " monedas de oro.", FontTypeNames.FONTTYPE_SUBASTA)
                Call WriteConsoleMsg(Userindex, "Podes realizar una oferta escribiendo /OFERTAR " & Subasta.OfertaInicial + 100, FontTypeNames.FONTTYPE_SUBASTA)

            End If

            Call WriteConsoleMsg(Userindex, "Tiempo Restante de subasta:  " & SumarTiempo(Subasta.TiempoRestanteSubasta), FontTypeNames.FONTTYPE_SUBASTA)
            
        Else
            Call WriteConsoleMsg(Userindex, "No hay ninguna subasta activa en este momento.", FontTypeNames.FONTTYPE_SUBASTA)

        End If
                
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleScrollInfo(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim activo As Boolean

        Dim HR     As Integer

        Dim MS     As Integer

        Dim SS     As Integer

        Dim secs   As Integer
        
        If UserList(Userindex).flags.ScrollExp > 1 Then
            secs = UserList(Userindex).Counters.ScrollExperiencia
            HR = secs \ 3600
            MS = (secs Mod 3600) \ 60
            SS = (secs Mod 3600) Mod 60

            If SS > 9 Then
                Call WriteConsoleMsg(Userindex, "Scroll de experiencia activo. Tiempo restante: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_INFOIAO)
            Else
                Call WriteConsoleMsg(Userindex, "Scroll de experiencia activo. Tiempo restante: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_INFOIAO)

            End If

            activo = True

        End If

        If UserList(Userindex).flags.ScrollOro > 1 Then
            secs = UserList(Userindex).Counters.ScrollOro
            HR = secs \ 3600
            MS = (secs Mod 3600) \ 60
            SS = (secs Mod 3600) Mod 60

            If SS > 9 Then
                Call WriteConsoleMsg(Userindex, "Scroll de oro activo. Tiempo restante: " & MS & ":" & SS & " minuto(s).", FontTypeNames.FONTTYPE_INFOIAO)
            Else
                Call WriteConsoleMsg(Userindex, "Scroll de oro activo. Tiempo restante: " & MS & ":0" & SS & " minuto(s).", FontTypeNames.FONTTYPE_INFOIAO)

            End If

            activo = True

        End If

        If Not activo Then
            Call WriteConsoleMsg(Userindex, "No tenes ningun scroll activo.", FontTypeNames.FONTTYPE_INFOIAO)

        End If
                
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCancelarExit(ByVal Userindex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleCancelarExit_Err
        

100     With UserList(Userindex)
            'Remove Packet ID
102         Call .incomingData.ReadInteger
    
            'If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

104         Call CancelExit(Userindex)

        End With
        
        
        Exit Sub

HandleCancelarExit_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleCancelarExit", Erl)
        Resume Next
        
End Sub

Private Sub HandleBanCuenta(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 6 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
            Call BanAccount(Userindex, UserName, Reason)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleUnBanCuenta(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String
        
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call UnBanAccount(Userindex, UserName)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBanSerial(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String
         
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanSerialOK(Userindex, UserName)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleUnBanSerial(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim UserName As String
         
        UserName = buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call UnBanSerialOK(Userindex, UserName)
            
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCerrarCliente(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 4 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                Call WriteConsoleMsg(Userindex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
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

ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleEventoInfo(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        If EventoActivo Then
            Call WriteConsoleMsg(Userindex, PublicidadEvento & ". Tiempo restante: " & TiempoRestanteEvento & " minuto(s).", FontTypeNames.FONTTYPE_New_Eventos)
        Else
            Call WriteConsoleMsg(Userindex, "Eventos> Actualmente no hay ningun evento en curso.", FontTypeNames.FONTTYPE_New_Eventos)

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
            Call WriteConsoleMsg(Userindex, "Eventos> El proximo evento " & DescribirEvento(HoraProximo) & " iniciara a las " & HoraProximo & ":00 horas.", FontTypeNames.FONTTYPE_New_Eventos)
        Else
            Call WriteConsoleMsg(Userindex, "Eventos> No hay eventos proximos.", FontTypeNames.FONTTYPE_New_Eventos)

        End If
 
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCrearEvento(ByVal Userindex As Integer)

    '***************************************************
    'Author: Pablo Mercavides
    '***************************************************
    If UserList(Userindex).incomingData.Length < 5 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
                    Call WriteConsoleMsg(Userindex, "Utilice /CREAREVENTO TIPO@DURACION@MULTIPLICACION.", FontTypeNames.FONTTYPE_New_Eventos)
                Else
                
                    Call ForzarEvento(Tipo, duracion, multiplicacion, UserList(Userindex).name)
                  
                End If

            Else
                Call WriteConsoleMsg(Userindex, "Ya hay un evento en curso. Finalicelo con /FINEVENTO primero.", FontTypeNames.FONTTYPE_New_Eventos)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleBanTemporal(ByVal Userindex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(Userindex).incomingData.Length < 7 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
            Call Admin.BanTemporal(UserName, dias, Reason, UserList(Userindex).name)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With

ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleTraerShop(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        If UserList(Userindex).flags.BattleModo = 1 Then
            Call WriteConsoleMsg(Userindex, "No disponible aquí.", FontTypeNames.FONTTYPE_INFOIAO)
        Else
            Call WriteShop(Userindex)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleTraerRanking(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Call WriteRanking(Userindex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandlePareja(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 2 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger
        
        Dim parejaindex As Integer

        If Not UserList(Userindex).flags.BattleModo Then
                
            If UserList(Userindex).donador.activo = 1 Then
                If MapInfo(UserList(Userindex).Pos.Map).Seguro = 1 Then
                    If UserList(Userindex).flags.Casado = 1 Then
                        parejaindex = NameIndex(UserList(Userindex).flags.Pareja)
                        
                        If parejaindex > 0 Then
                            If Not UserList(parejaindex).flags.BattleModo Then
                                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Runa, 600, False))
                                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 600, Accion_Barra.GoToPareja))
                                UserList(Userindex).Accion.AccionPendiente = True
                                UserList(Userindex).Accion.Particula = ParticulasIndex.Runa
                                UserList(Userindex).Accion.TipoAccion = Accion_Barra.GoToPareja
                            Else
                                Call WriteConsoleMsg(Userindex, "Tu pareja esta en modo battle. No podés teletransportarte hacia ella.", FontTypeNames.FONTTYPE_INFOIAO)

                            End If
                                
                        Else
                            Call WriteConsoleMsg(Userindex, "Tu pareja no esta online.", FontTypeNames.FONTTYPE_INFOIAO)

                        End If

                    Else
                        Call WriteConsoleMsg(Userindex, "No estas casado con nadie.", FontTypeNames.FONTTYPE_INFOIAO)

                    End If

                Else
                    Call WriteConsoleMsg(Userindex, "Solo disponible en zona segura.", FontTypeNames.FONTTYPE_INFOIAO)

                End If
                
            Else
                Call WriteConsoleMsg(Userindex, "Opcion disponible unicamente para usuarios donadores.", FontTypeNames.FONTTYPE_INFOIAO)

            End If

        Else
            Call WriteConsoleMsg(Userindex, "No podés usar esta opciín en el battle.", FontTypeNames.FONTTYPE_INFOIAO)
        
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Sub WriteShop(ByVal Userindex As Integer)

    '***************************************************
    On Error GoTo ErrHandler

    Dim i              As Long

    Dim obj            As ObjData

    Dim validIndexes() As Integer

    Dim Count          As Integer
    
    ReDim validIndexes(1 To UBound(ObjDonador()))
    
    With UserList(Userindex).outgoingData
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
        
        Call .WriteLong(CreditosDonadorCheck(UserList(Userindex).Cuenta))
        Call .WriteInteger(DiasDonadorCheck(UserList(Userindex).Cuenta))

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteRanking(ByVal Userindex As Integer)

    '***************************************************
    On Error GoTo ErrHandler

    Dim i As Byte
    
    With UserList(Userindex).outgoingData
        Call .WriteByte(ServerPacketID.Ranking)

        For i = 1 To 10
            Call .WriteASCIIString(Rankings(1).user(i).Nick)
            Call .WriteInteger(Rankings(1).user(i).Value)
        Next i
        
    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Private Sub HandleComprarItem(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 3 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

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
        
        For i = 1 To UserList(Userindex).CurrentInventorySlots

            If UserList(Userindex).Invent.Object(i).ObjIndex = 0 Then InvSlotsLibres = InvSlotsLibres + 1
        Next i
    
        'Nos fijamos si entra
        If InvSlotsLibres = 0 Then
            Call WriteConsoleMsg(Userindex, "Donación> Sin espacio en el inventario.", FontTypeNames.FONTTYPE_WARNING)
        Else

            If CreditosDonadorCheck(UserList(Userindex).Cuenta) - ObjDonador(ItemIndex).Valor >= 0 Then
                ObjComprado.Amount = ObjDonador(ItemIndex).Cantidad
                ObjComprado.ObjIndex = ObjDonador(ItemIndex).ObjIndex
            
                LogeoDonador = LogeoDonador & vbCrLf & "****************************************************" & vbCrLf
                LogeoDonador = LogeoDonador & "Compra iniciada. Balance de la cuenta " & CreditosDonadorCheck(UserList(Userindex).Cuenta) & " creditos." & vbCrLf
                LogeoDonador = LogeoDonador & "El personaje " & UserList(Userindex).name & "(" & UserList(Userindex).Cuenta & ") Compro el item " & ObjData(ObjDonador(ItemIndex).ObjIndex).name & vbCrLf
                LogeoDonador = LogeoDonador & "Se descontaron " & CLng(ObjDonador(ItemIndex).Valor) & " creditos de la cuenta " & UserList(Userindex).Cuenta & "." & vbCrLf
            
                If Not MeterItemEnInventario(Userindex, ObjComprado) Then
                    LogeoDonador = LogeoDonador & "El item se tiro al piso" & vbCrLf
                    Call TirarItemAlPiso(UserList(Userindex).Pos, ObjComprado)

                End If
                
                LogeoDonador = LogeoDonador & "****************************************************" & vbCrLf
             
                Call RestarCreditosDonador(UserList(Userindex).Cuenta, CLng(ObjDonador(ItemIndex).Valor))
                Call WriteConsoleMsg(Userindex, "Donación> Gracias por tu compra. Tu saldo es de " & CreditosDonadorCheck(UserList(Userindex).Cuenta) & " creditos.", FontTypeNames.FONTTYPE_WARNING)
                Call LogearEventoDeDonador(LogeoDonador)
                Call SaveUser(Userindex)
                Call WriteActShop(Userindex)
            Else
                Call WriteConsoleMsg(Userindex, "Donación> Tu saldo es insuficiente. Actualmente tu saldo es de " & CreditosDonadorCheck(UserList(Userindex).Cuenta) & " creditos.", FontTypeNames.FONTTYPE_WARNING)
                Call WriteActShop(Userindex)

            End If

        End If
    
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Private Sub HandleCompletarViaje(ByVal Userindex As Integer)
    'Author: Pablo Mercavides

    If UserList(Userindex).incomingData.Length < 7 Then
        Err.raise UserList(Userindex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo ErrHandler

    With UserList(Userindex)

        Dim buffer As New clsByteQueue

        Call buffer.CopyBuffer(.incomingData)
        'Remove packet ID
        Call buffer.ReadInteger

        Dim Destino As Byte

        Dim costo   As Long

        Destino = buffer.ReadByte()
        costo = buffer.ReadLong()
        
        Dim DeDonde As CityWorldPos

        If UserList(Userindex).Stats.GLD < costo Then
            Call WriteConsoleMsg(Userindex, "No tienes suficiente dinero.", FontTypeNames.FONTTYPE_INFO)
            
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
                If UserList(Userindex).Stats.UserSkills(eSkill.Navegacion) < 80 Then
                    Rem Call WriteConsoleMsg(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(Userindex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_WARNING)
                Else

                    If UserList(Userindex).flags.TargetNPC <> 0 Then
                        If Npclist(UserList(Userindex).flags.TargetNPC).SoundClose <> 0 Then
                            Call WritePlayWave(Userindex, Npclist(UserList(Userindex).flags.TargetNPC).SoundClose, NO_3D_SOUND, NO_3D_SOUND)

                        End If

                    End If

                    Call WarpToLegalPos(Userindex, DeDonde.MapaViaje, DeDonde.ViajeX, DeDonde.ViajeY, True)
                    Call WriteConsoleMsg(Userindex, "Has viajado por varios días, te sientes exhausto!", FontTypeNames.FONTTYPE_WARNING)
                    UserList(Userindex).Stats.MinAGU = 0
                    UserList(Userindex).Stats.MinHam = 0
                    UserList(Userindex).flags.Sed = 1
                    UserList(Userindex).flags.Hambre = 1
                    
                    UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD - costo
                    Call WriteUpdateHungerAndThirst(Userindex)
                    Call WriteUpdateUserStats(Userindex)

                End If

            Else
            
                Dim Map As Integer

                Dim X   As Byte

                Dim Y   As Byte
            
                Map = DeDonde.MapaViaje
                X = DeDonde.ViajeX
                Y = DeDonde.ViajeY

                If UserList(Userindex).flags.TargetNPC <> 0 Then
                    If Npclist(UserList(Userindex).flags.TargetNPC).SoundClose <> 0 Then
                        Call WritePlayWave(Userindex, Npclist(UserList(Userindex).flags.TargetNPC).SoundClose, NO_3D_SOUND, NO_3D_SOUND)

                    End If

                End If
                
                Call WarpUserChar(Userindex, Map, X, Y, True)
                Call WriteConsoleMsg(Userindex, "Has viajado por varios días, te sientes exhausto!", FontTypeNames.FONTTYPE_WARNING)
                UserList(Userindex).Stats.MinAGU = 0
                UserList(Userindex).Stats.MinHam = 0
                UserList(Userindex).flags.Sed = 1
                UserList(Userindex).flags.Hambre = 1
                
                UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD - costo
                Call WriteUpdateHungerAndThirst(Userindex)
                Call WriteUpdateUserStats(Userindex)
        
            End If

        End If
    
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(buffer)

    End With
    
ErrHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.raise Error

End Sub

Public Function PrepareMessageCreateRenderValue(ByVal X As Byte, ByVal Y As Byte, ByVal rValue As Double, ByVal rType As Byte)
        '***************************************************
        'Author: maTih.-
        'Last Modification: 09/06/2012 - ^[GS]^
        '***************************************************
        
        On Error GoTo PrepareMessageCreateRenderValue_Err
        

        ' @ Envia el paquete para crear un valor en el render
     
100     With auxiliarBuffer
102         .WriteByte ServerPacketID.CreateRenderText
104         .WriteByte X
106         .WriteByte Y
108         .WriteDouble rValue
110         .WriteByte rType
         
112         PrepareMessageCreateRenderValue = .ReadASCIIStringFixed(.Length)
         
        End With
     
        
        Exit Function

PrepareMessageCreateRenderValue_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.PrepareMessageCreateRenderValue", Erl)
        Resume Next
        
End Function

Public Sub WriteActShop(ByVal Userindex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
    
        Call .WriteByte(ServerPacketID.ActShop)
        Call .WriteLong(CreditosDonadorCheck(UserList(Userindex).Cuenta))
        
        Call .WriteInteger(DiasDonadorCheck(UserList(Userindex).Cuenta))

    End With

    Exit Sub

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteViajarForm(ByVal Userindex As Integer, ByVal NpcIndex As Integer)

    'Author: Pablo Mercavides
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
    
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

ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub HandleQuest(ByVal Userindex As Integer)
        
        On Error GoTo HandleQuest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete Quest.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim NpcIndex As Integer

        Dim tmpByte  As Byte
 
        'Leemos el paquete
    
100     Call UserList(Userindex).incomingData.ReadInteger
 
102     NpcIndex = UserList(Userindex).flags.TargetNPC
    
104     If NpcIndex = 0 Then Exit Sub
    
        'Esta el personaje en la distancia correcta?
106     If Distancia(UserList(Userindex).Pos, Npclist(NpcIndex).Pos) > 5 Then
108         Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        'El NPC hace quests?
110     If Npclist(NpcIndex).NumQuest = 0 Then
112         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("No tengo ninguna mision para ti.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
            Exit Sub

        End If
    
        'El personaje ya hizo la quest?
114   '  If UserDoneQuest(UserIndex, Npclist(NpcIndex).QuestNumber) Then
116     '    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Ya has hecho una mision para mi.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
         '   Exit Sub

       ' End If
        
        
        
        
        
        
 
        'El personaje tiene suficiente nivel?
118    ' If UserList(UserIndex).Stats.ELV < QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel Then
120       '  Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Debes ser por lo menos nivel " & QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel & " para emprender esta mision.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
         '   Exit Sub

        'End If
    
        'A esta altura ya analizo todas las restricciones y esta preparado para el handle propiamente dicho
 
122    ' tmpByte = TieneQuest(UserIndex, Npclist(NpcIndex).QuestNumber)
    
124   '  If tmpByte Then
            'El usuario esta haciendo la quest, entonces va a hablar con el NPC para recibir la recompensa.
126      '   Call FinishQuest(UserIndex, Npclist(NpcIndex).QuestNumber, tmpByte)
      '  Else
            'El usuario no esta haciendo la quest, entonces primero recibe un informe con los detalles de la mision.
128      '   tmpByte = FreeQuestSlot(UserIndex)
        
            'El personaje tiene algun slot de quest para la nueva quest?
130      '   If tmpByte = 0 Then
132             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
         ''       Exit Sub

       '     End If
        
            'Enviamos los detalles de la quest
134      '   Call WriteQuestDetails(UserIndex, Npclist(NpcIndex).QuestNumber)

       ' End If

        
        Exit Sub

HandleQuest_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleQuest", Erl)
        Resume Next
        
End Sub

Public Sub HandleQuestAccept(ByVal Userindex As Integer)
        
        On Error GoTo HandleQuestAccept_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el evento de aceptar una quest.
        'Last modified: 31/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim NpcIndex  As Integer

        Dim QuestSlot As Byte
        
        Dim Indice As Byte
 
100     Call UserList(Userindex).incomingData.ReadInteger

        Indice = UserList(Userindex).incomingData.ReadByte
 
102     NpcIndex = UserList(Userindex).flags.TargetNPC
    
104     If NpcIndex = 0 Then Exit Sub
        If Indice = 0 Then Exit Sub
    
        'Esta el personaje en la distancia correcta?
106     If Distancia(UserList(Userindex).Pos, Npclist(NpcIndex).Pos) > 5 Then
108         Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If TieneQuest(Userindex, Npclist(NpcIndex).QuestNumber(Indice)) Then
            Call WriteConsoleMsg(Userindex, "La quest ya esta en curso.", FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub
        End If
        
        
        
        'El personaje completo la quest que requiere?
        If QuestList(Npclist(NpcIndex).QuestNumber(Indice)).RequiredQuest > 0 Then
            If Not UserDoneQuest(Userindex, QuestList(Npclist(NpcIndex).QuestNumber(Indice)).RequiredQuest) Then
                Call WriteChatOverHead(Userindex, "Debes completas la quest " & QuestList(QuestList(Npclist(NpcIndex).QuestNumber(Indice)).RequiredQuest).nombre & " para emprender esta mision.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
                Exit Sub
            End If
        End If
        

        'El personaje tiene suficiente nivel?
        If UserList(Userindex).Stats.ELV < QuestList(Npclist(NpcIndex).QuestNumber(Indice)).RequiredLevel Then
            Call WriteChatOverHead(Userindex, "Debes ser por lo menos nivel " & QuestList(Npclist(NpcIndex).QuestNumber(Indice)).RequiredLevel & " para emprender esta mision.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
            Exit Sub
        End If
        
        
        'El personaje ya hizo la quest?
        If UserDoneQuest(Userindex, Npclist(NpcIndex).QuestNumber(Indice)) Then
            Call WriteChatOverHead(Userindex, "QUESTNEXT*" & Npclist(NpcIndex).QuestNumber(Indice), Npclist(NpcIndex).Char.CharIndex, vbYellow)
            Exit Sub
        End If
    
110     QuestSlot = FreeQuestSlot(Userindex)


        If QuestSlot = 0 Then
            Call WriteChatOverHead(Userindex, "Debes completar las misiones en curso para poder aceptar más misiones.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
            Exit Sub
        End If
        
        
        



    
        'Agregamos la quest.
112     With UserList(Userindex).QuestStats.Quests(QuestSlot)
114         .QuestIndex = Npclist(NpcIndex).QuestNumber(Indice)
        
116         If QuestList(.QuestIndex).RequiredNPCs Then ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
            If QuestList(.QuestIndex).RequiredTargetNPCs Then ReDim .NPCsTarget(1 To QuestList(.QuestIndex).RequiredTargetNPCs)
118         Call WriteConsoleMsg(Userindex, "Has aceptado la mision " & Chr(34) & QuestList(.QuestIndex).nombre & Chr(34) & ".", FontTypeNames.FONTTYPE_INFOIAO)
            Call WriteUpdateNPCSimbolo(Userindex, NpcIndex, 4)
        
        End With

        
        Exit Sub

HandleQuestAccept_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleQuestAccept", Erl)
        Resume Next
        
End Sub

Public Sub HandleQuestDetailsRequest(ByVal Userindex As Integer)
        
        On Error GoTo HandleQuestDetailsRequest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestInfoRequest.
        'Last modified: 30/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim QuestSlot As Byte
 
        'Leemos el paquete
100     Call UserList(Userindex).incomingData.ReadInteger
    
102     QuestSlot = UserList(Userindex).incomingData.ReadByte
    
104     Call WriteQuestDetails(Userindex, UserList(Userindex).QuestStats.Quests(QuestSlot).QuestIndex, QuestSlot)

        
        Exit Sub

HandleQuestDetailsRequest_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleQuestDetailsRequest", Erl)
        Resume Next
        
End Sub
 
Public Sub HandleQuestAbandon(ByVal Userindex As Integer)
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestAbandon.
        'Last modified: 31/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Leemos el paquete.
        
        On Error GoTo HandleQuestAbandon_Err
        
100     Call UserList(Userindex).incomingData.ReadInteger
    
        'Borramos la quest.
102     Call CleanQuestSlot(Userindex, UserList(Userindex).incomingData.ReadByte)
    
        'Ordenamos la lista de quests del usuario.
104     Call ArrangeUserQuests(Userindex)
    
        'Enviamos la lista de quests actualizada.
106     Call WriteQuestListSend(Userindex)

        
        Exit Sub

HandleQuestAbandon_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleQuestAbandon", Erl)
        Resume Next
        
End Sub

Public Sub HandleQuestListRequest(ByVal Userindex As Integer)
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestListRequest.
        'Last modified: 30/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        
        On Error GoTo HandleQuestListRequest_Err
        
 
        'Leemos el paquete
100     Call UserList(Userindex).incomingData.ReadInteger
    
102     If UserList(Userindex).flags.BattleModo = 0 Then
104         Call WriteQuestListSend(Userindex)
        Else
106         Call WriteConsoleMsg(Userindex, "No disponible aquí.", FontTypeNames.FONTTYPE_INFOIAO)

        End If

        
        Exit Sub

HandleQuestListRequest_Err:
        Call RegistrarError(Err.Number, Err.description, "Protocol.HandleQuestListRequest", Erl)
        Resume Next
        
End Sub

Public Sub WriteQuestDetails(ByVal Userindex As Integer, ByVal QuestIndex As Integer, Optional QuestSlot As Byte = 0)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Envía el paquete QuestDetails y la informaciín correspondiente.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
 
    On Error GoTo ErrHandler

    With UserList(Userindex).outgoingData
        'ID del paquete
        Call .WriteByte(ServerPacketID.QuestDetails)
        
        'Se usa la variable QuestSlot para saber si enviamos la info de una quest ya empezada o la info de una quest que no se aceptí todavía (1 para el primer caso y 0 para el segundo)
        Call .WriteByte(IIf(QuestSlot, 1, 0))
        
        'Enviamos nombre, descripciín y nivel requerido de la quest
        'Call .WriteASCIIString(QuestList(QuestIndex).Nombre)
        'Call .WriteASCIIString(QuestList(QuestIndex).Desc)
        Call .WriteInteger(QuestIndex)
        Call .WriteByte(QuestList(QuestIndex).RequiredLevel)
        
        Call .WriteInteger(QuestList(QuestIndex).RequiredQuest)
        
        'Enviamos la cantidad de npcs requeridos
        Call .WriteByte(QuestList(QuestIndex).RequiredNPCs)

        If QuestList(QuestIndex).RequiredNPCs Then

            'Si hay npcs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredNPCs
                Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).Amount)
                Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).NpcIndex)

                'Si es una quest ya empezada, entonces mandamos los NPCs que matí.
                If QuestSlot Then
                    Call .WriteInteger(UserList(Userindex).QuestStats.Quests(QuestSlot).NPCsKilled(i))

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
                
                'escribe si tiene ese objeto en el inventario y que cantidad
                Call .WriteInteger(CantidadObjEnInv(Userindex, QuestList(QuestIndex).RequiredOBJ(i).ObjIndex))
               ' Call .WriteInteger(0)
                
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
 
ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub
 
Public Sub WriteQuestListSend(ByVal Userindex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Envía el paquete QuestList y la informaciín correspondiente.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i       As Integer

    Dim tmpStr  As String

    Dim tmpByte As Byte
 
    On Error GoTo ErrHandler
 
    With UserList(Userindex)
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
 
ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

Public Sub WriteNpcQuestListSend(ByVal Userindex As Integer, ByVal NpcIndex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Envía el paquete QuestList y la informaciín correspondiente.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i       As Integer
    Dim j       As Integer

    Dim tmpStr  As String

    Dim tmpByte As Byte
 
    On Error GoTo ErrHandler
    
    Dim QuestIndex As Integer
    
    
    
 
    With UserList(Userindex).outgoingData
        .WriteByte ServerPacketID.NpcQuestListSend
        
        
        Call .WriteByte(Npclist(NpcIndex).NumQuest) 'Escribimos primero cuantas quest tiene el NPC
    
        For j = 1 To Npclist(NpcIndex).NumQuest
        
        QuestIndex = Npclist(NpcIndex).QuestNumber(j)
            
        Call .WriteInteger(QuestIndex)
        Call .WriteByte(QuestList(QuestIndex).RequiredLevel)
        
        
        Call .WriteInteger(QuestList(QuestIndex).RequiredQuest)
        
        'Enviamos la cantidad de npcs requeridos
        Call .WriteByte(QuestList(QuestIndex).RequiredNPCs)

        If QuestList(QuestIndex).RequiredNPCs Then

            'Si hay npcs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredNPCs
                Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).Amount)
                Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).NpcIndex)

                'Si es una quest ya empezada, entonces mandamos los NPCs que matí.
                'If QuestSlot Then
                   ' Call .WriteInteger(UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i))

               ' End If

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
        
        
        'Enviamos el estado de la QUEST
        '0 Disponible
        '1 EN CURSO
        '2 REALIZADA
        '3 no puede hacerla
        
        Dim PuedeHacerla As Boolean
        
        'La tiene aceptada el usuario?
        If TieneQuest(Userindex, QuestIndex) Then
            Call .WriteByte(1)
        Else
            If UserDoneQuest(Userindex, QuestIndex) Then
                Call .WriteByte(2)
            Else
                PuedeHacerla = True
                If QuestList(QuestIndex).RequiredQuest > 0 Then
                    If Not UserDoneQuest(Userindex, QuestList(QuestIndex).RequiredQuest) Then
                        PuedeHacerla = False
                    End If
                End If
                
                If UserList(Userindex).Stats.ELV < QuestList(QuestIndex).RequiredLevel Then
                    PuedeHacerla = False
                End If
                
                If PuedeHacerla Then
                    Call .WriteByte(0)
                Else
                    Call .WriteByte(3)
                End If
                
            End If
                
        End If
                 

        Next j
        
        'Escribimos la cantidad de quests
       ' Call .WriteByte(tmpByte)
        
        'Escribimos la lista de quests (sacamos el íltimo caracter)
       ' If tmpByte Then
         '   Call .WriteASCIIString(Left$(tmpStr, Len(tmpStr) - 1))

       ' End If

    End With

    Exit Sub
 
ErrHandler:

    If Err.Number = UserList(Userindex).outgoingData.NotEnoughSpaceErrCode Then
        
        Resume

    End If

End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreatePretorianClan(ByVal Userindex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    '***************************************************

    On Error GoTo ErrHandler

    Dim Map   As Integer
    Dim X     As Byte
    Dim Y     As Byte
    Dim Index As Long
    
    With UserList(Userindex)
        
        'Remove packet ID
        Call .incomingData.ReadInteger
        
        Map = .incomingData.ReadInteger()
        X = .incomingData.ReadByte()
        Y = .incomingData.ReadByte()
        
        ' User Admin?
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) = 0 Then Exit Sub
        
        ' Valid pos?
        If Not InMapBounds(Map, X, Y) Then
            Call WriteConsoleMsg(Userindex, "Posicion invalida.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        ' Choose pretorian clan index
        If Map = MAPA_PRETORIANO Then
            Index = ePretorianType.Default ' Default clan
        Else
            Index = ePretorianType.Custom ' Custom Clan
        End If
            
        ' Is already active any clan?
        If Not ClanPretoriano(Index).Active Then
            
            If Not ClanPretoriano(Index).SpawnClan(Map, X, Y, Index) Then
                Call WriteConsoleMsg(Userindex, "La posicion no es apropiada para crear el clan", FontTypeNames.FONTTYPE_INFO)

            End If
        
        Else
            Call WriteConsoleMsg(Userindex, "El clan pretoriano se encuentra activo en el mapa " & ClanPretoriano(Index).ClanMap & ". Utilice /EliminarPretorianos MAPA y reintente.", FontTypeNames.FONTTYPE_INFO)

        End If
    
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en HandleCreatePretorianClan. Error: " & Err.Number & " - " & Err.description)

End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDeletePretorianClan(ByVal Userindex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    '***************************************************

    On Error GoTo ErrHandler
    
    Dim Map   As Integer
    Dim Index As Long
    
    With UserList(Userindex)
        
        'Remove packet ID
        Call .incomingData.ReadInteger
        
        Map = .incomingData.ReadInteger()
        
        ' User Admin?
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) = 0 Then Exit Sub
        
        ' Valid map?
        If Map < 1 Or Map > NumMaps Then
            Call WriteConsoleMsg(Userindex, "Mapa invalido.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Los sacamos correctamente.
        Call EliminarPretorianos(Map)
    
    End With

    Exit Sub

ErrHandler:
    Call LogError("Error en HandleDeletePretorianClan. Error: " & Err.Number & " - " & Err.description)

End Sub

