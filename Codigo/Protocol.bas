Attribute VB_Name = "Protocol"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P

Public Const SEPARATOR As String * 1 = vbNullChar
Private Const SPELL_UNASSISTED_DARDO = 1
Private Const SPELL_UNASSISTED_RUGIDO_SALVAJE = 5
Private Const SPELL_UNASSISTED_RUGIDO_ARCANO = 348
Private Const SPELL_UNASSISTED_FULGOR_IGNEO = 52
Private Const SPELL_UNASSISTED_LATIDO_IGNEO = 349
Private Const SPELL_UNASSISTED_ECO_IGNEO = 61
Private Const SPELL_UNASSISTED_DESTELLO_MALVA = 62
Private Const SPELL_UNASSISTED_FRACTURA_GLACIAL = 63
Private Const SPELL_UNASSISTED_ALIENTO_CARMESI = 64
Private Const SPELL_UNASSISTED_ENERGIA_ANCESTRAL = 65


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
    FONTTYPE_CITIZEN_ARMADA
    FONTTYPE_CRIMINAL
    FONTTYPE_CRIMINAL_CAOS
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

Public Type t_PersonajeCuenta
    nombre As String
    nivel As Byte
    Mapa As Integer
    PosX As Integer
    PosY As Integer
    cuerpo As Integer
    Cabeza As Integer
    Status As Byte
    clase As Byte
    Arma As Integer
    Escudo As Integer
    Casco As Integer
    ClanIndex As Integer
    BackPack As Integer
End Type


#If DIRECT_PLAY = 0 Then
Public reader As Network.reader

#If LOGIN_STRESS_TEST = 1 And PYMMO = 1 Then
Private mStressSeq As Long

Private Function Stress_Enabled() As Boolean
    ' If you want a runtime toggle, read from INI/env here.
    Stress_Enabled = True
End Function

Private Function Stress_NextName() As String
    mStressSeq = mStressSeq + 1
    Stress_NextName = "STRESS_" & Format$(mStressSeq, "000000")
End Function

Private Function Stress_IsToken(ByVal s As String) As Boolean
    ' Minimal detector: anything starting with "STRESS." is treated as a stress login
    Stress_IsToken = (Len(s) >= 7 And Left$(s, 7) = "STRESS.")
End Function
#End If


Public Sub InitializePacketList()
    Call Protocol_Writes.InitializeAuxiliaryBuffer
End Sub

Public Function HandleIncomingData(ByVal ConnectionID As Long, ByVal Message As Network.reader, Optional ByVal optional_user_index As Variant) As Boolean
#Else
    Public reader As New clsNetReader
    Public Function HandleIncomingData(ByVal ConnectionID As Long, Message As DxVBLibA.DPNMSG_RECEIVE, Optional ByVal optional_user_index As Variant) As Boolean
    #End If
    On Error Resume Next
    Dim UserIndex As Integer
    If Not IsMissing(optional_user_index) Then
        UserIndex = CInt(optional_user_index)
    Else
        UserIndex = 0
    End If
    #If DIRECT_PLAY = 0 Then
        Set reader = Message
    #Else
        reader.set_data Message
    #End If
    Dim PacketId As Long
    PacketId = reader.ReadInt16
    Dim actual_time       As Long
    Dim performance_timer As Long
    actual_time = GetTickCountRaw()
    performance_timer = actual_time
    #If DIRECT_PLAY = 0 Then
        If TicksElapsed(Mapping(ConnectionID).TimeLastReset, actual_time) >= 5000 Then
            Mapping(ConnectionID).TimeLastReset = actual_time
            Mapping(ConnectionID).PacketCount = 0
        End If
        Mapping(ConnectionID).PacketCount = Mapping(ConnectionID).PacketCount + 1
        If Mapping(ConnectionID).PacketCount > 100 Then
            'Lo kickeo
            If UserIndex > 0 Then
                If Not IsMissing(optional_user_index) Then ' userindex may be invalid here
                    Call SendData(SendTarget.ToAdminsYDioses, UserIndex, PrepareMessageConsoleMsg("Control Paquetes---> El usuario " & UserList(UserIndex).name & _
                            " | Iteración paquetes | Último paquete: " & PacketId & ".", e_FontTypeNames.FONTTYPE_FIGHT))
                End If
                Mapping(ConnectionID).PacketCount = 0
                If IsFeatureEnabled("kick_packet_overflow") Then
                    Call KickConnection(ConnectionID)
                End If
            Else
                If Not IsMissing(optional_user_index) Then ' userindex may be invalid here
                    Call SendData(SendTarget.ToAdminsYDioses, UserIndex, PrepareMessageConsoleMsg( _
                            "Control Paquetes---> Usuario desconocido | Iteración paquetes | Último paquete: " & PacketId & ".", e_FontTypeNames.FONTTYPE_FIGHT))
                End If
                Mapping(ConnectionID).PacketCount = 0
                If IsFeatureEnabled("kick_packet_overflow") Then
                    Call KickConnection(ConnectionID)
                End If
            End If
            Exit Function
        End If
    #End If
    If PacketId < ClientPacketID.eMinPacket Or PacketId >= ClientPacketID.PacketCount Then
        If Not IsMissing(optional_user_index) Then ' userindex may be invalid here
            Call LogEdicionPaquete("El usuario " & UserList(UserIndex).ConnectionDetails.IP & " mando fake paquet " & PacketId)
            Call SendData(SendTarget.ToGM, UserIndex, PrepareMessageConsoleMsg("Control Paquetes---> El usuario " & UserList(UserIndex).name & " | IP: " & UserList( _
                    UserIndex).ConnectionDetails.IP & " ESTÁ ENVIANDO PAQUETES INVÁLIDOS", e_FontTypeNames.FONTTYPE_GUILD))
        End If
        Call KickConnection(ConnectionID)
        Exit Function
    End If
    #If PYMMO = 1 Then
        'Does the packet requires a logged user??
        If Not (PacketId = ClientPacketID.eLoginExistingChar Or PacketId = ClientPacketID.eLoginNewChar) Then
            ' All these packets require a logged user with a valid user_index
            Debug.Assert Not IsMissing(optional_user_index)
            If Not IsMissing(optional_user_index) Then ' userindex may be invalid here
                'Is the user actually logged?
                If Not UserList(UserIndex).flags.UserLogged Then
                    Call CloseSocket(UserIndex)
                    Exit Function
                    'He is logged. Reset idle counter if id is valid.
                ElseIf PacketId <= ClientPacketID.[PacketCount] Then
                    UserList(UserIndex).Counters.IdleCount = 0
                End If
            Else
                'If UserIndex is missing then kick out
                Call KickConnection(ConnectionID)
            End If
        Else
            'Got eLoginExistingChar/eLoginNewChar, here UserIndex must not be assigned
            Debug.Assert IsMissing(optional_user_index)
            If Not IsMissing(optional_user_index) Then
                'If UserIndex is not missing then kick out
                Call KickConnection(ConnectionID)
            End If
        End If
    #ElseIf PYMMO = 0 Then
        'Does the packet requires a logged account??
        If Not (PacketId = ClientPacketID.eCreateAccount Or PacketId = ClientPacketID.eLoginAccount) Then
            'Is the account actually logged?
            If UserList(UserIndex).AccountID = 0 Then
                Call CloseSocket(UserIndex)
                Exit Function
            End If
            If Not (PacketId = ClientPacketID.eLoginExistingChar Or PacketId = ClientPacketID.eLoginNewChar) Then
                'Is the user actually logged?
                If Not UserList(UserIndex).flags.UserLogged Then
                    Call CloseSocket(UserIndex)
                    Exit Function
                    'He is logged. Reset idle counter if id is valid.
                ElseIf PacketId <= ClientPacketID.[PacketCount] Then
                    UserList(UserIndex).Counters.IdleCount = 0
                End If
            End If
        End If
    #End If
    Select Case PacketId
        Case ClientPacketID.eLoginExistingChar
            Call HandleLoginExistingChar(ConnectionID)
        Case ClientPacketID.eLoginNewChar
            Call HandleLoginNewChar(ConnectionID)
        Case ClientPacketID.eWalk
            Call HandleWalk(UserIndex)
        Case ClientPacketID.eAttack
            Call HandleAttack(UserIndex)
        Case ClientPacketID.eTalk
            Call HandleTalk(UserIndex)
        Case ClientPacketID.eYell
            Call HandleYell(UserIndex)
        Case ClientPacketID.eWhisper
            Call HandleWhisper(UserIndex)
        Case ClientPacketID.eRequestPositionUpdate
            Call HandleRequestPositionUpdate(UserIndex)
        Case ClientPacketID.ePickUp
            Call HandlePickUp(UserIndex)
        Case ClientPacketID.eSafeToggle
            Call HandleSafeToggle(UserIndex)
        Case ClientPacketID.ePartySafeToggle
            Call HandlePartyToggle(UserIndex)
        Case ClientPacketID.eRequestGuildLeaderInfo
            Call HandleRequestGuildLeaderInfo(UserIndex)
        Case ClientPacketID.eRequestAtributes
            Call HandleRequestAtributes(UserIndex)
        Case ClientPacketID.eRequestSkills
            Call HandleRequestSkills(UserIndex)
        Case ClientPacketID.eRequestMiniStats
            Call HandleRequestMiniStats(UserIndex)
        Case ClientPacketID.eCommerceEnd
            Call HandleCommerceEnd(UserIndex)
        Case ClientPacketID.eUserCommerceEnd
            Call HandleUserCommerceEnd(UserIndex)
        Case ClientPacketID.eBankEnd
            Call HandleBankEnd(UserIndex)
        Case ClientPacketID.eUserCommerceOk
            Call HandleUserCommerceOk(UserIndex)
        Case ClientPacketID.eUserCommerceReject
            Call HandleUserCommerceReject(UserIndex)
        Case ClientPacketID.eDrop
            Call HandleDrop(UserIndex)
        Case ClientPacketID.eCastSpell
            Call HandleCastSpell(UserIndex) ', crc)
        Case ClientPacketID.eLeftClick
            Call HandleLeftClick(UserIndex)
        Case ClientPacketID.eDoubleClick
            Call HandleDoubleClick(UserIndex)
        Case ClientPacketID.eWork
            Call HandleWork(UserIndex)
        Case ClientPacketID.eUseSpellMacro
            Call HandleUseSpellMacro(UserIndex)
        Case ClientPacketID.eUseItem
            Call HandleUseItem(UserIndex)
        Case ClientPacketID.eUseItemU
            Call HandleUseItemU(UserIndex)
        Case ClientPacketID.eCraftBlacksmith
            Call HandleCraftBlacksmith(UserIndex)
        Case ClientPacketID.eCraftCarpenter
            Call HandleCraftCarpenter(UserIndex)
        Case ClientPacketID.eWorkLeftClick
            Call HandleWorkLeftClick(UserIndex)
        Case ClientPacketID.eStartAutomatedAction
            Call HandleStartAutomatedAction(UserIndex)
        Case ClientPacketID.eCreateNewGuild
            Call HandleCreateNewGuild(UserIndex)
        Case ClientPacketID.eSpellInfo
            Call HandleSpellInfo(UserIndex)
        Case ClientPacketID.eEquipItem
            Call HandleEquipItem(UserIndex)
        Case ClientPacketID.eChangeHeading
            Call HandleChange_Heading(UserIndex)
        Case ClientPacketID.eModifySkills
            Call HandleModifySkills(UserIndex)
        Case ClientPacketID.eTrain
            Call HandleTrain(UserIndex)
        Case ClientPacketID.eCommerceBuy
            Call HandleCommerceBuy(UserIndex)
        Case ClientPacketID.eBankExtractItem
            Call HandleBankExtractItem(UserIndex)
        Case ClientPacketID.eCommerceSell
            Call HandleCommerceSell(UserIndex)
        Case ClientPacketID.eBankDeposit
            Call HandleBankDeposit(UserIndex)
        Case ClientPacketID.eForumPost
            Call HandleForumPost(UserIndex)
        Case ClientPacketID.eMoveSpell
            Call HandleMoveSpell(UserIndex)
        Case ClientPacketID.eClanCodexUpdate
            Call HandleClanCodexUpdate(UserIndex)
        Case ClientPacketID.eUserCommerceOffer
            Call HandleUserCommerceOffer(UserIndex)
        Case ClientPacketID.eGuildAcceptPeace
            Call HandleGuildAcceptPeace(UserIndex)
        Case ClientPacketID.eGuildRejectAlliance
            Call HandleGuildRejectAlliance(UserIndex)
        Case ClientPacketID.eGuildRejectPeace
            Call HandleGuildRejectPeace(UserIndex)
        Case ClientPacketID.eGuildAcceptAlliance
            Call HandleGuildAcceptAlliance(UserIndex)
        Case ClientPacketID.eGuildOfferPeace
            Call HandleGuildOfferPeace(UserIndex)
        Case ClientPacketID.eGuildOfferAlliance
            Call HandleGuildOfferAlliance(UserIndex)
        Case ClientPacketID.eGuildAllianceDetails
            Call HandleGuildAllianceDetails(UserIndex)
        Case ClientPacketID.eGuildPeaceDetails
            Call HandleGuildPeaceDetails(UserIndex)
        Case ClientPacketID.eGuildRequestJoinerInfo
            Call HandleGuildRequestJoinerInfo(UserIndex)
        Case ClientPacketID.eGuildAlliancePropList
            Call HandleGuildAlliancePropList(UserIndex)
        Case ClientPacketID.eGuildPeacePropList
            Call HandleGuildPeacePropList(UserIndex)
        Case ClientPacketID.eGuildDeclareWar
            Call HandleGuildDeclareWar(UserIndex)
        Case ClientPacketID.eGuildNewWebsite
            Call HandleGuildNewWebsite(UserIndex)
        Case ClientPacketID.eGuildAcceptNewMember
            Call HandleGuildAcceptNewMember(UserIndex)
        Case ClientPacketID.eGuildRejectNewMember
            Call HandleGuildRejectNewMember(UserIndex)
        Case ClientPacketID.eGuildKickMember
            Call HandleGuildKickMember(UserIndex)
        Case ClientPacketID.eGuildUpdateNews
            Call HandleGuildUpdateNews(UserIndex)
        Case ClientPacketID.eGuildMemberInfo
            Call HandleGuildMemberInfo(UserIndex)
        Case ClientPacketID.eGuildOpenElections
            Call HandleGuildOpenElections(UserIndex)
        Case ClientPacketID.eGuildRequestMembership
            Call HandleGuildRequestMembership(UserIndex)
        Case ClientPacketID.eGuildRequestDetails
            Call HandleGuildRequestDetails(UserIndex)
        Case ClientPacketID.eOnline
            Call HandleOnline(UserIndex)
        Case ClientPacketID.eQuit
            Call HandleQuit(UserIndex)
        Case ClientPacketID.eGuildLeave
            Call HandleGuildLeave(UserIndex)
        Case ClientPacketID.eRequestAccountState
            Call HandleRequestAccountState(UserIndex)
        Case ClientPacketID.ePetStand
            Call HandlePetStand(UserIndex)
        Case ClientPacketID.ePetFollow
            Call HandlePetFollow(UserIndex)
        Case ClientPacketID.ePetFollowAll
            Call HandlePetFollowAll(UserIndex)
        Case ClientPacketID.ePetLeave
            Call HandlePetLeave(UserIndex)
        Case ClientPacketID.eGrupoMsg
            Call HandleGrupoMsg(UserIndex)
        Case ClientPacketID.eTrainList
            Call HandleTrainList(UserIndex)
        Case ClientPacketID.eRest
            Call HandleRest(UserIndex)
        Case ClientPacketID.eMeditate
            Call HandleMeditate(UserIndex)
        Case ClientPacketID.eResucitate
            Call HandleResucitate(UserIndex)
        Case ClientPacketID.eHeal
            Call HandleHeal(UserIndex)
        Case ClientPacketID.eHelp
            Call HandleHelp(UserIndex)
        Case ClientPacketID.eRequestStats
            Call HandleRequestStats(UserIndex)
        Case ClientPacketID.eCommerceStart
            Call HandleCommerceStart(UserIndex)
        Case ClientPacketID.eBankStart
            Call HandleBankStart(UserIndex)
        Case ClientPacketID.eEnlist
            Call HandleEnlist(UserIndex)
        Case ClientPacketID.eInformation
            Call HandleInformation(UserIndex)
        Case ClientPacketID.eReward
            Call HandleReward(UserIndex)
        Case ClientPacketID.eRequestMOTD
            Call HandleRequestMOTD(UserIndex)
        Case ClientPacketID.eUpTime
            Call HandleUpTime(UserIndex)
        Case ClientPacketID.eGuildMessage
            Call HandleGuildMessage(UserIndex)
        Case ClientPacketID.eGuildOnline
            Call HandleGuildOnline(UserIndex)
        Case ClientPacketID.eCouncilMessage
            Call HandleCouncilMessage(UserIndex)
        Case ClientPacketID.eRoleMasterRequest
            Call HandleRoleMasterRequest(UserIndex)
        Case ClientPacketID.eChangeDescription
            Call HandleChangeDescription(UserIndex)
        Case ClientPacketID.eGuildVote
            Call HandleGuildVote(UserIndex)
        Case ClientPacketID.epunishments
            Call HandlePunishments(UserIndex)
        Case ClientPacketID.eGamble
            Call HandleGamble(UserIndex)
        Case ClientPacketID.eMapPriceEntrance
            Call HandleMapPriceEntrance(UserIndex)
        Case ClientPacketID.eLeaveFaction
            Call HandleLeaveFaction(UserIndex)
        Case ClientPacketID.eBankExtractGold
            Call HandleBankExtractGold(UserIndex)
        Case ClientPacketID.eBankDepositGold
            Call HandleBankDepositGold(UserIndex)
        Case ClientPacketID.eDenounce
            Call HandleDenounce(UserIndex)
        Case ClientPacketID.eGMMessage
            Call HandleGMMessage(UserIndex)
        Case ClientPacketID.eshowName
            Call HandleShowName(UserIndex)
        Case ClientPacketID.eOnlineRoyalArmy
            Call HandleOnlineRoyalArmy(UserIndex)
        Case ClientPacketID.eOnlineChaosLegion
            Call HandleOnlineChaosLegion(UserIndex)
        Case ClientPacketID.eGoNearby
            Call HandleGoNearby(UserIndex)
        Case ClientPacketID.ecomment
            Call HandleComment(UserIndex)
        Case ClientPacketID.eWhere
            Call HandleWhere(UserIndex)
        Case ClientPacketID.eCreaturesInMap
            Call HandleCreaturesInMap(UserIndex)
        Case ClientPacketID.eWarpMeToTarget
            Call HandleWarpMeToTarget(UserIndex)
        Case ClientPacketID.eWarpChar
            Call HandleWarpChar(UserIndex)
        Case ClientPacketID.eSilence
            Call HandleSilence(UserIndex)
        Case ClientPacketID.eSOSShowList
            Call HandleSOSShowList(UserIndex)
        Case ClientPacketID.eSOSRemove
            Call HandleSOSRemove(UserIndex)
        Case ClientPacketID.eGoToChar
            Call HandleGoToChar(UserIndex)
        Case ClientPacketID.einvisible
            Call HandleInvisible(UserIndex)
        Case ClientPacketID.eGMPanel
            Call HandleGMPanel(UserIndex)
        Case ClientPacketID.eRequestUserList
            Call HandleRequestUserList(UserIndex)
        Case ClientPacketID.eWorking
            Call HandleWorking(UserIndex)
        Case ClientPacketID.eHiding
            Call HandleHiding(UserIndex)
        Case ClientPacketID.eJail
            Call HandleJail(UserIndex)
        Case ClientPacketID.eKillNPC
            Call HandleKillNPC(UserIndex)
        Case ClientPacketID.eWarnUser
            Call HandleWarnUser(UserIndex)
        Case ClientPacketID.eEditChar
            Call HandleEditChar(UserIndex)
        Case ClientPacketID.eRequestCharInfo
            Call HandleRequestCharInfo(UserIndex)
        Case ClientPacketID.eRequestCharStats
            Call HandleRequestCharStats(UserIndex)
        Case ClientPacketID.eRequestCharGold
            Call HandleRequestCharGold(UserIndex)
        Case ClientPacketID.eRequestCharInventory
            Call HandleRequestCharInventory(UserIndex)
        Case ClientPacketID.eRequestCharBank
            Call HandleRequestCharBank(UserIndex)
        Case ClientPacketID.eRequestCharSkills
            Call HandleRequestCharSkills(UserIndex)
        Case ClientPacketID.eReviveChar
            Call HandleReviveChar(UserIndex)
        Case ClientPacketID.eNotifyInventarioHechizos
            Call HandleNotifyInventariohechizos(UserIndex)
        Case ClientPacketID.eOnlineGM
            Call HandleOnlineGM(UserIndex)
        Case ClientPacketID.eOnlineMap
            Call HandleOnlineMap(UserIndex)
        Case ClientPacketID.eForgive
            Call HandleForgive(UserIndex)
        Case ClientPacketID.ePerdonFaccion
            Call HandlePerdonFaccion(UserIndex)
        Case ClientPacketID.eStartEvent
            Call HandleStartEvent(UserIndex)
        Case ClientPacketID.eCancelarEvento
            Call HandleCancelarEvento(UserIndex)
        Case ClientPacketID.eKick
            Call HandleKick(UserIndex)
        Case ClientPacketID.eExecute
            Call HandleExecute(UserIndex)
        Case ClientPacketID.eBanChar
            Call HandleBanChar(UserIndex)
        Case ClientPacketID.eUnbanChar
            Call HandleUnbanChar(UserIndex)
        Case ClientPacketID.eNPCFollow
            Call HandleNPCFollow(UserIndex)
        Case ClientPacketID.eSummonChar
            Call HandleSummonChar(UserIndex)
        Case ClientPacketID.eSpawnListRequest
            Call HandleSpawnListRequest(UserIndex)
        Case ClientPacketID.eSpawnCreature
            Call HandleSpawnCreature(UserIndex)
        Case ClientPacketID.eResetNPCInventory
            Call HandleResetNPCInventory(UserIndex)
        Case ClientPacketID.eCleanWorld
            Call HandleCleanWorld(UserIndex)
        Case ClientPacketID.eServerMessage
            Call HandleServerMessage(UserIndex)
        Case ClientPacketID.eNickToIP
            Call HandleNickToIP(UserIndex)
        Case ClientPacketID.eIPToNick
            Call HandleIPToNick(UserIndex)
        Case ClientPacketID.eGuildOnlineMembers
            Call HandleGuildOnlineMembers(UserIndex)
        Case ClientPacketID.eTeleportCreate
            Call HandleTeleportCreate(UserIndex)
        Case ClientPacketID.eTeleportDestroy
            Call HandleTeleportDestroy(UserIndex)
        Case ClientPacketID.eRainToggle
            Call HandleRainToggle(UserIndex)
        Case ClientPacketID.eSetCharDescription
            Call HandleSetCharDescription(UserIndex)
        Case ClientPacketID.eForceMIDIToMap
            Call HanldeForceMIDIToMap(UserIndex)
        Case ClientPacketID.eForceWAVEToMap
            Call HandleForceWAVEToMap(UserIndex)
        Case ClientPacketID.eRoyalArmyMessage
            Call HandleRoyalArmyMessage(UserIndex)
        Case ClientPacketID.eChaosLegionMessage
            Call HandleChaosLegionMessage(UserIndex)
        Case ClientPacketID.eTalkAsNPC
            Call HandleTalkAsNPC(UserIndex)
        Case ClientPacketID.eDestroyAllItemsInArea
            Call HandleDestroyAllItemsInArea(UserIndex)
        Case ClientPacketID.eAcceptRoyalCouncilMember
            Call HandleAcceptRoyalCouncilMember(UserIndex)
        Case ClientPacketID.eAcceptChaosCouncilMember
            Call HandleAcceptChaosCouncilMember(UserIndex)
        Case ClientPacketID.eItemsInTheFloor
            Call HandleItemsInTheFloor(UserIndex)
        Case ClientPacketID.eMakeDumb
            Call HandleMakeDumb(UserIndex)
        Case ClientPacketID.eMakeDumbNoMore
            Call HandleMakeDumbNoMore(UserIndex)
        Case ClientPacketID.eCouncilKick
            Call HandleCouncilKick(UserIndex)
        Case ClientPacketID.eSetTrigger
            Call HandleSetTrigger(UserIndex)
        Case ClientPacketID.eAskTrigger
            Call HandleAskTrigger(UserIndex)
        Case ClientPacketID.eGuildMemberList
            Call HandleGuildMemberList(UserIndex)
        Case ClientPacketID.eGuildBan
            Call HandleGuildBan(UserIndex)
        Case ClientPacketID.eCreateItem
            Call HandleCreateItem(UserIndex)
        Case ClientPacketID.eDestroyItems
            Call HandleDestroyItems(UserIndex)
        Case ClientPacketID.eChaosLegionKick
            Call HandleChaosLegionKick(UserIndex)
        Case ClientPacketID.eRoyalArmyKick
            Call HandleRoyalArmyKick(UserIndex)
        Case ClientPacketID.eForceMIDIAll
            Call HandleForceMIDIAll(UserIndex)
        Case ClientPacketID.eForceWAVEAll
            Call HandleForceWAVEAll(UserIndex)
        Case ClientPacketID.eRemovePunishment
            Call HandleRemovePunishment(UserIndex)
        Case ClientPacketID.eTileBlockedToggle
            Call HandleTile_BlockedToggle(UserIndex)
        Case ClientPacketID.eKillNPCNoRespawn
            Call HandleKillNPCNoRespawn(UserIndex)
        Case ClientPacketID.eKillAllNearbyNPCs
            Call HandleKillAllNearbyNPCs(UserIndex)
        Case ClientPacketID.eLastIP
            Call HandleLastIP(UserIndex)
        Case ClientPacketID.eChangeMOTD
            Call HandleChangeMOTD(UserIndex)
        Case ClientPacketID.eSetMOTD
            Call HandleSetMOTD(UserIndex)
        Case ClientPacketID.eSystemMessage
            Call HandleSystemMessage(UserIndex)
        Case ClientPacketID.eCreateNPC
            Call HandleCreateNPC(UserIndex)
        Case ClientPacketID.eCreateNPCWithRespawn
            Call HandleCreateNPCWithRespawn(UserIndex)
        Case ClientPacketID.eImperialArmour
            Call HandleImperialArmour(UserIndex)
        Case ClientPacketID.eChaosArmour
            Call HandleChaosArmour(UserIndex)
        Case ClientPacketID.eNavigateToggle
            Call HandleNavigateToggle(UserIndex)
        Case ClientPacketID.eServerOpenToUsersToggle
            Call HandleServerOpenToUsersToggle(UserIndex)
        Case ClientPacketID.eParticipar
            Call HandleParticipar(UserIndex)
        Case ClientPacketID.eTurnCriminal
            Call HandleTurnCriminal(UserIndex)
        Case ClientPacketID.eResetFactions
            Call HandleResetFactions(UserIndex)
        Case ClientPacketID.eRemoveCharFromGuild
            Call HandleRemoveCharFromGuild(UserIndex)
        Case ClientPacketID.eAlterName
            Call HandleAlterName(UserIndex)
        Case ClientPacketID.eDoBackUp
            Call HandleDoBackUp(UserIndex)
        Case ClientPacketID.eShowGuildMessages
            Call HandleShowGuildMessages(UserIndex)
        Case ClientPacketID.eChangeMapInfoPK
            Call HandleChangeMapInfoPK(UserIndex)
        Case ClientPacketID.eChangeMapInfoBackup
            Call HandleChangeMapInfoBackup(UserIndex)
        Case ClientPacketID.eChangeMapInfoRestricted
            Call HandleChangeMapInfoRestricted(UserIndex)
        Case ClientPacketID.eChangeMapInfoNoMagic
            Call HandleChangeMapInfoNoMagic(UserIndex)
        Case ClientPacketID.eChangeMapInfoNoInvi
            Call HandleChangeMapInfoNoInvi(UserIndex)
        Case ClientPacketID.eChangeMapInfoNoResu
            Call HandleChangeMapInfoNoResu(UserIndex)
        Case ClientPacketID.eChangeMapInfoLand
            Call HandleChangeMapInfoLand(UserIndex)
        Case ClientPacketID.eChangeMapInfoZone
            Call HandleChangeMapInfoZone(UserIndex)
        Case ClientPacketID.eChangeMapSetting
            Call HandleChangeMapSetting(UserIndex)
        Case ClientPacketID.eSaveChars
            Call HandleSaveChars(UserIndex)
        Case ClientPacketID.eCleanSOS
            Call HandleCleanSOS(UserIndex)
        Case ClientPacketID.eShowServerForm
            Call HandleShowServerForm(UserIndex)
        Case ClientPacketID.eKickAllChars
            Call HandleKickAllChars(UserIndex)
        Case ClientPacketID.eChatColor
            Call HandleChatColor(UserIndex)
        Case ClientPacketID.eIgnored
            Call HandleIgnored(UserIndex)
        Case ClientPacketID.eCheckSlot
            Call HandleCheckSlot(UserIndex)
        Case ClientPacketID.eSetSpeed
            Call HandleSetSpeed(UserIndex)
        Case ClientPacketID.eGlobalMessage
            Call HandleGlobalMessage(UserIndex)
        Case ClientPacketID.eGlobalOnOff
            Call HandleGlobalOnOff(UserIndex)
        Case ClientPacketID.eUseKey
            Call HandleUseKey(UserIndex)
        Case ClientPacketID.eDonateGold
            Call HandleDonateGold(UserIndex)
        Case ClientPacketID.ePromedio
            Call HandlePromedio(UserIndex)
        Case ClientPacketID.eGiveItem
            Call HandleGiveItem(UserIndex)
        Case ClientPacketID.eOfertaInicial
            Call HandleOfertaInicial(UserIndex)
        Case ClientPacketID.eOfertaDeSubasta
            Call HandleOfertaDeSubasta(UserIndex)
        Case ClientPacketID.eQuestionGM
            Call HandleQuestionGM(UserIndex)
        Case ClientPacketID.eCuentaRegresiva
            Call HandleCuentaRegresiva(UserIndex)
        Case ClientPacketID.ePossUser
            Call HandlePossUser(UserIndex)
        Case ClientPacketID.eDuel
            Call HandleDuel(UserIndex)
        Case ClientPacketID.eAcceptDuel
            Call HandleAcceptDuel(UserIndex)
        Case ClientPacketID.eCancelDuel
            Call HandleCancelDuel(UserIndex)
        Case ClientPacketID.eQuitDuel
            Call HandleQuitDuel(UserIndex)
        Case ClientPacketID.eNieveToggle
            Call HandleNieveToggle(UserIndex)
        Case ClientPacketID.eNieblaToggle
            Call HandleNieblaToggle(UserIndex)
        Case ClientPacketID.eTransFerGold
            Call HandleTransFerGold(UserIndex)
        Case ClientPacketID.eMoveitem
            Call HandleMoveItem(UserIndex)
        Case ClientPacketID.eGenio
            Call HandleGenio(UserIndex)
        Case ClientPacketID.eCasarse
            Call HandleCasamiento(UserIndex)
        Case ClientPacketID.eCraftAlquimista
            Call HandleCraftAlquimia(UserIndex)
        Case ClientPacketID.eFlagTrabajar
            Call HandleFlagTrabajar(UserIndex)
        Case ClientPacketID.eCraftSastre
            Call HandleCraftSastre(UserIndex)
        Case ClientPacketID.eMensajeUser
            Call HandleMensajeUser(UserIndex)
        Case ClientPacketID.eTraerBoveda
            Call HandleTraerBoveda(UserIndex)
        Case ClientPacketID.eCompletarAccion
            Call HandleCompletarAccion(UserIndex)
        Case ClientPacketID.eInvitarGrupo
            Call HandleInvitarGrupo(UserIndex)
        Case ClientPacketID.eResponderPregunta
            Call HandleResponderPregunta(UserIndex)
        Case ClientPacketID.eRequestGrupo
            Call HandleRequestGrupo(UserIndex)
        Case ClientPacketID.eAbandonarGrupo
            Call HandleAbandonarGrupo(UserIndex)
        Case ClientPacketID.eHecharDeGrupo
            Call HandleHecharDeGrupo(UserIndex)
        Case ClientPacketID.eMacroPossent
            Call HandleMacroPos(UserIndex)
        Case ClientPacketID.eSubastaInfo
            Call HandleSubastaInfo(UserIndex)
        Case ClientPacketID.eBanCuenta
            Call HandleBanCuenta(UserIndex)
        Case ClientPacketID.eUnbanCuenta
            Call HandleUnBanCuenta(UserIndex)
        Case ClientPacketID.eCerrarCliente
            Call HandleCerrarCliente(UserIndex)
        Case ClientPacketID.eEventoInfo
            Call HandleEventoInfo(UserIndex)
        Case ClientPacketID.eCrearEvento
            Call HandleCrearEvento(UserIndex)
        Case ClientPacketID.eBanTemporal
            Call HandleBanTemporal(UserIndex)
        Case ClientPacketID.eCancelarExit
            Call HandleCancelarExit(UserIndex)
        Case ClientPacketID.eCrearTorneo
            Call HandleCrearTorneo(UserIndex)
        Case ClientPacketID.eComenzarTorneo
            Call HandleComenzarTorneo(UserIndex)
        Case ClientPacketID.eCancelarTorneo
            Call HandleCancelarTorneo(UserIndex)
        Case ClientPacketID.eBusquedaTesoro
            Call HandleBusquedaTesoro(UserIndex)
        Case ClientPacketID.eCompletarViaje
            Call HandleCompletarViaje(UserIndex)
        Case ClientPacketID.eBovedaMoveItem
            Call HandleBovedaMoveItem(UserIndex)
        Case ClientPacketID.eQuieroFundarClan
            Call HandleQuieroFundarClan(UserIndex)
        Case ClientPacketID.ellamadadeclan
            Call HandleLlamadadeClan(UserIndex)
        Case ClientPacketID.eMarcaDeClanPack
            Call HandleMarcaDeClan(UserIndex)
        Case ClientPacketID.eMarcaDeGMPack
            Call HandleMarcaDeGM(UserIndex)
        Case ClientPacketID.eQuest
            Call HandleQuest(UserIndex)
        Case ClientPacketID.eQuestAccept
            Call HandleQuestAccept(UserIndex)
        Case ClientPacketID.eQuestListRequest
            Call HandleQuestListRequest(UserIndex)
        Case ClientPacketID.eQuestDetailsRequest
            Call HandleQuestDetailsRequest(UserIndex)
        Case ClientPacketID.eQuestAbandon
            Call HandleQuestAbandon(UserIndex)
        Case ClientPacketID.eSeguroClan
            Call HandleSeguroClan(UserIndex)
        Case ClientPacketID.ehome
            Call HandleHome(UserIndex)
        Case ClientPacketID.eConsulta
            Call HandleConsulta(UserIndex)
        Case ClientPacketID.eGetMapInfo
            Call HandleGetMapInfo(UserIndex)
        Case ClientPacketID.eFinEvento
            Call HandleFinEvento(UserIndex)
        Case ClientPacketID.eSeguroResu
            Call HandleSeguroResu(UserIndex)
        Case ClientPacketID.eLegionarySecure
            Call HandleLegionarySecure(UserIndex)
        Case ClientPacketID.eCuentaExtractItem
            Call HandleCuentaExtractItem(UserIndex)
        Case ClientPacketID.eCuentaDeposit
            Call HandleCuentaDeposit(UserIndex)
        Case ClientPacketID.eCreateEvent
            Call HandleCreateEvent(UserIndex)
        Case ClientPacketID.eCommerceSendChatMessage
            Call HandleCommerceSendChatMessage(UserIndex)
        Case ClientPacketID.eLogMacroClickHechizo
            Call HandleLogMacroClickHechizo(UserIndex)
        Case ClientPacketID.eAddItemCrafting
            Call HandleAddItemCrafting(UserIndex)
        Case ClientPacketID.eRemoveItemCrafting
            Call HandleRemoveItemCrafting(UserIndex)
        Case ClientPacketID.eAddCatalyst
            Call HandleAddCatalyst(UserIndex)
        Case ClientPacketID.eRemoveCatalyst
            Call HandleRemoveCatalyst(UserIndex)
        Case ClientPacketID.eCraftItem
            Call HandleCraftItem(UserIndex)
        Case ClientPacketID.eCloseCrafting
            Call HandleCloseCrafting(UserIndex)
        Case ClientPacketID.eMoveCraftItem
            Call HandleMoveCraftItem(UserIndex)
        Case ClientPacketID.ePetLeaveAll
            Call HandlePetLeaveAll(UserIndex)
        Case ClientPacketID.eResetChar
            Call HandleResetChar(UserIndex)
        Case ClientPacketID.eResetearPersonaje
            Call HandleResetearPersonaje(UserIndex)
        Case ClientPacketID.eDeleteItem
            Call HandleDeleteItem(UserIndex)
        Case ClientPacketID.eFinalizarPescaEspecial
            Call HandleFinalizarPescaEspecial(UserIndex)
        Case ClientPacketID.eRomperCania
            Call HandleRomperCania(UserIndex)
        Case ClientPacketID.eRepeatMacro
            Call HandleRepeatMacro(UserIndex)
        Case ClientPacketID.eBuyShopItem
            Call HandleBuyShopItem(UserIndex)
        Case ClientPacketID.ePublicarPersonajeMAO
            Call HandlePublicarPersonajeMAO(UserIndex)
        Case ClientPacketID.eEventoFaccionario
            Call HandleEventoFaccionario(UserIndex)
        Case ClientPacketID.eRequestDebug
            Call HandleDebugRequest(UserIndex)
        Case ClientPacketID.eLobbyCommand
            Call HandleLobbyCommand(UserIndex)
        Case ClientPacketID.eFeatureToggle
            Call HandleFeatureToggle(UserIndex)
        Case ClientPacketID.eActionOnGroupFrame
            Call HandleActionOnGroupFrame(UserIndex)
        Case ClientPacketID.eSetHotkeySlot
            Call HandleSetHotkeySlot(UserIndex)
        Case ClientPacketID.eUseHKeySlot
            Call HandleUseHKeySlot(UserIndex)
        Case ClientPacketID.eAntiCheatMessage
            Call HandleAntiCheatMessage(UserIndex)
        Case ClientPacketID.eFactionMessage
            Call HandleFactionMessage(UserIndex)
            #If PYMMO = 0 Then
            Case ClientPacketID.eCreateAccount
                Call HandleCreateAccount(ConnectionID)
            Case ClientPacketID.eLoginAccount
                Call HandleLoginAccount(ConnectionID)
            Case ClientPacketID.eDeleteCharacter
                Call HandleDeleteCharacter(ConnectionID)
            #End If
        Case Else
            Call TraceError(&HDEAD0001, "Invalid or unhandled message ID: " & PacketId, "Protocol.HandleIncomingData", Erl)
            If Not IsMissing(optional_user_index) Then
                Call SendData(SendTarget.ToGM, UserIndex, PrepareMessageConsoleMsg("[Error] Paquete desconocido: " & PacketId, e_FontTypeNames.FONTTYPE_GUILD))
            End If
            Call KickConnection(ConnectionID)
            HandleIncomingData = False
            Exit Function
    End Select
    If (reader.GetAvailable() > 0) Then
        Dim errMsg As String
        errMsg = "Server message ID: " & PacketId & " has too many bytes; " & reader.GetAvailable() & " extra bytes found"
        If Not IsMissing(optional_user_index) Then
            errMsg = errMsg & " from user: " & UserList(UserIndex).name
        End If
        Call TraceError(&HDEADBEEF, errMsg, "Protocol.HandleIncomingData", Erl)
        If Not IsMissing(optional_user_index) Then
            Call SendData(SendTarget.ToGM, UserIndex, PrepareMessageConsoleMsg("[Warning] " & errMsg, e_FontTypeNames.FONTTYPE_GUILD))
        End If
        Call KickConnection(ConnectionID)
        HandleIncomingData = False
        Exit Function
    End If
    Call PerformTimeLimitCheck(performance_timer, "Protocol handling message " & PacketID_to_string(PacketId), 100)
    HandleIncomingData = True
End Function

#If PYMMO = 0 Then
    Private Sub HandleCreateAccount(ByVal ConnectionID As Long)
        On Error GoTo HandleCreateAccount_Err:
        Dim username As String
        Dim Password As String
        username = reader.ReadString8
        Password = reader.ReadString8
        Dim UserIndex As Integer
        UserIndex = MapConnectionToUser(ConnectionID)
        If UserIndex < 1 Then
            Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(2094)) ', "No hay slot disponibles para el usuario."))
            Call KickConnection(ConnectionID)
            Exit Sub
        End If
        If (username = "" Or Password = "" Or LenB(Password) <= 3) Then
            Call WriteErrorMsg(UserIndex, "Parametros incorrectos")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        Dim Result As ADODB.Recordset
        Set Result = Query("INSERT INTO account (email, password, salt, validate_code) VALUES (?,?,?,?)", LCase(username), Password, Password, "123")
        If (Result Is Nothing) Then
            Call WriteErrorMsg(UserIndex, "Ya hay una cuenta asociada con ese email")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        Set Result = Query("SELECT id FROM account WHERE email=?", username)
        UserList(UserIndex).AccountID = Result!Id
        Dim Personajes() As t_PersonajeCuenta
        Call WriteAccountCharacterList(UserIndex, Personajes, 0)
        Exit Sub
HandleCreateAccount_Err:
        Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateAccount", Erl)
    End Sub

Private Sub HandleLoginAccount(ByVal ConnectionID As Long)
    On Error GoTo LoginAccount_Err:
    Dim username As String
    Dim Password As String
    username = reader.ReadString8
    Password = reader.ReadString8
    Dim UserIndex As Integer
    UserIndex = MapConnectionToUser(ConnectionID)
    If UserIndex < 1 Then
        Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(2094)) ', "No hay slot disponibles para el usuario."))
        Call KickConnection(ConnectionID)
        Exit Sub
    End If
    If (username = "" Or Password = "" Or LenB(Password) <= 3) Then
        Call WriteErrorMsg(UserIndex, "Parametros incorrectos")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    Dim Result As ADODB.Recordset
    Set Result = Query("SELECT * FROM account WHERE UPPER(email)=UPPER(?) AND password=?", username, Password)
    If (Result.EOF) Then
        Call WriteErrorMsg(UserIndex, "Usuario o Contraseña erronea.")
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    UserList(UserIndex).AccountID = Result!Id
    Dim Personajes(1 To 10) As t_PersonajeCuenta
    Dim count               As Long
    count = GetPersonajesCuentaDatabase(Result!Id, Personajes)
    Call WriteAccountCharacterList(UserIndex, Personajes, count)
    Exit Sub
LoginAccount_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleLoginAccount", Erl)
End Sub

Private Sub HandleDeleteCharacter(ByVal ConnectionID As Long)
    On Error GoTo DeleteCharacter_Err:
DeleteCharacter_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleDeleteCharacter", Erl)
End Sub

Private Sub HandleLoginExistingChar(ByVal ConnectionID As Long)
    On Error GoTo ErrHandler
    Dim user_name As String
    Dim UserIndex As Integer
    UserIndex = Mapping(ConnectionID).UserRef.ArrayIndex
    user_name = reader.ReadString8
    Call ConnectUser(UserIndex, user_name)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleLoginExistingChar", Erl)
End Sub

#End If
#If PYMMO = 1 Then


Private Sub HandleLoginExistingChar(ByVal ConnectionID As Long)
        On Error GoTo ErrHandler

        Dim user_name    As String
        Dim CuentaEmail As String
        Dim Version     As String
        Dim md5         As String
        Dim encrypted_session_token As String
        Dim encrypted_username As String
        
        encrypted_session_token = reader.ReadString8
        encrypted_username = reader.ReadString8
        Version = CStr(reader.ReadInt8()) & "." & CStr(reader.ReadInt8()) & "." & CStr(reader.ReadInt8())
        md5 = reader.ReadString8()

        If Len(encrypted_session_token) <> 88 Then
            Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(2092)) ', "Cliente inválido, por favor realice una actualización."))
            Call KickConnection(ConnectionID)
            Exit Sub
        End If
                
        
        Dim encrypted_session_token_byte() As Byte
        Call AO20CryptoSysWrapper.Str2ByteArr(encrypted_session_token, encrypted_session_token_byte)
        
        Dim decrypted_session_token As String
        decrypted_session_token = AO20CryptoSysWrapper.DECRYPT(PrivateKey, cnvStringFromHexStr(cnvToHex(encrypted_session_token_byte)))
                
        If Not IsBase64(decrypted_session_token) Then
            Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(2092)) ', "Cliente inválido, por favor realice una actualización"))
            Call KickConnection(ConnectionID)
            Exit Sub
        End If
        
        ' Para recibir el ID del user
        Dim RS As ADODB.Recordset
        Set RS = Query("select * from tokens where decrypted_token = '" & decrypted_session_token & "'")
                
        If RS Is Nothing Or RS.RecordCount = 0 Then
            Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(2093)) ', "Sesión inválida, conéctese nuevamente."))
            Call KickConnection(ConnectionID)
            Exit Sub
        End If
        
        CuentaEmail = CStr(RS!username)
                    
        If RS!encrypted_token <> encrypted_session_token Then
            Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(2092)) ', "Cliente inválido, por favor realice una actualización."))
            Call KickConnection(ConnectionID)
            Exit Sub
        End If
        Dim UserIndex As Integer
        UserIndex = MapConnectionToUser(ConnectionID)
        If UserIndex < 1 Then
            Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(2094)) ', "No hay slot disponibles para el usuario."))
            Call KickConnection(ConnectionID)
            Exit Sub
        End If
        
        UserList(UserIndex).encrypted_session_token_db_id = RS!Id
        UserList(UserIndex).encrypted_session_token = encrypted_session_token
        UserList(UserIndex).decrypted_session_token = decrypted_session_token
        UserList(UserIndex).public_key = mid$(decrypted_session_token, 1, 16)
        
        user_name = AO20CryptoSysWrapper.DECRYPT(cnvHexStrFromString(UserList(UserIndex).public_key), encrypted_username)
         
        If Not EntrarCuenta(UserIndex, CuentaEmail, md5) Then
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        Call ConnectUser(UserIndex, user_name, False)
        Exit Sub
    
ErrHandler:
        Call TraceError(Err.Number, Err.Description, "Protocol.HandleLoginExistingChar", Erl)

End Sub

Private Sub HandleLoginNewChar(ByVal ConnectionID As Long)
    On Error GoTo ErrHandler
    Dim username                As String
    Dim CuentaEmail             As String
    Dim Version                 As String
    Dim md5                     As String
    Dim encrypted_session_token As String
    Dim encrypted_username      As String
    Dim race                    As e_Raza
    Dim gender                  As e_Genero
    Dim Hogar                   As e_Ciudad
    Dim Class                   As e_Class
    Dim head                    As Integer

    ' --- read payload exactly as today ---
    encrypted_session_token = reader.ReadString8
    encrypted_username = reader.ReadString8
    Version = CStr(reader.ReadInt8()) & "." & CStr(reader.ReadInt8()) & "." & CStr(reader.ReadInt8())
    md5 = reader.ReadString8()
    race = reader.ReadInt8()
    gender = reader.ReadInt8()
    Class = reader.ReadInt8()
    head = reader.ReadInt16()
    Hogar = reader.ReadInt8()

#If LOGIN_STRESS_TEST = 1 Then
    ' ====== STRESS PATH ======
    If Stress_Enabled() And Stress_IsToken(encrypted_session_token) Then
        Dim UI As Integer
        UI = MapConnectionToUser(ConnectionID)
        If UI < 1 Then
            Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(MSG_CONNECTION_SLOT_ERROR))
            Call KickConnection(ConnectionID)
            Exit Sub
        End If

        ' Optional IP gate (lock to localhost while testing)
        'If Not Stress_IpOk(UserList(ui).ConnectionDetails.IP) Then
        '    Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(MSG_CONNECTION_SLOT_ERROR))
        '    Call KickConnection(ConnectionID)
        '    Exit Sub
        'End If

        ' Build an ephemeral name and mark as stress (no DB, no account checks)
        Dim sName As String
        sName = Stress_NextName()

        UserList(UI).AccountID = -9999               ' sentinel for ephemerals
        UserList(UI).encrypted_session_token = encrypted_session_token
        UserList(UI).decrypted_session_token = "STRESS"
        UserList(UI).public_key = String$(16, "S")   ' harmless filler

        ' Go straight to character creation/spawn using the provided appearance
        If Not ConnectNewUser(UI, sName, race, gender, Class, head, Hogar) Then
            Call CloseSocket(UI)
            Exit Sub
        End If

        Exit Sub  ' IMPORTANT: do not fall through to normal auth
    End If
#End If

    ' ====== NORMAL PATH (unchanged) ======
    If Len(encrypted_session_token) <> 88 Then
        Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(MSG_CLIENT_UPDATE_REQUIRED))
        Exit Sub
    End If
        
    Dim encrypted_session_token_byte() As Byte
    Call AO20CryptoSysWrapper.Str2ByteArr(encrypted_session_token, encrypted_session_token_byte)

    Dim decrypted_session_token As String
    decrypted_session_token = AO20CryptoSysWrapper.DECRYPT(PrivateKey, cnvStringFromHexStr(cnvToHex(encrypted_session_token_byte)))
    If Not IsBase64(decrypted_session_token) Then
        Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(MSG_CLIENT_UPDATE_REQUIRED))
        Call KickConnection(ConnectionID)
        Exit Sub
    End If

    ' Para recibir el ID del user
    Dim RS As ADODB.Recordset
    Set RS = Query("select * from tokens where decrypted_token = '" & decrypted_session_token & "'")
    If RS Is Nothing Or RS.RecordCount = 0 Then
        Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(MSG_INVALID_SESSION_TOKEN))
        Call KickConnection(ConnectionID)
        Exit Sub
    End If

    CuentaEmail = CStr(RS!username)
    If RS!encrypted_token <> encrypted_session_token Then
        Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(MSG_CLIENT_UPDATE_REQUIRED))
        Call KickConnection(ConnectionID)
        Exit Sub
    End If

    Dim UserIndex As Integer
    UserIndex = MapConnectionToUser(ConnectionID)
    If UserIndex < 1 Then
        Call modSendData.SendToConnection(ConnectionID, PrepareShowMessageBox(MSG_CONNECTION_SLOT_ERROR))
        Call KickConnection(ConnectionID)
        Exit Sub
    End If
    UserList(UserIndex).encrypted_session_token_db_id = RS!Id
    UserList(UserIndex).encrypted_session_token = encrypted_session_token
    UserList(UserIndex).decrypted_session_token = decrypted_session_token
    UserList(UserIndex).public_key = mid$(decrypted_session_token, 1, 16)

    username = AO20CryptoSysWrapper.DECRYPT(cnvHexStrFromString(UserList(UserIndex).public_key), encrypted_username)

    If PuedeCrearPersonajes = 0 Then
        Call WriteShowMessageBox(UserIndex, MSG_DISABLED_NEW_CHARACTERS, vbNullString)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If

    If aClon.MaxPersonajes(UserList(UserIndex).ConnectionDetails.IP) Then
        Call WriteShowMessageBox(UserIndex, MSG_YOU_HAVE_TOO_MANY_CHARS, vbNullString)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If

    If EsGmChar(username) Then
        If AdministratorAccounts(UCase$(username)) <> UCase$(CuentaEmail) Then
            Call WriteShowMessageBox(UserIndex, MSG_USERNAME_ALREADY_TAKEN, vbNullString)
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
    End If
    UserList(UserIndex).AccountID = -1
    If Not EntrarCuenta(UserIndex, CuentaEmail, md5) Then
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    Debug.Assert UserList(UserIndex).AccountID > -1
    Dim num_pc As Byte
    num_pc = GetPersonajesCountByIDDatabase(UserList(UserIndex).AccountID)
    Debug.Assert num_pc > 0
    Dim user_tier As e_TipoUsuario
    user_tier = GetPatronTierFromAccountID(UserList(UserIndex).AccountID)
    Dim max_pc_for_tier As Byte
    max_pc_for_tier = MaxCharacterForTier(user_tier)
    Debug.Assert max_pc_for_tier > 0

    If num_pc >= Min(max_pc_for_tier, MAX_PERSONAJES) Then
        Call WriteShowMessageBox(UserIndex, MSG_UPGRADE_ACCOUNT_TO_CREATE_MORE_CHARS, vbNullString)
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    If Not ConnectNewUser(UserIndex, username, race, gender, Class, head, Hogar) Then
        Call CloseSocket(UserIndex)
        Exit Sub
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleLoginNewChar", Erl)
End Sub

#ElseIf PYMMO = 0 Then
    

'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler

        Dim name As String
        Dim race     As e_Raza
        Dim gender   As e_Genero
        Dim Hogar    As e_Ciudad
        Dim Class As e_Class
        Dim head        As Integer

        name = reader.ReadString8
110     race = reader.ReadInt()
112     gender = reader.ReadInt()
113     Class = reader.ReadInt()
116     head = reader.ReadInt()
118     Hogar = reader.ReadInt()

126     If PuedeCrearPersonajes = 0 Then
128         Call WriteShowMessageBox(UserIndex, 1780, vbNullString) 'Msg1780=La creación de personajes en este servidor se ha deshabilitado.
130         Call CloseSocket(UserIndex)
            Exit Sub

        End If

132     If aClon.MaxPersonajes(UserList(UserIndex).ConnectionDetails.IP) Then
134         Call WriteShowMessageBox(UserIndex, 1781, vbNullString) 'Msg1781=Has creado demasiados personajes.

136         Call CloseSocket(UserIndex)
            Exit Sub

        End If

        'Check if we reached MAX_PERSONAJES for this account after updateing the UserList(userindex).AccountID in the if above
        If GetPersonajesCountByIDDatabase(UserList(UserIndex).AccountID) >= MAX_PERSONAJES Then
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        If Not ConnectNewUser(UserIndex, name, race, gender, Class, head, Hogar) Then
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        
        
        Exit Sub
    
ErrHandler:
     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLoginNewChar", Erl)
End Sub
#End If

''
' Handles the "Talk" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleTalk(ByVal UserIndex As Integer)
    'Now hidden on boat pirats recover the proper boat body.
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim chat As String
        chat = reader.ReadString8()
        Dim PacketCounter As Long
        PacketCounter = reader.ReadInt32
        Dim Packet_ID As Long
        Packet_ID = PacketNames.Talk
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "Talk", PacketTimerThreshold( _
                Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        '[Consejeros & GMs]
        If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
            Call LogGM(.name, "Dijo: " & chat)
        End If
        If .flags.Silenciado = 1 Then
            Call WriteLocaleMsg(UserIndex, 110, e_FontTypeNames.FONTTYPE_INFO, .flags.MinutosRestantes)
        Else
            If LenB(chat) <> 0 Then
                '  Foto-denuncias - Push message
                Dim i As Long
                For i = 1 To UBound(.flags.ChatHistory) - 1
                    .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                Next
                .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
                If .flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToPCDeadArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.charindex, CHAT_COLOR_DEAD_CHAR))
                Else
                    If Trim(chat) = "" Then
                        .Counters.timeChat = 0
                    Else
                        .Counters.timeChat = 1 + Ceil((3000 + 60 * Len(chat)) / 1000)
                    End If
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.charindex, .flags.ChatColor, , .pos.x, .pos.y))
                End If
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleTalk", Erl)
End Sub

''
' Handles the "Yell" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleYell(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim chat As String
        chat = reader.ReadString8()
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            'Msg77=¡¡Estás muerto!!.
        Else
            '[Consejeros & GMs]
            If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
                Call LogGM(.name, "Grito: " & chat)
            End If
            'I see you....
            If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                If .flags.Navegando = 1 Then
                    'TODO: Revisar con WyroX
                    If .clase = e_Class.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
                        Call EquiparBarco(UserIndex)
                        ' Msg592=¡Has recuperado tu apariencia normal!
                        Call WriteLocaleMsg(UserIndex, 592, e_FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NoCart, NoBackPack)
                        Call RefreshCharStatus(UserIndex)
                    End If
                Else
                    If .flags.invisible = 0 Then
                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                        'Msg1115= ¡Has vuelto a ser visible!
                        Call WriteLocaleMsg(UserIndex, 1115, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
            If .flags.Silenciado = 1 Then
                Call WriteLocaleMsg(UserIndex, 110, e_FontTypeNames.FONTTYPE_VENENO, .flags.MinutosRestantes)
                'Msg1116= Los administradores te han impedido hablar durante los proximos ¬1
                Call WriteLocaleMsg(UserIndex, 1116, e_FontTypeNames.FONTTYPE_INFO, .flags.MinutosRestantes)
            Else
                If LenB(chat) <> 0 Then
                    '  Foto-denuncias - Push message
                    Dim i As Long
                    For i = 1 To UBound(.flags.ChatHistory) - 1
                        .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                    Next
                    .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
                    If Trim(chat) = "" Then
                        .Counters.timeChat = 0
                    Else
                        .Counters.timeChat = 1 + Ceil((3000 + 60 * Len(chat)) / 1000)
                    End If
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.charindex, vbRed, , .pos.x, .pos.y))
                End If
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleYell", Erl)
End Sub

''
' Handles the "Whisper" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleWhisper(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim chat            As String
        Dim targetCharIndex As String
        Dim TargetUser      As t_UserReference
        targetCharIndex = reader.ReadString8()
        chat = reader.ReadString8()
        If CompararPrivilegios(.flags.Privilegios, UserDarPrivilegioLevel(targetCharIndex)) < 0 Then Exit Sub
        TargetUser = NameIndex(targetCharIndex)
        If UserList(UserIndex).flags.Muerto = 1 Then
            'Msg1117= No puedes susurrar estando muerto.
            Call WriteLocaleMsg(UserIndex, 1117, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not IsValidUserRef(TargetUser) Then
            'Msg1118= El usuario esta muy lejos o desconectado.
            Call WriteLocaleMsg(UserIndex, 1118, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If EstaPCarea(UserIndex, TargetUser.ArrayIndex) Then
            If UserList(TargetUser.ArrayIndex).flags.Muerto = 1 Then
                'Msg1119= No puedes susurrar a un muerto.
                Call WriteLocaleMsg(UserIndex, 1119, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If LenB(chat) <> 0 Then
                Dim i As Long
                For i = 1 To UBound(.flags.ChatHistory) - 1
                    .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                Next
                .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
                Call SendData(SendTarget.ToSuperioresArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.charindex, RGB(157, 226, 20), , .pos.x, .pos.y))
                Call SendData(SendTarget.ToIndex, UserIndex, PrepareConsoleCharText(chat, RGB(157, 226, 20), UserList(UserIndex).name, UserList(UserIndex).Faccion.Status, _
                        UserList(UserIndex).flags.Privilegios))
                Call SendData(SendTarget.ToIndex, TargetUser.ArrayIndex, PrepareConsoleCharText(chat, RGB(157, 226, 20), UserList(UserIndex).name, UserList( _
                        UserIndex).Faccion.Status, UserList(UserIndex).flags.Privilegios))
            End If
        Else
            'Msg1120= El usuario esta muy lejos o desconectado.
            Call WriteLocaleMsg(UserIndex, 1120, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleWhisper", Erl)
End Sub

''
' Handles the "Walk" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleWalk(ByVal UserIndex As Integer)
    On Error GoTo HandleWalk_Err
    Dim Heading As e_Heading
    With UserList(UserIndex)
        Heading = reader.ReadInt8()
        Dim PacketCount As Long
        PacketCount = reader.ReadInt32
        If .flags.Muerto = 0 Then
            If .flags.Navegando Then
                Call verifyTimeStamp(PacketCount, .PacketCounters(PacketNames.Sailing), .PacketTimers(PacketNames.Sailing), .MacroIterations(PacketNames.Sailing), UserIndex, _
                        "Sailing", PacketTimerThreshold(PacketNames.Sailing), MacroIterations(PacketNames.Sailing))
            Else
                Call verifyTimeStamp(PacketCount, .PacketCounters(PacketNames.Walk), .PacketTimers(PacketNames.Walk), .MacroIterations(PacketNames.Walk), UserIndex, "Walk", _
                        PacketTimerThreshold(PacketNames.Walk), MacroIterations(PacketNames.Walk))
            End If
        End If
        If .flags.PescandoEspecial Then
            .Stats.NumObj_PezEspecial = 0
            .flags.PescandoEspecial = False
        End If
        If UserMod.CanMove(.flags, .Counters) Then
            If .flags.Comerciando Or .flags.Crafteando <> 0 Then Exit Sub
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                UserList(UserIndex).Char.FX = 0
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.charindex, 0))
            End If
            Dim currentTick As Long
            currentTick = GetTickCountRaw()
            'Prevent SpeedHack (refactored by WyroX)
            If Not EsGM(UserIndex) And .Char.speeding > 0 Then
                Dim ElapsedTimeStep As Double, MinTimeStep As Long, DeltaStep As Single
                ElapsedTimeStep = TicksElapsed(.Counters.LastStep, currentTick)
                MinTimeStep = .Intervals.Caminar / .Char.speeding
                DeltaStep = (MinTimeStep - ElapsedTimeStep) / MinTimeStep
                If DeltaStep > 0 Then
                    .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep
                    If .Counters.SpeedHackCounter > SvrConfig.GetValue("MaximoSpeedHack") Then
                        Call WritePosUpdate(UserIndex)
                        Exit Sub
                    End If
                Else
                    .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep * 5
                    If .Counters.SpeedHackCounter < 0 Then .Counters.SpeedHackCounter = 0
                End If
            End If
            'Move user
            If MoveUserChar(UserIndex, Heading) Then
                ' Save current step for anti-sh
                .Counters.LastStep = currentTick
                Call ResetUserAutomatedActions(UserIndex)
                If UserList(UserIndex).Grupo.EnGrupo Then
                    Call CompartirUbicacion(UserIndex)
                End If
                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False
                    Call WriteRestOK(UserIndex)
                    'Msg1121= Has dejado de descansar.
                    Call WriteLocaleMsg(UserIndex, 1121, e_FontTypeNames.FONTTYPE_INFO)
                    Call WriteLocaleMsg(UserIndex, 178, e_FontTypeNames.FONTTYPE_INFO)
                End If
                Call CancelExit(UserIndex)
                'Esta usando el /HOGAR, no se puede mover
                If .flags.Traveling = 1 Then
                    .flags.Traveling = 0
                    .Counters.goHome = 0
                    'Msg1122= Has cancelado el viaje a casa.
                    Call WriteLocaleMsg(UserIndex, 1122, e_FontTypeNames.FONTTYPE_INFO)
                End If
                ' Si no pudo moverse
            Else
                .Counters.LastStep = 0
                Call WritePosUpdate(UserIndex)
            End If
        Else    'paralized
            If Not .flags.UltimoMensaje = MSG_PARALYZED Then
                .flags.UltimoMensaje = MSG_PARALYZED
                'Msg1123= No podes moverte porque estas paralizado.
                Call WriteLocaleMsg(UserIndex, MSG_PARALYZED, e_FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(UserIndex, 54, e_FontTypeNames.FONTTYPE_INFO)
            End If
            Call WritePosUpdate(UserIndex)
        End If
        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
            If .clase <> e_Class.Thief And .clase <> e_Class.Bandit Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                If .flags.Navegando = 1 Then
                    If .clase = e_Class.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
                        Call EquiparBarco(UserIndex)
                        ' Msg592=¡Has recuperado tu apariencia normal!
                        Call WriteLocaleMsg(UserIndex, 592, e_FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NoCart, NoBackPack)
                        Call RefreshCharStatus(UserIndex)
                    End If
                Else
                    'If not under a spell effect, show char
                    If .flags.invisible = 0 Then
                        'Msg1124= Has vuelto a ser visible.
                        Call WriteLocaleMsg(UserIndex, 1124, e_FontTypeNames.FONTTYPE_INFO)
                        Call WriteLocaleMsg(UserIndex, 307, e_FontTypeNames.FONTTYPE_INFO)
                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                    End If
                End If
            End If
        End If
    End With
    Exit Sub
HandleWalk_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleWalk", Erl)
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
    On Error GoTo HandleRequestPositionUpdate_Err
    Call WritePosUpdate(UserIndex)

    Exit Sub
HandleRequestPositionUpdate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandlRequestPositionUpdate", Erl)
End Sub

''
' Handles the "Attack" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleAttack(ByVal UserIndex As Integer)
    On Error GoTo HandleAttack_Err
    'Se cancela la salida del juego si el user esta saliendo
    With UserList(UserIndex)
        Dim PacketCounter As Long
        PacketCounter = reader.ReadInt32
        Dim Packet_ID As Long
        Packet_ID = PacketNames.Attack
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "Attack", PacketTimerThreshold( _
                Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            'Msg77=¡¡Estás muerto!!.
            Exit Sub
        End If
        'If equiped weapon is ranged, can't attack this way
        If .invent.EquippedWeaponObjIndex > 0 Then
            If ObjData(.invent.EquippedWeaponObjIndex).Proyectil = 1 And ObjData(.invent.EquippedWeaponObjIndex).Municion > 0 Then
                'Msg1125= No podés usar así esta arma.
                Call WriteLocaleMsg(UserIndex, 1125, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If IsItemInCooldown(UserList(UserIndex), .invent.Object(.invent.EquippedWeaponSlot)) Then
                Exit Sub
            End If
        End If
        If .invent.EquippedWorkingToolObjIndex > 0 Then
            ' Msg694=Para atacar debes desequipar la herramienta.
            Call WriteLocaleMsg(UserIndex, 694, e_FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Meditando Then
            UserList(UserIndex).flags.Meditando = False
            UserList(UserIndex).Char.FX = 0
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.charindex, 0))
        End If
        'If exiting, cancel
        Call CancelExit(UserIndex)
        'Attack!
        Call UsuarioAtaca(UserIndex)
    End With
    Exit Sub
HandleAttack_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleAttack", Erl)
End Sub

''
' Handles the "PickUp" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandlePickUp(ByVal UserIndex As Integer)
    On Error GoTo HandlePickUp_Err
    With UserList(UserIndex)
        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Lower rank administrators can't pick up items
        If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
            ' Msg695=No podés tomar ningun objeto.
            Call WriteLocaleMsg(UserIndex, 695, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call PickObj(UserIndex)
    End With
    Exit Sub
HandlePickUp_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandlePickUp", Erl)
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleSafeToggle(ByVal UserIndex As Integer)
    On Error GoTo HandleSafeToggle_Err
    With UserList(UserIndex)
        Dim cambiaSeguro As Boolean
        cambiaSeguro = False
        If .GuildIndex > 0 And (GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Or GuildAlignmentIndex(.GuildIndex) = _
                e_ALINEACION_GUILD.ALINEACION_ARMADA) Then
            cambiaSeguro = False
        Else
            cambiaSeguro = True
        End If
        If cambiaSeguro Or .flags.Seguro = 0 Then
            If esCiudadano(UserIndex) Then
                If .flags.Seguro Then
                    Call WriteSafeModeOff(UserIndex)
                Else
                    Call WriteSafeModeOn(UserIndex)
                End If
                .flags.Seguro = Not .flags.Seguro
            Else
                ' Msg696=Solo los ciudadanos pueden cambiar el seguro.
                Call WriteLocaleMsg(UserIndex, 696, e_FontTypeNames.FONTTYPE_TALK)
            End If
        Else
            ' Msg697=Debes abandonar el clan para poder sacar el seguro.
            Call WriteLocaleMsg(UserIndex, 697, e_FontTypeNames.FONTTYPE_TALK)
        End If
    End With
    Exit Sub
HandleSafeToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSafeToggle", Erl)
End Sub

' Handles the "PartySafeToggle" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandlePartyToggle(ByVal UserIndex As Integer)
    On Error GoTo HandlePartyToggle_Err
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
    Call TraceError(Err.Number, Err.Description, "Protocol.HandlePartyToggle", Erl)
End Sub

Private Sub HandleSeguroClan(ByVal UserIndex As Integer)
    On Error GoTo HandleSeguroClan_Err
    With UserList(UserIndex)
        .flags.SeguroClan = Not .flags.SeguroClan
        Call WriteClanSeguro(UserIndex, .flags.SeguroClan)
    End With
    Exit Sub
HandleSeguroClan_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSeguroClan", Erl)
End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleRequestGuildLeaderInfo(ByVal UserIndex As Integer)
    On Error GoTo HandleRequestGuildLeaderInfo_Err
    Call modGuilds.SendGuildLeaderInfo(UserIndex)
    Exit Sub
HandleRequestGuildLeaderInfo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestGuildLeaderInfo", Erl)
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleRequestAtributes(ByVal UserIndex As Integer)
    On Error GoTo HandleRequestAtributes_Err
    Call WriteAttributes(UserIndex)
    Exit Sub
HandleRequestAtributes_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestAtributes", Erl)
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
    On Error GoTo HandleRequestSkills_Err
    Call WriteSendSkills(UserIndex)
    Exit Sub
HandleRequestSkills_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestSkills", Erl)
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)
    On Error GoTo HandleRequestMiniStats_Err
    Call WriteMiniStats(UserIndex)
    Exit Sub
HandleRequestMiniStats_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestMiniStats", Erl)
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
    On Error GoTo HandleCommerceEnd_Err
    'User quits commerce mode
    If IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
        If NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundClose <> 0 Then
            Call WritePlayWave(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundClose, NO_3D_SOUND, NO_3D_SOUND, , 1)
        End If
    End If
    UserList(UserIndex).flags.Comerciando = False
    Call WriteCommerceEnd(UserIndex)
    Exit Sub
HandleCommerceEnd_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceEnd", Erl)
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)
    On Error GoTo HandleUserCommerceEnd_Err
    With UserList(UserIndex)
        'Quits commerce mode with user
        If IsValidUserRef(.ComUsu.DestUsu) Then
            If UserList(.ComUsu.DestUsu.ArrayIndex).ComUsu.DestUsu.ArrayIndex = UserIndex Then
                Call WriteConsoleMsg(.ComUsu.DestUsu.ArrayIndex, PrepareMessageLocaleMsg(1949, .name, e_FontTypeNames.FONTTYPE_TALK)) ' Msg1949=¬1 ha dejado de comerciar con vos.
                Call FinComerciarUsu(.ComUsu.DestUsu.ArrayIndex)
                'Send data in the outgoing buffer of the other user
            End If
        End If
        Call FinComerciarUsu(UserIndex)
    End With
    Exit Sub
HandleUserCommerceEnd_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleUserCommerceEnd", Erl)
End Sub

''
' Handles the "BankEnd" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleBankEnd(ByVal UserIndex As Integer)
    On Error GoTo HandleBankEnd_Err
    With UserList(UserIndex)
        If .flags.Comerciando Then
            .flags.Comerciando = False
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave("171", UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
            Call WriteBankEnd(UserIndex)
        End If
    End With
    Exit Sub
HandleBankEnd_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankEnd", Erl)
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
    On Error GoTo HandleUserCommerceOk_Err
    'Trade accepted
    Call AceptarComercioUsu(UserIndex)
    Exit Sub
HandleUserCommerceOk_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleUserCommerceOk", Erl)
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)
    On Error GoTo HandleUserCommerceReject_Err
    Dim otherUser As Integer
    With UserList(UserIndex)
        otherUser = .ComUsu.DestUsu.ArrayIndex
        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, PrepareMessageLocaleMsg(1950, .name, e_FontTypeNames.FONTTYPE_TALK)) ' Msg1950=¬1 ha rechazado tu oferta.
                Call FinComerciarUsu(otherUser)
                'Send data in the outgoing buffer of the other user
            End If
        End If
        ' Msg698=Has rechazado la oferta del otro usuario.
        Call WriteLocaleMsg(UserIndex, 698, e_FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)
    End With
    Exit Sub
HandleUserCommerceReject_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleUserCommerceReject", Erl)
End Sub

''
' Handles the "Drop" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleDrop(ByVal UserIndex As Integer)
    On Error GoTo HandleDrop_Err
    'Agregue un checkeo para patear a los usuarios que tiran items mientras comercian.
    Dim Slot          As Byte
    Dim amount        As Long
    Dim PacketCounter As Long
    Dim Packet_ID     As Long
    With UserList(UserIndex)
        Slot = reader.ReadInt8()
        amount = reader.ReadInt32()
        PacketCounter = reader.ReadInt32
        Packet_ID = PacketNames.Drop
        If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then
            If Slot <> GOLD_SLOT Then
                Exit Sub
            End If
        End If
        If IsInMapCarcelRestrictedArea(.pos) Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MSG_CANNOT_PICK_UP_ITEMS_IN_JAIL, vbNullString, e_FontTypeNames.FONTTYPE_INFO))
            Exit Sub
        End If
        If Not IntervaloPermiteTirar(UserIndex) Then Exit Sub
        If .flags.PescandoEspecial = True Then Exit Sub
        If amount <= 0 Then Exit Sub
        'low rank admins can't drop item. Neither can the dead nor those sailing or riding a horse.
        If .flags.Muerto = 1 Then Exit Sub
        'If the user is trading, he can't drop items => He's cheating, we kick him.
        If .flags.Comerciando Then Exit Sub
        If .flags.Montado = 1 Then
            ' Msg699=Debes descender de tu montura para dejar objetos en el suelo.
            Call WriteLocaleMsg(UserIndex, 699, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        'Are we dropping gold or other items??
        If Slot = FLAGORO Then
            If amount > 100000 Then amount = 100000
            Call TirarOro(amount, UserIndex)
        Else
            If Slot <= getMaxInventorySlots(UserIndex) Then
                '04-05-08 Ladder
                If (.flags.Privilegios And e_PlayerType.Admin) <> 16 Then
                    If EsNewbie(UserIndex) And ObjData(.invent.Object(Slot).ObjIndex).Newbie = 1 Then
                        ' Msg701=No se pueden tirar los objetos Newbies.
                        Call WriteLocaleMsg(UserIndex, 701, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    If ObjData(.invent.Object(Slot).ObjIndex).Intirable = 1 And Not EsGM(UserIndex) Then
                        ' Msg702=Acción no permitida.
                        Call WriteLocaleMsg(UserIndex, 702, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    ElseIf ObjData(.invent.Object(Slot).ObjIndex).Intirable = 1 And EsGM(UserIndex) Then
                        If Slot <= UserList(UserIndex).CurrentInventorySlots And Slot > 0 Then
                            If .invent.Object(Slot).ObjIndex = 0 Then Exit Sub
                            Call DropObj(UserIndex, Slot, amount, .pos.Map, .pos.x, .pos.y)
                        End If
                        Exit Sub
                    End If
                    If ObjData(.invent.Object(Slot).ObjIndex).Instransferible = 1 Then
                        ' Msg702=Acción no permitida.
                        Call WriteLocaleMsg(UserIndex, 702, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                If ObjData(.invent.Object(Slot).ObjIndex).OBJType = e_OBJType.otShips And UserList(UserIndex).flags.Navegando Then
                    ' Msg703=Para tirar la barca deberias estar en tierra firme.
                    Call WriteLocaleMsg(UserIndex, 703, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Else
                'ver de banear al usuario
                'Call BanearIP(0, UserList(UserIndex).name, UserList(UserIndex).IP, UserList(UserIndex).Cuenta)
                Call LogEdicionPaquete("El usuario " & UserList(UserIndex).name & " editó el slot del inventario | Valor: " & Slot & ".")
            End If
            '04-05-08 Ladder
            'Only drop valid slots
            If Slot <= UserList(UserIndex).CurrentInventorySlots And Slot > 0 Then
                If .invent.Object(Slot).ObjIndex = 0 Then Exit Sub
                Call DropObj(UserIndex, Slot, amount, .pos.Map, .pos.x, .pos.y)
            End If
        End If
    End With
    Exit Sub
HandleDrop_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleDrop", Erl)
End Sub

Public Function verifyTimeStamp(ByVal ActualCount As Long, _
                                ByRef LastCount As Long, _
                                ByRef LastTick As Long, _
                                ByRef Iterations, _
                                ByVal UserIndex As Integer, _
                                ByVal PacketName As String, _
                                Optional ByVal DeltaThreshold As Long = 100, _
                                Optional ByVal MaxIterations As Long = 5, _
                                Optional ByVal CloseClient As Boolean = False) As Boolean
    Dim Ticks As Long, Delta As Double
    Ticks = GetTickCountRaw()
    Delta = TicksElapsed(LastTick, Ticks)
    LastTick = Ticks
    'Controlamos secuencia para ver que no haya paquetes duplicados.
    If ActualCount <= LastCount Then
        Call SendData(SendTarget.ToAdminsYDioses, UserIndex, PrepareMessageConsoleMsg("Paquete grabado: " & PacketName & " | Cuenta: " & UserList(UserIndex).Cuenta & " | Ip: " & _
                UserList(UserIndex).ConnectionDetails.IP & " (Baneado automaticamente)", e_FontTypeNames.FONTTYPE_INFOBOLD))
        Call LogEdicionPaquete("El usuario " & UserList(UserIndex).name & " editó el paquete " & PacketName & ".")
        Call SendData(SendTarget.ToAdminsYDioses, UserIndex, PrepareMessageConsoleMsg("Paquete grabado: " & PacketName & " | Cuenta: " & UserList(UserIndex).Cuenta & " | Ip: " & _
                UserList(UserIndex).ConnectionDetails.IP & " (Baneado automaticamente)", e_FontTypeNames.FONTTYPE_INFOBOLD))
        Call LogEdicionPaquete("El usuario " & UserList(UserIndex).name & " editó el paquete " & PacketName & ".")
        LastCount = ActualCount
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    'controlamos speedhack/macro
    If Delta < DeltaThreshold Then
        Iterations = Iterations + 1
        If Iterations >= MaxIterations Then
            'Call WriteShowMessageBox(UserIndex, "Relajate andá a tomarte un té con Gulfas.")
            verifyTimeStamp = False
            'Call LogMacroServidor("El usuario " & UserList(UserIndex).name & " iteró el paquete " & PacketName & " " & MaxIterations & " veces.")
            Call SendData(SendTarget.ToAdminsYDioses, UserIndex, PrepareMessageConsoleMsg("Control de macro---> El usuario " & UserList(UserIndex).name & "| Revisar --> " & _
                    PacketName & " (Envíos: " & Iterations & ").", e_FontTypeNames.FONTTYPE_INFOBOLD))
            'Call WriteCerrarleCliente(UserIndex)
            'Call CloseSocket(UserIndex)
            LastCount = ActualCount
            Iterations = 0
            Debug.Print "CIERRO CLIENTE"
        End If
        'Exit Function
    Else
        Iterations = 0
    End If
    verifyTimeStamp = True
    LastCount = ActualCount
End Function

''
' Handles the "CastSpell" message.
' @param    UserIndex The index of the user sending the message.
Private Sub HandleCastSpell(ByVal UserIndex As Integer)
    On Error GoTo HandleCastSpell_Err
    Dim Spell As Byte
    Spell = reader.ReadInt8()
    Dim PacketCounter As Long
    PacketCounter = reader.ReadInt32
    Dim Packet_ID As Long
    Packet_ID = PacketNames.CastSpell
    Call UseSpellSlot(UserIndex, Spell)
    Exit Sub
HandleCastSpell_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCastSpell", Erl)
End Sub

''
' Handles the "LeftClick" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleLeftClick(ByVal UserIndex As Integer)
    On Error GoTo HandleLeftClick_Err
    With UserList(UserIndex)
        Dim x As Byte
        Dim y As Byte
        x = reader.ReadInt8()
        y = reader.ReadInt8()
        Dim PacketCounter As Long
        PacketCounter = reader.ReadInt32
        Dim Packet_ID As Long
        Packet_ID = PacketNames.LeftClick
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "LeftClick", PacketTimerThreshold( _
                Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        Call LookatTile(UserIndex, .pos.Map, x, y)
    End With
    Exit Sub
HandleLeftClick_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleLeftClick", Erl)
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleDoubleClick(ByVal UserIndex As Integer)
    On Error GoTo HandleDoubleClick_Err
    With UserList(UserIndex)
        Dim x As Byte
        Dim y As Byte
        x = reader.ReadInt8()
        y = reader.ReadInt8()
        Call Accion(UserIndex, .pos.Map, x, y)
    End With
    Exit Sub
HandleDoubleClick_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleDoubleClick", Erl)
End Sub

Private Sub HandleWork(ByVal UserIndex As Integer)
    On Error GoTo HandleWork_Err
    With UserList(UserIndex)
        Dim Skill As e_Skill
        Skill = reader.ReadInt8()
        Dim PacketCounter As Long
        PacketCounter = reader.ReadInt32
        If UserList(UserIndex).flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'If exiting, cancel
        Call CancelExit(UserIndex)
        Select Case Skill
            Case Robar, Magia, Domar
                Call WriteWorkRequestTarget(UserIndex, Skill)
            Case Ocultarse
                If Not verifyTimeStamp(PacketCounter, .PacketCounters(PacketNames.Hide), .PacketTimers(PacketNames.Hide), .MacroIterations(PacketNames.Hide), UserIndex, _
                        "Ocultar", PacketTimerThreshold(PacketNames.Hide), MacroIterations(PacketNames.Hide)) Then Exit Sub
                If .flags.Montado = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = MSG_CANNOT_HIDE_MOUNTED Then
                        ' Msg704=No podés ocultarte si estás montado.
                        Call WriteLocaleMsg(UserIndex, MSG_CANNOT_HIDE_MOUNTED, e_FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = MSG_CANNOT_HIDE_MOUNTED
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                If .flags.Oculto = 1 Then
                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = MSG_ALREADY_HIDDEN Then
                        Call WriteLocaleMsg(UserIndex, 55, e_FontTypeNames.FONTTYPE_INFO)
                        'Msg1127= Ya estás oculto.
                        Call WriteLocaleMsg(UserIndex, MSG_ALREADY_HIDDEN, e_FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = MSG_ALREADY_HIDDEN
                    End If
                    '[/CDT]
                    Exit Sub
                End If
                If .flags.EnReto Then
                    ' Msg705=No podés ocultarte durante un reto.
                    Call WriteLocaleMsg(UserIndex, 705, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .flags.EnConsulta Then
                    ' Msg706=No podés ocultarte si estas en consulta.
                    Call WriteLocaleMsg(UserIndex, 706, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .flags.invisible Then
                    ' Msg707=No podés ocultarte si estás invisible.
                    Call WriteLocaleMsg(UserIndex, 707, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If MapInfo(.pos.Map).SinInviOcul Then
                    ' Msg708=Una fuerza divina te impide ocultarte en esta zona.
                    Call WriteLocaleMsg(UserIndex, 708, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call DoOcultarse(UserIndex)
        End Select
    End With
    Exit Sub
HandleWork_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleWork", Erl)
End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleUseSpellMacro(ByVal UserIndex As Integer)
    On Error GoTo HandleUseSpellMacro_Err
    With UserList(UserIndex)
        Call SendData(SendTarget.ToAdminsYDioses, UserIndex, PrepareMessageConsoleMsg(.name & " fue expulsado por Anti-macro de hechizos", e_FontTypeNames.FONTTYPE_VENENO))
        Call WriteShowMessageBox(UserIndex, 1782, vbNullString) 'Msg1782=Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros.
        Call CloseSocket(UserIndex)
    End With
    Exit Sub
HandleUseSpellMacro_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleUseSpellMacro", Erl)
End Sub

''
' Handles the "UseItem" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleUseItem(ByVal UserIndex As Integer)
    On Error GoTo HandleUseItem_Err
    With UserList(UserIndex)
        Dim Slot As Byte
        Slot = reader.ReadInt8()
        Dim DesdeInventario As Boolean
        DesdeInventario = reader.ReadInt8
        If Not DesdeInventario Then
            Call SendData(SendTarget.ToAdminsYDioses, UserIndex, PrepareMessageConsoleMsg("El usuario " & .name & _
                    " está tomando pociones con click estando en hechizos....Fue kickeado automaticamente", e_FontTypeNames.FONTTYPE_INFOBOLD))
            Call modNetwork.Kick(UserList(UserIndex).ConnectionDetails.ConnID)
        End If
        Dim PacketCounter As Long
        PacketCounter = reader.ReadInt32
        Dim Packet_ID As Long
        Packet_ID = PacketNames.UseItem
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "UseItem", PacketTimerThreshold( _
                Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        '  Debug.Print "LLEGA PAQUETE"
        If Slot <= UserList(UserIndex).CurrentInventorySlots And Slot > 0 Then
            If .invent.Object(Slot).ObjIndex = 0 Then Exit Sub
            Call UseInvItem(UserIndex, Slot, 1)
        End If
    End With
    Exit Sub
HandleUseItem_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleUseItem", Erl)
End Sub

''
' Handles the "UseItem" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleUseItemU(ByVal UserIndex As Integer)
    On Error GoTo HandleUseItemU_Err
    With UserList(UserIndex)
        Dim Slot As Byte
        Slot = reader.ReadInt8()
        Dim PacketCounter As Long
        PacketCounter = reader.ReadInt32
        Dim Packet_ID As Long
        Packet_ID = PacketNames.UseItemU
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "UseItemU", PacketTimerThreshold( _
                Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        If Slot <= UserList(UserIndex).CurrentInventorySlots And Slot > 0 Then
            If .invent.Object(Slot).ObjIndex = 0 Then Exit Sub
            Call UseInvItem(UserIndex, Slot, 0)
        End If
    End With
    Exit Sub
HandleUseItemU_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleUseItemU", Erl)
End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)
    On Error GoTo HandleCraftBlacksmith_Err
    Dim Item As Integer
    Item = reader.ReadInt16()
    If Item < 1 Then Exit Sub
    ' If ObjData(Item).SkHerreria = 0 Then Exit Sub
    Call HerreroConstruirItem(UserIndex, Item)
    Exit Sub
HandleCraftBlacksmith_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftBlacksmith", Erl)
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)
    On Error GoTo HandleCraftCarpenter_Err
    Dim Item As Integer
    Item = reader.ReadInt16()
    Dim Cantidad As Long
    Cantidad = reader.ReadInt32()
    If Item = 0 Then Exit Sub
    'Valido que haya puesto una cantidad > 0
    If Cantidad > 0 Then
        UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
        UserList(UserIndex).Trabajo.TargetSkill = e_Skill.Carpinteria
        UserList(UserIndex).Trabajo.Cantidad = Cantidad
        UserList(UserIndex).Trabajo.Item = Item
        Call WriteMacroTrabajoToggle(UserIndex, True)
    Else
    End If
    '106         Call CarpinteroConstruirItem(UserIndex, Item, Cantidad)
    Exit Sub
HandleCraftCarpenter_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftCarpenter", Erl)
End Sub

Private Sub HandleCraftAlquimia(ByVal UserIndex As Integer)
    On Error GoTo HandleCraftAlquimia_Err
    Dim Item As Integer
    Item = reader.ReadInt16()
    If Item < 1 Then Exit Sub
    Call AlquimistaConstruirItem(UserIndex, Item)
    Exit Sub
HandleCraftAlquimia_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftAlquimia", Erl)
End Sub

Private Sub HandleCraftSastre(ByVal UserIndex As Integer)
    On Error GoTo HandleCraftSastre_Err
    Dim Item As Integer
    Item = reader.ReadInt16()
    If Item < 1 Then Exit Sub
    Call SastreConstruirItem(UserIndex, Item)
    Exit Sub
HandleCraftSastre_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftSastre", Erl)
End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)
    On Error GoTo HandleWorkLeftClick_Err
    With UserList(UserIndex)
        Dim x        As Byte
        Dim y        As Byte
        Dim Skill    As e_Skill
        Dim DummyInt As Integer
        Dim tU       As Integer   'Target user
        Dim tN       As Integer   'Target NPC
        x = reader.ReadInt8()
        y = reader.ReadInt8()
        Skill = reader.ReadInt8()
        Dim PacketCounter As Long
        PacketCounter = reader.ReadInt32
        Dim Packet_ID As Long
        Packet_ID = PacketNames.WorkLeftClick
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "WorkLeftClick", PacketTimerThreshold( _
                Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        .Trabajo.Target_X = x
        .Trabajo.Target_Y = y
        .Trabajo.TargetSkill = Skill
        If .flags.Muerto = 1 Or .flags.Descansar Or Not InMapBounds(.pos.Map, x, y) Then Exit Sub
        If UserMod.IsStun(.flags, .Counters) Then Exit Sub
        If Not InRangoVision(UserIndex, x, y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
        End If
        If .flags.Meditando Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
        End If
        'If exiting, cancel
        Call CancelExit(UserIndex)
        Select Case Skill
                Dim consumirMunicion As Boolean
            Case e_Skill.Proyectiles
                Dim WeaponData     As t_ObjData
                Dim ProjectileType As Byte
                'Check attack interval
                If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Sub
                'Check Magic interval
                If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                'Make sure the item is valid and there is ammo equipped.
                With .invent
                    If .EquippedWeaponObjIndex < 1 Then Exit Sub
                    WeaponData = ObjData(.EquippedWeaponObjIndex)
                    If IsItemInCooldown(UserList(UserIndex), .Object(.EquippedWeaponSlot)) Then Exit Sub
                    ProjectileType = GetProjectileView(UserList(UserIndex))
                    If WeaponData.Proyectil = 1 And WeaponData.Municion = 0 Then
                        DummyInt = 0
                    ElseIf .EquippedWeaponObjIndex = 0 Then
                        DummyInt = 1
                    ElseIf .EquippedWeaponSlot < 1 Or .EquippedWeaponSlot > UserList(UserIndex).CurrentInventorySlots Then
                        DummyInt = 1
                    ElseIf .EquippedMunitionSlot < 1 Or .EquippedMunitionSlot > UserList(UserIndex).CurrentInventorySlots Then
                        DummyInt = 1
                    ElseIf .EquippedMunitionObjIndex = 0 Then
                        DummyInt = 1
                    ElseIf ObjData(.EquippedWeaponObjIndex).Proyectil <> 1 Then
                        DummyInt = 2
                    ElseIf ObjData(.EquippedMunitionObjIndex).OBJType <> e_OBJType.otArrows Then
                        DummyInt = 1
                    ElseIf .Object(.EquippedMunitionSlot).amount < 1 Then
                        DummyInt = 1
                    ElseIf ObjData(.EquippedMunitionObjIndex).Subtipo <> WeaponData.Municion Then
                        DummyInt = 1
                    End If
                    If DummyInt <> 0 Then
                        If DummyInt = 1 Then
                            ' Msg709=No tenés municiones.
                            Call WriteLocaleMsg(UserIndex, 709, e_FontTypeNames.FONTTYPE_INFO)
                        End If
                        Call Desequipar(UserIndex, .EquippedMunitionSlot)
                        Call WriteWorkRequestTarget(UserIndex, 0)
                        Exit Sub
                    End If
                End With
                'Quitamos stamina
                If .Stats.MinSta >= 10 Then
                    Call QuitarSta(UserIndex, RandomNumber(1, 10))
                Else
                    'Msg2129=¡No tengo energía!
                    Call SendData(SendTarget.ToIndex, UserIndex, PrepareLocalizedChatOverHead(2129, UserList(UserIndex).Char.charindex, vbWhite))
                    'Msg1128= Estás muy cansado para luchar.
                    Call WriteLocaleMsg(UserIndex, 1128, e_FontTypeNames.FONTTYPE_INFO)
                    Call WriteWorkRequestTarget(UserIndex, 0)
                    Exit Sub
                End If
                Call LookatTile(UserIndex, .pos.Map, x, y)
                tU = .flags.TargetUser.ArrayIndex
                tN = .flags.TargetNPC.ArrayIndex
                consumirMunicion = False
                'Validate target
                If IsValidUserRef(.flags.TargetUser) Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(UserList(tU).pos.y - .pos.y) > RANGO_VISION_Y Then
                        ' Msg8=Estas demasiado lejos.
                        Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(UserIndex, 0)
                        Exit Sub
                    End If
                    'Prevent from hitting self
                    If tU = UserIndex Then
                        ' Msg710=¡No podés atacarte a vos mismo!
                        Call WriteLocaleMsg(UserIndex, 710, e_FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(UserIndex, 0)
                        Exit Sub
                    End If
                    'Attack!
                    If Not PuedeAtacar(UserIndex, tU) Then Exit Sub 'TODO: Por ahora pongo esto para solucionar lo anterior.
                    Dim backup    As Byte
                    Dim envie     As Boolean
                    Dim Particula As Integer
                    Dim Tiempo    As Long
                    If .flags.invisible > 0 Then
                        If IsFeatureEnabled("remove-inv-on-attack") Then
                            Call RemoveUserInvisibility(UserIndex)
                        End If
                    End If
                    Call UsuarioAtacaUsuario(UserIndex, tU, Ranged)
                    Dim FX As Integer
                    If .invent.EquippedMunitionObjIndex Then
                        FX = ObjData(.invent.EquippedMunitionObjIndex).CreaFX
                    End If
                    If FX <> 0 Then
                        UserList(tU).Counters.timeFx = 3
                        Call SendData(SendTarget.ToPCAliveArea, tU, PrepareMessageCreateFX(UserList(tU).Char.charindex, FX, 0, UserList(tU).pos.x, UserList(tU).pos.y))
                    End If
                    If ProjectileType > 0 And (.flags.Oculto = 0 Or Not MapInfo(.pos.Map).KeepInviOnAttack) Then
                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, x, y, ProjectileType))
                    End If
                    'Si no es GM invisible, le envio el movimiento del arma.
                    If UserList(UserIndex).flags.AdminInvisible = 0 Then
                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageArmaMov(UserList(UserIndex).Char.charindex, 1))
                    End If
                    If .invent.EquippedMunitionObjIndex > 0 Then
                        If ObjData(.invent.EquippedMunitionObjIndex).CreaParticula <> "" Then
                            Particula = val(ReadField(1, ObjData(.invent.EquippedMunitionObjIndex).CreaParticula, Asc(":")))
                            Tiempo = val(ReadField(2, ObjData(.invent.EquippedMunitionObjIndex).CreaParticula, Asc(":")))
                            UserList(tU).Counters.timeFx = 3
                            Call SendData(SendTarget.ToPCAliveArea, tU, PrepareMessageParticleFX(UserList(tU).Char.charindex, Particula, Tiempo, False, , UserList(tU).pos.x, _
                                    UserList(tU).pos.y))
                        End If
                    End If
                    consumirMunicion = True
                ElseIf tN > 0 Then
                    'Only allow to atack if the other one can retaliate (can see us)
                    If Abs(NpcList(tN).pos.y - .pos.y) > RANGO_VISION_Y And Abs(NpcList(tN).pos.x - .pos.x) > RANGO_VISION_X Then
                        ' Msg8=Estas demasiado lejos.
                        Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(UserIndex, 0)
                        Exit Sub
                    End If
                    'Is it attackable???
                    If NpcList(tN).Attackable <> 0 Then
                        Dim UserAttackInteractionResult As t_AttackInteractionResult
                        UserAttackInteractionResult = UserCanAttackNpc(UserIndex, tN)
                        Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResult.Result)
                        If UserAttackInteractionResult.CanAttack Then
                            If UserAttackInteractionResult.TurnPK Then Call VolverCriminal(UserIndex)
                            Call UsuarioAtacaNpc(UserIndex, tN, Ranged)
                            consumirMunicion = True
                            If ProjectileType > 0 And .flags.Oculto = 0 Then
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, x, y, _
                                        ProjectileType))
                            End If
                            'Si no es GM invisible, le envio el movimiento del arma.
                            If UserList(UserIndex).flags.AdminInvisible = 0 Then
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageArmaMov(UserList(UserIndex).Char.charindex, 1))
                            End If
                        Else
                            consumirMunicion = False
                        End If
                    End If
                End If
                With .invent
                    If WeaponData.Proyectil = 1 And WeaponData.Municion > 0 Then
                        DummyInt = .EquippedMunitionSlot
                        If ObjData(.EquippedWeaponObjIndex).CreaWav > 0 Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(ObjData(.EquippedWeaponObjIndex).CreaWav, UserList(UserIndex).pos.x, _
                                    UserList(UserIndex).pos.y))
                        End If
                        If DummyInt <> 0 Then
                            'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
                            If consumirMunicion And Not IsConsumableFreeZone(UserIndex) Then
                                Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                            End If
                            If .Object(DummyInt).amount > 0 Then
                                'QuitarUserInvItem unequipps the ammo, so we equip it again
                                .EquippedMunitionSlot = DummyInt
                                .EquippedMunitionObjIndex = .Object(DummyInt).ObjIndex
                                .Object(DummyInt).Equipped = 1
                            Else
                                .EquippedMunitionSlot = 0
                                .EquippedMunitionObjIndex = 0
                            End If
                            Call UpdateUserInv(False, UserIndex, DummyInt)
                        End If
                    ElseIf consumirMunicion Then
                        Call UpdateCd(UserIndex, WeaponData.cdType)
                    End If
                End With
                '-----------------------------------
            Case e_Skill.Magia
                'Target whatever is in that tile
                Call LookatTile(UserIndex, .pos.Map, x, y)
                'If it's outside range log it and exit
                If Abs(.pos.x - x) > RANGO_VISION_X Or Abs(.pos.y - y) > RANGO_VISION_Y Then
                    Call LogSecurity("Ataque fuera de rango de " & .name & "(" & .pos.Map & "/" & .pos.x & "/" & .pos.y & ") ip: " & .ConnectionDetails.IP & " a la posicion (" & _
                            .pos.Map & "/" & x & "/" & y & ")")
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
                    .Counters.controlHechizos.HechizosTotales = .Counters.controlHechizos.HechizosTotales + 1
                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
                    .flags.Hechizo = 0
                Else
                    ' Msg587=¡Primero selecciona el hechizo que quieres lanzar!
                    Call WriteLocaleMsg(UserIndex, 587, e_FontTypeNames.FONTTYPE_INFO)
                End If
            Case e_Skill.Pescar
                If .Counters.Trabajando = 0 And .Counters.LastTrabajo = 0 Then
                    If .invent.EquippedWorkingToolSlot = 0 Then Exit Sub
                    If IsItemInCooldown(UserList(UserIndex), .invent.Object(.invent.EquippedWorkingToolSlot)) Then Exit Sub
                    Call LookatTile(UserIndex, .pos.Map, x, y)
                    Call FishOrThrowNet(UserIndex)
                End If
            Case e_Skill.Talar
                If .Counters.Trabajando = 0 And .Counters.LastTrabajo = 0 Then
                    Call Trabajar(UserIndex, e_Skill.Talar)
                End If
            Case e_Skill.Alquimia
                If .invent.EquippedWorkingToolObjIndex = 0 Then Exit Sub
                If ObjData(.invent.EquippedWorkingToolObjIndex).OBJType <> e_OBJType.otWorkingTools Then Exit Sub
                'Check interval
                If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub
                Select Case ObjData(.invent.EquippedWorkingToolObjIndex).Subtipo
                    Case e_WorkingToolSubType.AlchemyScissors  ' Herramientas de Alquimia - Tijeras
                        If MapInfo(UserList(UserIndex).pos.Map).Seguro = 1 Then
                            Call WriteWorkRequestTarget(UserIndex, 0)
                            ' Msg711=Esta prohibido cortar raices en las ciudades.
                            Call WriteLocaleMsg(UserIndex, 711, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        If MapData(.pos.Map, x, y).ObjInfo.amount <= 0 Then
                            ' Msg712=El árbol ya no te puede entregar mas raices.
                            Call WriteLocaleMsg(UserIndex, 712, e_FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(UserIndex, 0)
                            Call WriteMacroTrabajoToggle(UserIndex, False)
                            Exit Sub
                        End If
                        DummyInt = MapData(.pos.Map, x, y).ObjInfo.ObjIndex
                        If DummyInt > 0 Then
                            If Abs(.pos.x - x) + Abs(.pos.y - y) > 2 Then
                                Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                                'Msg1129= Estas demasiado lejos.
                                Call WriteLocaleMsg(UserIndex, 1129, e_FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
                            End If
                            If .pos.x = x And .pos.y = y Then
                                ' Msg713=No podés quitar raices allí.
                                Call WriteLocaleMsg(UserIndex, 713, e_FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
                            End If
                        Else
                            ' Msg604=No podés quitar raices allí.
                            Call WriteLocaleMsg(UserIndex, 604, e_FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(UserIndex, 0)
                            Call WriteMacroTrabajoToggle(UserIndex, False)
                        End If
                End Select
            Case e_Skill.Mineria
                If .Counters.Trabajando = 0 And .Counters.LastTrabajo = 0 Then
                    Call Trabajar(UserIndex, e_Skill.Mineria)
                End If
            Case e_Skill.Robar
                'Does the map allow us to steal here?
                If MapInfo(.pos.Map).Seguro = 0 Then
                    'Check interval
                    If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub
                    'Target whatever is in that tile
                    Call LookatTile(UserIndex, UserList(UserIndex).pos.Map, x, y)
                    tU = .flags.TargetUser.ArrayIndex
                    If IsValidUserRef(.flags.TargetUser) And tU <> UserIndex Then
                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And e_PlayerType.User Then
                            If UserList(tU).flags.Muerto = 0 Then
                                Dim DistanciaMaxima As Integer
                                If .clase = e_Class.Thief Then
                                    DistanciaMaxima = 1
                                Else
                                    DistanciaMaxima = 1
                                End If
                                If Abs(.pos.x - UserList(tU).pos.x) + Abs(.pos.y - UserList(tU).pos.y) > DistanciaMaxima Then
                                    Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                                    'Msg1130= Estís demasiado lejos.
                                    Call WriteLocaleMsg(UserIndex, 1130, e_FontTypeNames.FONTTYPE_INFO)
                                    Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub
                                End If
                                '17/09/02
                                'Check the trigger
                                If MapData(UserList(tU).pos.Map, UserList(tU).pos.x, UserList(tU).pos.y).trigger = e_Trigger.ZonaSegura Then
                                    ' Msg714=No podés robar aquí.
                                    Call WriteLocaleMsg(UserIndex, 714, e_FontTypeNames.FONTTYPE_WARNING)
                                    Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub
                                End If
                                If MapData(.pos.Map, .pos.x, .pos.y).trigger = e_Trigger.ZonaSegura Then
                                    ' Msg714=No podés robar aquí.
                                    Call WriteLocaleMsg(UserIndex, 714, e_FontTypeNames.FONTTYPE_WARNING)
                                    Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub
                                End If
                                Call DoRobar(UserIndex, tU)
                            End If
                        End If
                    Else
                        ' Msg715=No a quien robarle!
                        Call WriteLocaleMsg(UserIndex, 715, e_FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(UserIndex, 0)
                    End If
                Else
                    ' Msg716=¡No podés robar en zonas seguras!
                    Call WriteLocaleMsg(UserIndex, 716, e_FontTypeNames.FONTTYPE_INFO)
                    Call WriteWorkRequestTarget(UserIndex, 0)
                End If
            Case e_Skill.Domar
                Call LookatTile(UserIndex, .pos.Map, x, y)
                If IsValidNpcRef(.flags.TargetNPC) Then
                    tN = .flags.TargetNPC.ArrayIndex
                    If NpcList(tN).flags.Domable > 0 Then
                        If Abs(.pos.x - x) + Abs(.pos.y - y) > 4 Then
                            ' Msg8=Estas demasiado lejos.
                            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        If GetOwnedBy(tN) <> 0 Then
                            ' Msg717=No puedes domar una criatura que esta luchando con un jugador.
                            Call WriteLocaleMsg(UserIndex, 717, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        Call DoDomar(UserIndex, tN)
                    Else
                        ' Msg718=No puedes domar a esa criatura.
                        Call WriteLocaleMsg(UserIndex, 718, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    ' Msg719=No hay ninguna criatura alli!
                    Call WriteLocaleMsg(UserIndex, 719, e_FontTypeNames.FONTTYPE_INFO)
                End If
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
                'Check interval
                If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub
                Call LookatTile(UserIndex, .pos.Map, x, y)
                'Check there is a proper item there
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = e_OBJType.otForge Then
                        'Validate other items
                        If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > UserList(UserIndex).CurrentInventorySlots Then
                            Exit Sub
                        End If
                        ''chequeamos que no se zarpe duplicando oro
                        If .invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                            If .invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .invent.Object(.flags.TargetObjInvSlot).amount = 0 Then
                                ' Msg605=No tienes más minerales
                                Call WriteLocaleMsg(UserIndex, 605, e_FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
                            End If
                            ''FUISTE
                            Call WriteShowMessageBox(UserIndex, 1783, vbNullString) 'Msg1783=Has sido expulsado por el sistema anti cheats.
                            Call CloseSocket(UserIndex)
                            Exit Sub
                        End If
                        Call FundirMineral(UserIndex)
                    Else
                        ' Msg606=Ahí no hay ninguna fragua.
                        Call WriteLocaleMsg(UserIndex, 606, e_FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(UserIndex, 0)
                        If UserList(UserIndex).Counters.Trabajando > 1 Then
                            Call WriteMacroTrabajoToggle(UserIndex, False)
                        End If
                    End If
                Else
                    ' Msg606=Ahí no hay ninguna fragua.
                    Call WriteLocaleMsg(UserIndex, 606, e_FontTypeNames.FONTTYPE_INFO)
                    Call WriteWorkRequestTarget(UserIndex, 0)
                    If UserList(UserIndex).Counters.Trabajando > 1 Then
                        Call WriteMacroTrabajoToggle(UserIndex, False)
                    End If
                End If
            Case e_Skill.Grupo
                Call LookatTile(UserIndex, .pos.Map, x, y)
                'Target whatever is in that tile
                tU = .flags.TargetUser.ArrayIndex
                If IsValidUserRef(.flags.TargetUser) And tU <> UserIndex Then
                    If UserList(UserIndex).Grupo.EnGrupo = False Then
                        If UserList(tU).flags.Muerto = 0 Then
                            If Abs(.pos.x - x) + Abs(.pos.y - y) > 8 Then
                                Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                                Call WriteWorkRequestTarget(UserIndex, 0)
                                Exit Sub
                            End If
                            If UserList(UserIndex).Grupo.CantidadMiembros = 0 Then
                                Call SetUserRef(UserList(UserIndex).Grupo.Lider, UserIndex)
                                Call SetUserRef(UserList(UserIndex).Grupo.Miembros(1), UserIndex)
                                UserList(UserIndex).Grupo.CantidadMiembros = 1
                                Call InvitarMiembro(UserIndex, tU)
                            Else
                                Call SetUserRef(UserList(UserIndex).Grupo.Lider, UserIndex)
                                Call InvitarMiembro(UserIndex, tU)
                            End If
                        Else
                            Call WriteLocaleMsg(UserIndex, 7, e_FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(UserIndex, 0)
                        End If
                    Else
                        If UserList(UserIndex).Grupo.Lider.ArrayIndex = UserIndex Then
                            Call InvitarMiembro(UserIndex, tU)
                        Else
                            'Msg1131= Tu no podés invitar usuarios, debe hacerlo ¬1
                            Call WriteLocaleMsg(UserIndex, 1131, e_FontTypeNames.FONTTYPE_INFO, UserList(UserList(UserIndex).Grupo.Lider.ArrayIndex).name)
                            Call WriteWorkRequestTarget(UserIndex, 0)
                        End If
                    End If
                Else
                    Call WriteLocaleMsg(UserIndex, 261, e_FontTypeNames.FONTTYPE_INFO)
                End If
            Case e_Skill.MarcaDeClan
                'Target whatever is in that tile
                Dim clan_nivel As Byte
                If UserList(UserIndex).GuildIndex = 0 Then
                    ' Msg720=Servidor » No perteneces a ningún clan.
                    Call WriteLocaleMsg(UserIndex, 720, e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
                End If
                clan_nivel = modGuilds.NivelDeClan(UserList(UserIndex).GuildIndex)
                If clan_nivel < 3 Then
                    ' Msg721=Servidor » El nivel de tu clan debe ser 3 para utilizar esta opción.
                    Call WriteLocaleMsg(UserIndex, 721, e_FontTypeNames.FONTTYPE_INFOIAO)
                    Exit Sub
                End If
                Call LookatTile(UserIndex, UserList(UserIndex).pos.Map, x, y)
                If Not IsValidUserRef(.flags.TargetUser) Then Exit Sub
                tU = .flags.TargetUser.ArrayIndex
                If UserList(UserIndex).GuildIndex = UserList(tU).GuildIndex Then
                    'Msg1132= Servidor » No podes marcar a un miembro de tu clan.
                    Call WriteLocaleMsg(UserIndex, 1132, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If tU > 0 And tU <> UserIndex Then
                    If UserList(tU).flags.AdminInvisible <> 0 Then Exit Sub
                    'Can't steal administrative players
                    If UserList(tU).flags.Muerto = 0 Then
                        'call marcar
                        If UserList(tU).flags.invisible = 1 Or UserList(tU).flags.Oculto = 1 Then
                            UserList(UserIndex).Counters.timeFx = 3
                            Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageParticleFX(UserList(tU).Char.charindex, 210, 50, False, , UserList(UserIndex).pos.x, _
                                    UserList(UserIndex).pos.y))
                        Else
                            UserList(UserIndex).Counters.timeFx = 3
                            Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageParticleFX(UserList(tU).Char.charindex, 210, 150, False, , UserList(UserIndex).pos.x, _
                                    UserList(UserIndex).pos.y))
                        End If
                        Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageLocaleMsg(1798, UserList(UserIndex).name & "¬" & UserList(tU).name, _
                                e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1798=Clan> [¬1] marcó a ¬2.
                    Else
                        Call WriteLocaleMsg(UserIndex, 7, e_FontTypeNames.FONTTYPE_INFO)
                        Call WriteWorkRequestTarget(UserIndex, 0)
                    End If
                Else
                    Call WriteLocaleMsg(UserIndex, 261, e_FontTypeNames.FONTTYPE_INFO)
                End If
            Case e_Skill.MarcaDeGM
                Call LookatTile(UserIndex, UserList(UserIndex).pos.Map, x, y)
                tU = .flags.TargetUser.ArrayIndex
                If IsValidUserRef(.flags.TargetUser) Then
                    'Msg1133= Servidor » [¬1
                    Call WriteLocaleMsg(UserIndex, 1133, e_FontTypeNames.FONTTYPE_INFO, UserList(tU).name)
                Else
                    Call WriteLocaleMsg(UserIndex, 261, e_FontTypeNames.FONTTYPE_INFO)
                End If
            Case e_Skill.TargetableItem
                If .Stats.MinSta < ObjData(.invent.Object(.flags.TargetObjInvSlot).ObjIndex).MinSta Then
                    Call WriteLocaleMsg(UserIndex, MsgNotEnoughtStamina, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg2129=¡No tengo energía!
                    Call SendData(SendTarget.ToIndex, UserIndex, PrepareLocalizedChatOverHead(2129, UserList(UserIndex).Char.charindex, vbWhite))
                    Exit Sub
                End If
                Call LookatTile(UserIndex, UserList(UserIndex).pos.Map, x, y)
                Call UserTargetableItem(UserIndex, x, y)
        End Select
    End With
    Exit Sub
HandleWorkLeftClick_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleWorkLeftClick", Erl)
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleCreateNewGuild(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Desc       As String
        Dim GuildName  As String
        Dim errorStr   As String
        Dim Alineacion As Byte
        Desc = reader.ReadString8()
        GuildName = reader.ReadString8()
        Alineacion = reader.ReadInt8()
        If modGuilds.CrearNuevoClan(UserIndex, Desc, GuildName, Alineacion, errorStr) Then
            Call QuitarObjetos(407, 1, UserIndex)
            Call QuitarObjetos(408, 1, UserIndex)
            Call QuitarObjetos(409, 1, UserIndex)
            Call QuitarObjetos(412, 1, UserIndex)
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageLocaleMsg(1642, .name & "¬" & GuildName & "¬" & GuildAlignment(.GuildIndex), _
                    e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1642=¬1 ha fundado el clan <¬2> de alineación ¬3.
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
            'Update tag
            Call RefreshCharStatus(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(errorStr, vbNullString, e_FontTypeNames.FONTTYPE_GUILD))
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateNewGuild", Erl)
End Sub

''
' Handles the "SpellInfo" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleSpellInfo(ByVal UserIndex As Integer)
    On Error GoTo HandleSpellInfo_Err
    With UserList(UserIndex)
        Dim spellSlot As Byte
        Dim Spell     As Integer
        spellSlot = reader.ReadInt8()
        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            'Msg1134= ¡Primero selecciona el hechizo!
            Call WriteLocaleMsg(UserIndex, 1134, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(spellSlot)
        If Spell > 0 And Spell < NumeroHechizos + 1 Then
            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(UserIndex, "HECINF*" & Spell, e_FontTypeNames.FONTTYPE_INFO)
            End With
        End If
    End With
    Exit Sub
HandleSpellInfo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSpellInfo", Erl)
End Sub

''
' Handles the "EquipItem" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleEquipItem(ByVal UserIndex As Integer)

Dim bSkins                      As Boolean
Dim itemSlot                    As Byte
Dim Packet_ID                   As Long
Dim PacketCounter               As Long
Dim eSkinType                   As e_OBJType

    On Error GoTo HandleEquipItem_Err

    With UserList(UserIndex)

        itemSlot = reader.ReadInt8()
        bSkins = reader.ReadBool

        If bSkins Then
            eSkinType = reader.ReadInt8()
        End If

        PacketCounter = reader.ReadInt32
        Packet_ID = PacketNames.EquipItem

        'Dead users can't equip items
        If .flags.Muerto = 1 Then
            'Msg1136= ¡¡Estás muerto!! Sólo podés usar items cuando estás vivo.
            Call WriteLocaleMsg(UserIndex, 1136, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Validate item slot
        If Not bSkins Then
            If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
            'Auto Fix errores de dateos en ï¿½tems.
            If .invent.Object(itemSlot).amount = 0 Then
                .invent.Object(itemSlot).ObjIndex = 0
                Call UpdateSingleItemInv(UserIndex, itemSlot, False)
                Exit Sub
            End If
            Call EquiparInvItem(UserIndex, itemSlot)
        Else
            If itemSlot > MAX_SKINSINVENTORY_SLOTS Or itemSlot < 1 Then Exit Sub
            If .Invent_Skins.Object(itemSlot).ObjIndex = 0 Then Exit Sub
            If .Invent_Skins.Object(itemSlot).Equipped Then
                Call DesequiparSkin(UserIndex, itemSlot)
                Exit Sub
            End If
            If CanEquipSkin(UserIndex, itemSlot, False) Then
                Call SkinEquip(UserIndex, itemSlot, .Invent_Skins.Object(itemSlot).ObjIndex)
            End If
        End If

    End With
    Exit Sub
HandleEquipItem_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleEquipItem", Erl)
End Sub

''
' Handles the "Change_Heading" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleChange_Heading(ByVal UserIndex As Integer)
    On Error GoTo HandleChange_Heading_Err
    'Se cancela la salida del juego si el user esta saliendo
    With UserList(UserIndex)
        Dim Heading As e_Heading
        Heading = reader.ReadInt8()
        Dim PacketCounter As Long
        PacketCounter = reader.ReadInt32
        Dim Packet_ID As Long
        Packet_ID = PacketNames.ChangeHeading
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "ChangeHeading", PacketTimerThreshold( _
                Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        If .flags.Paralizado > 0 Then
            Exit Sub
        End If
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If Heading > 0 And Heading < 5 Then
            .Char.Heading = Heading
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(.Char.body, .Char.head, .Char.Heading, .Char.charindex, .Char.WeaponAnim, _
                    .Char.ShieldAnim, .Char.CartAnim, .Char.BackpackAnim, .Char.FX, .Char.loops, .Char.CascoAnim, False, .flags.Navegando))
        End If
    End With
    Exit Sub
HandleChange_Heading_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChange_Heading", Erl)
End Sub

''
' Handles the "ModifySkills" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleModifySkills(ByVal UserIndex As Integer)
    On Error GoTo HandleModifySkills_Err
    With UserList(UserIndex)
        Dim i                      As Long
        Dim count                  As Integer
        Dim points(1 To NUMSKILLS) As Byte
        'Codigo para prevenir el hackeo de los skills
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        For i = 1 To NUMSKILLS
            points(i) = reader.ReadInt8()
            If points(i) < 0 Then
                Call LogSecurity(.name & " IP:" & .ConnectionDetails.IP & " trató de hackear los skills.")
                .Stats.SkillPts = 0
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            count = count + points(i)
        Next i
        If count > .Stats.SkillPts Then
            Call LogSecurity(.name & " IP:" & .ConnectionDetails.IP & " trató de hackear los skills.")
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        With .Stats
            For i = 1 To NUMSKILLS
                .SkillPts = .SkillPts - points(i)
                If .UserSkills(i) <> .UserSkills(i) + points(i) Then
                    .UserSkills(i) = .UserSkills(i) + points(i)
                    'Client should prevent this, but just in case...
                    If .UserSkills(i) > 100 Then
                        .SkillPts = .SkillPts + .UserSkills(i) - 100
                        .UserSkills(i) = 100
                    End If
                    UserList(UserIndex).flags.ModificoSkills = True
                End If
            Next i
        End With
    End With
    Exit Sub
HandleModifySkills_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleModifySkills", Erl)
End Sub

''
' Handles the "Train" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleTrain(ByVal UserIndex As Integer)
    On Error GoTo HandleTrain_Err
    With UserList(UserIndex)
        Dim SpawnedNpc As Integer
        Dim PetIndex   As Byte
        PetIndex = reader.ReadInt8()
        If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Entrenador Then Exit Sub
        If NpcList(.flags.TargetNPC.ArrayIndex).Mascotas < MAXMASCOTASENTRENADOR Then
            If PetIndex > 0 And PetIndex < NpcList(.flags.TargetNPC.ArrayIndex).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(NpcList(.flags.TargetNPC.ArrayIndex).Criaturas(PetIndex).NpcIndex, NpcList(.flags.TargetNPC.ArrayIndex).pos, True, False)
                If SpawnedNpc > 0 Then
                    NpcList(SpawnedNpc).MaestroNPC = .flags.TargetNPC
                    NpcList(.flags.TargetNPC.ArrayIndex).Mascotas = NpcList(.flags.TargetNPC.ArrayIndex).Mascotas + 1
                End If
            End If
        Else
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareLocalizedChatOverHead(2082, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite))
        End If
    End With
    Exit Sub
HandleTrain_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleTrain", Erl)
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)
    On Error GoTo HandleCommerceBuy_Err
    With UserList(UserIndex)
        Dim Slot   As Byte
        Dim amount As Integer
        Slot = reader.ReadInt8()
        amount = reader.ReadInt16()
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'El target es un NPC valido?
        If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
        'íEl NPC puede comerciar?
        If NpcList(.flags.TargetNPC.ArrayIndex).Comercia = 0 Then
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareLocalizedChatOverHead(2084, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite))
            Exit Sub
        End If
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            'Msg1137= No estás comerciando
            Call WriteLocaleMsg(UserIndex, 1137, e_FontTypeNames.FONTTYPE_INFO)
            Call WriteCommerceEnd(UserIndex)
            Exit Sub
        End If
        'User compra el item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC.ArrayIndex, Slot, amount)
    End With
    Exit Sub
HandleCommerceBuy_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceBuy", Erl)
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)
    On Error GoTo HandleBankExtractItem_Err
    With UserList(UserIndex)
        Dim Slot        As Byte
        Dim slotdestino As Byte
        Dim amount      As Integer
        Slot = reader.ReadInt8()
        amount = reader.ReadInt16()
        slotdestino = reader.ReadInt8()
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        '¿El target es un NPC valido?
        If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
        '¿Es el banquero?
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then Exit Sub
        'User retira el item del slot
        Call UserRetiraItem(UserIndex, Slot, amount, slotdestino)
    End With
    Exit Sub
HandleBankExtractItem_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankExtractItem", Erl)
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleCommerceSell(ByVal UserIndex As Integer)
    On Error GoTo HandleCommerceSell_Err
    With UserList(UserIndex)
        Dim Slot   As Byte
        Dim amount As Integer
        Slot = reader.ReadInt8()
        amount = reader.ReadInt16()
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'íEl target es un NPC valido?
        If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
        'íEl NPC puede comerciar?
        If NpcList(.flags.TargetNPC.ArrayIndex).Comercia = 0 Then
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareLocalizedChatOverHead(2084, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite))
            Exit Sub
        End If
        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC.ArrayIndex, Slot, amount)
    End With
    Exit Sub
HandleCommerceSell_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceSell", Erl)
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleBankDeposit(ByVal UserIndex As Integer)
    On Error GoTo HandleBankDeposit_Err
    With UserList(UserIndex)
        Dim Slot        As Byte
        Dim slotdestino As Byte
        Dim amount      As Integer
        Slot = reader.ReadInt8()
        amount = reader.ReadInt16()
        slotdestino = reader.ReadInt8()
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'íEl target es un NPC valido?
        If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
        'íEl NPC puede comerciar?
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then
            Exit Sub
        End If
        If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'User deposita el item del slot rdata
        Call UserDepositaItem(UserIndex, Slot, amount, slotdestino)
    End With
    Exit Sub
HandleBankDeposit_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankDeposit", Erl)
End Sub

''
' Handles the "ForumPost" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleForumPost(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim File     As String
        Dim title    As String
        Dim Msg      As String
        Dim postFile As String
        Dim handle   As Integer
        Dim i        As Long
        Dim count    As Integer
        title = reader.ReadString8()
        Msg = reader.ReadString8()
        If .flags.TargetObj > 0 Then
            File = App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
            If FileExist(File, vbNormal) Then
                count = val(GetVar(File, "INFO", "CantMSG"))
                'If there are too many messages, delete the forum
                If count > MAX_MENSAJES_FORO Then
                    For i = 1 To count
                        Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & i & ".for"
                    Next i
                    Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
                    count = 0
                End If
            Else
                'Starting the forum....
                count = 0
            End If
            handle = FreeFile()
            postFile = Left$(File, Len(File) - 4) & CStr(count + 1) & ".for"
            'Create file
            Open postFile For Output As handle
            Print #handle, title
            Print #handle, Msg
            Close #handle
            'Update post count
            Call WriteVar(File, "INFO", "CantMSG", count + 1)
        End If
    End With
    Exit Sub
ErrHandler:
    Close #handle
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleForumPost", Erl)
End Sub

''
' Handles the "MoveSpell" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleMoveSpell(ByVal UserIndex As Integer)
    On Error GoTo HandleMoveSpell_Err
    Dim dir As Integer
    If reader.ReadBool() Then
        dir = 1
    Else
        dir = -1
    End If
    Call DesplazarHechizo(UserIndex, dir, reader.ReadInt8())
    Exit Sub
HandleMoveSpell_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleMoveSpell", Erl)
End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleClanCodexUpdate(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Desc As String
        Desc = reader.ReadString8()
        Call modGuilds.ChangeCodexAndDesc(Desc, .GuildIndex)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleMoveSpell", Erl)
End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)
    On Error GoTo HandleUserCommerceOffer_Err
    With UserList(UserIndex)
        Dim tUser         As Integer
        Dim Slot          As Byte
        Dim amount        As Long
        Dim ElementalTags As Long
        Slot = reader.ReadInt8()
        amount = reader.ReadInt32()
        If Slot <> FLAGORO Then
            'Natural elemental tags are the one in the object
            'User added elemental tags are the one in the user slots
            ElementalTags = UserList(UserIndex).invent.Object(Slot).ElementalTags
        End If
        'Is the commerce attempt valid??
        If Not IsValidUserRef(.ComUsu.DestUsu) Then
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        End If
        'Get the other player
        tUser = .ComUsu.DestUsu.ArrayIndex
        If UserList(tUser).ComUsu.DestUsu.ArrayIndex <> UserIndex Then
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        End If
        'If Amount is invalid, or slot is invalid and it's not gold, then ignore it.
        If ((Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots) And Slot <> FLAGORO) Or amount <= 0 Then Exit Sub
        'Is the other player valid??
        If tUser < 1 Or tUser > MaxUsers Then Exit Sub
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
            If Slot = FLAGORO Then
                'gold
                If amount > .Stats.GLD Then
                    'Msg1138= No tienes esa cantidad.
                    Call WriteLocaleMsg(UserIndex, 1138, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Else
                'inventory
                If amount > .invent.Object(Slot).amount Then
                    'Msg1139= No tienes esa cantidad.
                    Call WriteLocaleMsg(UserIndex, 1139, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .invent.Object(Slot).ObjIndex > 0 Then
                    If ObjData(.invent.Object(Slot).ObjIndex).Instransferible = 1 Then
                        'Msg1140= Este objeto es intransferible, no podés venderlo.
                        Call WriteLocaleMsg(UserIndex, 1140, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    If ObjData(.invent.Object(Slot).ObjIndex).Newbie = 1 Then
                        'Msg1141= No puedes comerciar objetos newbie.
                        Call WriteLocaleMsg(UserIndex, 1141, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
            End If
            'Prevent offer changes (otherwise people would ripp off other players)
            'If .ComUsu.Objeto > 0 Then
            'Msg1142= No podés cambiar tu oferta.
            Call WriteLocaleMsg(UserIndex, 1142, e_FontTypeNames.FONTTYPE_INFO)
            '  End If
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .invent.EquippedShipSlot = Slot Then
                    'Msg1143= No podés vender tu barco mientras lo estás usando.
                    Call WriteLocaleMsg(UserIndex, 1143, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
            If .flags.Montado = 1 Then
                If .invent.EquippedSaddleSlot = Slot Then
                    'Msg1144= No podés vender tu montura mientras la estás usando.
                    Call WriteLocaleMsg(UserIndex, 1144, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
            .ComUsu.Objeto = Slot
            .ComUsu.cant = amount
            'If the other one had accepted, we turn that back and inform of the new offer (just to be cautious).
            If UserList(tUser).ComUsu.Acepto Then
                UserList(tUser).ComUsu.Acepto = False
                Call WriteConsoleMsg(tUser, PrepareMessageLocaleMsg(1951, .name, e_FontTypeNames.FONTTYPE_TALK)) ' Msg1951=¬1 ha cambiado su oferta.
            End If
            Dim ObjAEnviar As t_Obj
            ObjAEnviar.amount = amount
            ObjAEnviar.ElementalTags = ElementalTags
            'Si no es oro tmb le agrego el objInex
            If Slot <> FLAGORO Then ObjAEnviar.ObjIndex = UserList(UserIndex).invent.Object(Slot).ObjIndex
            'Llamos a la funcion
            Call EnviarObjetoTransaccion(tUser, UserIndex, ObjAEnviar)
        End If
    End With
    Exit Sub
HandleUserCommerceOffer_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleUserCommerceOffer", Erl)
End Sub

Private Sub HandleGuildAcceptPeace(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        guild = reader.ReadString8()
        Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1799, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1799=No se pueden actualizar relaciones.
        Exit Sub
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1800, guild, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1800=Tu clan ha firmado la paz con ¬1
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageLocaleMsg(1800, modGuilds.GuildName(.GuildIndex), e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1800=Tu clan ha firmado la paz con ¬1
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptPeace", Erl)
End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildRejectAlliance(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        guild = reader.ReadString8()
        Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1801, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1801=Relaciones de clan desactivadas por el momento.
        Exit Sub
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1802, guild, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1802=Tu clan ha rechazado la propuesta de alianza de ¬1
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageLocaleMsg(1803, modGuilds.GuildName(.GuildIndex), e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1803=¬1 ha rechazado nuestra propuesta de alianza con su clan.
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRejectAlliance", Erl)
End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildRejectPeace(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        guild = reader.ReadString8()
        Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1801, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1801=Relaciones de clan desactivadas por el momento.
        Exit Sub
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1804, guild, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1804=Tu clan ha rechazado la propuesta de paz de ¬1
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageLocaleMsg(1805, modGuilds.GuildName(.GuildIndex), e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1805=¬1 ha rechazado nuestra propuesta de paz con su clan.
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRejectPeace", Erl)
End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildAcceptAlliance(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim guild          As String
        Dim errorStr       As String
        Dim otherClanIndex As String
        guild = reader.ReadString8()
        Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1801, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1801=Relaciones de clan desactivadas por el momento.
        Exit Sub
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1806, guild, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1806=Tu clan ha firmado la alianza con ¬1
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageLocaleMsg(1800, modGuilds.GuildName(.GuildIndex), e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1800=Tu clan ha firmado la paz con ¬1
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptAlliance", Erl)
End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildOfferPeace(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim guild    As String
        Dim proposal As String
        Dim errorStr As String
        guild = reader.ReadString8()
        proposal = reader.ReadString8()
        Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1801, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1801=Relaciones de clan desactivadas por el momento.
        Exit Sub
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOfferPeace", Erl)
End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildOfferAlliance(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim guild    As String
        Dim proposal As String
        Dim errorStr As String
        guild = reader.ReadString8()
        proposal = reader.ReadString8()
        'Msg1145= Relaciones de clan desactivadas por el momento.
        Call WriteLocaleMsg(UserIndex, 1145, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOfferPeace", Erl)
End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildAllianceDetails(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim guild    As String
        Dim errorStr As String
        Dim details  As String
        guild = reader.ReadString8()
        'Msg1146= Relaciones de clan desactivadas por el momento.
        Call WriteLocaleMsg(UserIndex, 1146, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOfferPeace", Erl)
End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildPeaceDetails(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim guild    As String
        Dim errorStr As String
        Dim details  As String
        guild = reader.ReadString8()
        'Msg1147= Relaciones de clan desactivadas por el momento.
        Call WriteLocaleMsg(UserIndex, 1147, e_FontTypeNames.FONTTYPE_INFO)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildPeaceDetails", Erl)
End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildRequestJoinerInfo(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim User    As String
        Dim details As String
        User = reader.ReadString8()
        details = modGuilds.a_DetallesAspirante(UserIndex, User)
        If LenB(details) = 0 Then
            'Msg1148= El personaje no ha mandado solicitud, o no estás habilitado para verla.
            Call WriteLocaleMsg(UserIndex, 1148, e_FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteShowUserRequest(UserIndex, details)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRequestJoinerInfo", Erl)
End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildAlliancePropList(ByVal UserIndex As Integer)
    On Error GoTo HandleGuildAlliancePropList_Err
    'Msg1149= Relaciones de clan desactivadas por el momento.
    Call WriteLocaleMsg(UserIndex, 1149, e_FontTypeNames.FONTTYPE_INFO)
    Exit Sub
HandleGuildAlliancePropList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAlliancePropList", Erl)
End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildPeacePropList(ByVal UserIndex As Integer)
    On Error GoTo HandleGuildPeacePropList_Err
    'Msg1150= Relaciones de clan desactivadas por el momento.
    Call WriteLocaleMsg(UserIndex, 1150, e_FontTypeNames.FONTTYPE_INFO)
    Exit Sub
HandleGuildPeacePropList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildPeacePropList", Erl)
End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildDeclareWar(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim guild           As String
        Dim errorStr        As String
        Dim otherGuildIndex As Integer
        guild = reader.ReadString8()
        'Msg1151= Relaciones de clan desactivadas por el momento.
        Call WriteLocaleMsg(UserIndex, 1151, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1807, guild, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1807=TU CLAN HA ENTRADO EN GUERRA CON ¬1
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageLocaleMsg(1808, modGuilds.GuildName(.GuildIndex), e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1808=¬1 LE DECLARA LA GUERRA A TU CLAN
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildPeacePropList", Erl)
End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildNewWebsite(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Call modGuilds.ActualizarWebSite(UserIndex, reader.ReadString8())
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildNewWebsite", Erl)
End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildAcceptNewMember(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim errorStr As String
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        tUser = NameIndex(username)
        If IsValidUserRef(tUser) Then
            If Not modGuilds.a_AceptarAspirante(UserIndex, username, errorStr) Then
                Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
            Else
                Call modGuilds.m_ConectarMiembroAClan(tUser.ArrayIndex, .GuildIndex)
                Call RefreshCharStatus(tUser.ArrayIndex)
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1809, username, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1809=[¬1] ha sido aceptado como miembro del clan.
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))
            End If
        Else
            If Not modGuilds.a_AceptarAspirante(UserIndex, username, errorStr) Then
                Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1809, username, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1809=[¬1] ha sido aceptado como miembro del clan.
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptNewMember", Erl)
End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildRejectNewMember(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim errorStr As String
        Dim username As String
        Dim Reason   As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        Reason = reader.ReadString8()
        If Not modGuilds.a_RechazarAspirante(UserIndex, username, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(username)
            If IsValidUserRef(tUser) Then
                Call WriteConsoleMsg(tUser.ArrayIndex, errorStr & " : " & Reason, e_FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(username, .GuildIndex, Reason)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptNewMember", Erl)
End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildKickMember(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username   As String
        Dim GuildIndex As Integer
        username = reader.ReadString8()
        Dim CharId As Long
        CharId = GetCharacterIdWithName(username)
        If CharId <= 0 Then
            Exit Sub
        End If
        GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, CharId)
        If GuildIndex > 0 Then
            Dim expulsado As t_UserReference
            expulsado = NameIndex(username)
            'Msg1152= Has sido expulsado del clan.
            Call WriteLocaleMsg(expulsado.ArrayIndex, "1152", e_FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageLocaleMsg(1810, username, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1810=¬1 fue expulsado del clan.
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        Else
            'Msg1153= No podés expulsar ese personaje del clan.
            Call WriteLocaleMsg(UserIndex, 1153, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildKickMember", Erl)
End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildUpdateNews(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Call modGuilds.ActualizarNoticias(UserIndex, reader.ReadString8())
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildUpdateNews", Erl)
End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildMemberInfo(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Call modGuilds.SendDetallesPersonaje(UserIndex, reader.ReadString8())
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildMemberInfo", Erl)
End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildOpenElections(ByVal UserIndex As Integer)
    On Error GoTo HandleGuildOpenElections_Err
    With UserList(UserIndex)
        Dim Error As String
        'Msg1154= Elecciones de clan desactivadas por el momento.
        Call WriteLocaleMsg(UserIndex, 1154, e_FontTypeNames.FONTTYPE_INFO)
    End With
    Exit Sub
HandleGuildOpenElections_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOpenElections", Erl)
End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildRequestMembership(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim guild       As String
        Dim application As String
        Dim errorStr    As String
        guild = reader.ReadString8()
        application = reader.ReadString8()
        If Not modGuilds.a_NuevoAspirante(UserIndex, guild, application, errorStr) Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(errorStr, vbNullString, e_FontTypeNames.FONTTYPE_GUILD))
        Else
            'Msg1155= Tu solicitud ha sido enviada. Espera prontas noticias del líder de ¬1
            Call WriteLocaleMsg(UserIndex, 1155, e_FontTypeNames.FONTTYPE_INFO, guild)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRequestMembership", Erl)
End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildRequestDetails(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Call modGuilds.SendGuildDetails(UserIndex, reader.ReadString8())
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRequestDetails", Erl)
End Sub

''
' Handles the "Quit" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleQuit(ByVal UserIndex As Integer)
    On Error GoTo HandleQuit_Err
    'If user is invisible, it automatically becomes
    'visible before doing the countdown to exit
    'No se reseteaban los contadores de invi ni de ocultar.
    Dim tUser As Integer
    With UserList(UserIndex)
        If .flags.Paralizado = 1 Then
            'Msg1156= No podés salir estando paralizado.
            Call WriteLocaleMsg(UserIndex, 1156, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'exit secure commerce
        If .ComUsu.DestUsu.ArrayIndex > 0 Then
            tUser = .ComUsu.DestUsu.ArrayIndex
            If IsValidUserRef(.ComUsu.DestUsu) And UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu.ArrayIndex = UserIndex Then
                    'Msg1157= Comercio cancelado por el otro usuario
                    Call WriteLocaleMsg(tUser, 1157, e_FontTypeNames.FONTTYPE_INFO)
                    Call FinComerciarUsu(tUser)
                End If
            End If
            'Msg1158= Comercio cancelado.
            Call WriteLocaleMsg(UserIndex, 1158, e_FontTypeNames.FONTTYPE_INFO)
            Call FinComerciarUsu(UserIndex)
        End If
        Call Cerrar_Usuario(UserIndex)
    End With
    Exit Sub
HandleQuit_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuit", Erl)
End Sub

''
' Handles the "GuildLeave" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildLeave(ByVal UserIndex As Integer)
    On Error GoTo HandleGuildLeave_Err
    Dim GuildIndex As Integer
    With UserList(UserIndex)
        'obtengo el guildindex
        GuildIndex = m_EcharMiembroDeClan(UserIndex, .Id)
        If GuildIndex > 0 Then
            'Msg1159= Dejas el clan.
            Call WriteLocaleMsg(UserIndex, 1159, e_FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageLocaleMsg(1811, .name, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1811=¬1 deja el clan.
        Else
            'Msg1160= Tu no puedes salir de ningún clan.
            Call WriteLocaleMsg(UserIndex, 1160, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleGuildLeave_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildLeave", Erl)
End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)
    On Error GoTo HandleRequestAccountState_Err
    Dim earnings   As Integer
    Dim percentage As Integer
    With UserList(UserIndex)
        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Validate target NPC
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            'Msg1161= Primero tenes que seleccionar un personaje, hace click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 1161, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 3 Then
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Select Case NpcList(.flags.TargetNPC.ArrayIndex).npcType
            Case e_NPCType.Banquero
                Call WriteLocaleChatOverHead(UserIndex, 1433, "", str$(PonerPuntos(.Stats.Banco)), vbWhite) ' Msg1433=Tenes ¬1 monedas de oro en tu cuenta.
            Case e_NPCType.Timbero
                If Not .flags.Privilegios And e_PlayerType.User Then
                    earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                        percentage = Int(earnings * 100 / Apuestas.Ganancias)
                    End If
                    If earnings < 0 And Apuestas.Perdidas <> 0 Then
                        percentage = Int(earnings * 100 / Apuestas.Perdidas)
                    End If
                    'Msg1162= Entradas: ¬1
                    Call WriteLocaleMsg(UserIndex, 1162, e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(Apuestas.Ganancias))
                End If
        End Select
    End With
    Exit Sub
HandleRequestAccountState_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestAccountState", Erl)
End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandlePetStand(ByVal UserIndex As Integer)
    On Error GoTo HandlePetStand_Err
    With UserList(UserIndex)
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Validate target NPC
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            ' Msg757=Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 757, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Make sure it's close enough
        If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 10 Then
            'Msg1163= Estás demasiado lejos.
            Call WriteLocaleMsg(UserIndex, 1163, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Make sure it's his pet
        If Not IsValidUserRef(NpcList(.flags.TargetNPC.ArrayIndex).MaestroUser) Or NpcList(.flags.TargetNPC.ArrayIndex).MaestroUser.ArrayIndex <> UserIndex Then Exit Sub
        'Do it!
        Call SetMovement(.flags.TargetNPC.ArrayIndex, e_TipoAI.Estatico)
        Call Expresar(.flags.TargetNPC.ArrayIndex, UserIndex)
    End With
    Exit Sub
HandlePetStand_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandlePetStand", Erl)
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandlePetFollow(ByVal UserIndex As Integer)
    On Error GoTo HandlePetFollow_Err
    With UserList(UserIndex)
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Validate target NPC
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            ' Msg757=Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 757, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Make sure it's close enough
        If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 10 Then
            'Msg1164= Estás demasiado lejos.
            Call WriteLocaleMsg(UserIndex, 1164, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Make usre it's the user's pet
        If Not IsValidUserRef(NpcList(.flags.TargetNPC.ArrayIndex).MaestroUser) Or NpcList(.flags.TargetNPC.ArrayIndex).MaestroUser.ArrayIndex <> UserIndex Then Exit Sub
        'Do it
        Call FollowAmo(.flags.TargetNPC.ArrayIndex)
        Call Expresar(.flags.TargetNPC.ArrayIndex, UserIndex)
    End With
    Exit Sub
HandlePetFollow_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandlePetFollow", Erl)
End Sub
Private Sub HandlePetFollowAll(ByVal UserIndex As Integer)
    On Error GoTo HandlePetFollowAll_Err
    With UserList(UserIndex)
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Validate target NPC
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            ' Msg757=Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 757, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Make sure it's close enough
        If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 10 Then
            'Msg1164= Estás demasiado lejos.
            Call WriteLocaleMsg(UserIndex, 1164, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim i As Integer
        Dim NpcIndex As Integer
        For i = 1 To MAXMASCOTAS
            If IsValidNpcRef(.MascotasIndex(i)) Then
                NpcIndex = .MascotasIndex(i).ArrayIndex
                If NpcList(NpcIndex).flags.NPCActive Then
                    If IsValidUserRef(NpcList(NpcIndex).MaestroUser) Then
                        If NpcList(NpcIndex).MaestroUser.ArrayIndex = UserIndex Then
                            Call FollowAmo(NpcIndex)
                            Call Expresar(NpcIndex, UserIndex)
                        End If
                    End If
                End If
            End If
        Next i
    End With
    Exit Sub
HandlePetFollowAll_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandlePetFollowAll", Erl)
End Sub
''
' Handles the "PetLeave" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandlePetLeave(ByVal UserIndex As Integer)
    On Error GoTo HandlePetLeave_Err
    With UserList(UserIndex)
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Validate target NPC
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            ' Msg757=Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 757, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Make usre it's the user's pet
        If Not IsValidUserRef(NpcList(.flags.TargetNPC.ArrayIndex).MaestroUser) Or NpcList(.flags.TargetNPC.ArrayIndex).MaestroUser.ArrayIndex <> UserIndex Then Exit Sub
        Call QuitarNPC(.flags.TargetNPC.ArrayIndex, e_DeleteSource.ePetLeave)
    End With
    Exit Sub
HandlePetLeave_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandlePetLeave", Erl)
End Sub

''
' Handles the "GrupoMsg" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGrupoMsg(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim chat As String
        chat = reader.ReadString8()
        If LenB(chat) <> 0 Then
            If .Grupo.EnGrupo = True Then
                Dim i As Byte
                For i = 1 To UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros
                    Call WriteConsoleMsg(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex, .name & "> " & chat, e_FontTypeNames.FONTTYPE_New_Amarillo_Verdoso)
                    Call WriteChatOverHead(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex, "NOCONSOLA*" & chat, UserList(UserIndex).Char.charindex, &HFF8000)
                Next i
            Else
                ' Msg758=Grupo> No estas en ningun grupo.
                Call WriteLocaleMsg(UserIndex, 758, e_FontTypeNames.FONTTYPE_New_GRUPO)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGrupoMsg", Erl)
End Sub

Private Sub HandleTrainList(ByVal UserIndex As Integer)
    On Error GoTo HandleTrainList_Err
    With UserList(UserIndex)
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Validate target NPC
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            ' Msg757=Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 757, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Make sure it's close enough
        If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Make sure it's the trainer
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Entrenador Then Exit Sub
        Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC.ArrayIndex)
    End With
    Exit Sub
HandleTrainList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleTrainList", Erl)
End Sub

''
' Handles the "Rest" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleRest(ByVal UserIndex As Integer)
    On Error GoTo HandleRest_Err
    With UserList(UserIndex)
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            ' Msg752=¡¡Estás muerto!! Solo podés usar items cuando estás vivo.
            Call WriteLocaleMsg(UserIndex, 752, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If HayOBJarea(.pos, FOGATA) Then
            Call WriteRestOK(UserIndex)
            If Not .flags.Descansar Then
                ' Msg753=Te acomodás junto a la fogata y comenzas a descansar.
                Call WriteLocaleMsg(UserIndex, 753, e_FontTypeNames.FONTTYPE_INFO)
            Else
                ' Msg754=Te levantas.
                Call WriteLocaleMsg(UserIndex, 754, e_FontTypeNames.FONTTYPE_INFO)
            End If
            .flags.Descansar = Not .flags.Descansar
        Else
            If .flags.Descansar Then
                Call WriteRestOK(UserIndex)
                ' Msg754=Te levantas.
                Call WriteLocaleMsg(UserIndex, 754, e_FontTypeNames.FONTTYPE_INFO)
                .flags.Descansar = False
                Exit Sub
            End If
            ' Msg755=No hay ninguna fogata junto a la cual descansar.
            Call WriteLocaleMsg(UserIndex, 755, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleRest_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRest", Erl)
End Sub

''
' Handles the "Meditate" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleMeditate(ByVal UserIndex As Integer)
    On Error GoTo HandleMeditate_Err
    'Arreglí un bug que mandaba un index de la meditacion diferente
    'al que decia el server.
    With UserList(UserIndex)
        'Si ya tiene el mana completo, no lo dejamos meditar.
        If .Stats.MinMAN = .Stats.MaxMAN Then Exit Sub
        'Las clases NO MAGICAS no meditan...
        If .clase = e_Class.Hunter Or .clase = e_Class.Trabajador Or .clase = e_Class.Warrior Or .clase = e_Class.Pirat Or .clase = e_Class.Thief Then Exit Sub
        If .flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.Montado = 1 Then
            ' Msg756=No podes meditar estando montado.
            Call WriteLocaleMsg(UserIndex, 756, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        .flags.Meditando = Not .flags.Meditando
        If .flags.Meditando Then
            .Counters.TimerMeditar = 0
            .Counters.TiempoInicioMeditar = 0
            Dim customEffect As Integer
            Dim Index        As Integer
            Dim obj          As t_ObjData
            For Index = 1 To UBound(.invent.Object)
                If .invent.Object(Index).ObjIndex > 0 Then
                    If .invent.Object(Index).ObjIndex > 0 Then
                        obj = ObjData(.invent.Object(Index).ObjIndex)
                        If obj.OBJType = otDonator And obj.Subtipo = 4 And .invent.Object(Index).Equipped Then
                            customEffect = obj.HechizoIndex
                            Exit For
                        End If
                    End If
                End If
            Next Index
            If customEffect > 0 Then
                .Char.FX = customEffect
            Else
                Dim isCriminal As Boolean
                
                isCriminal = (.Faccion.Status = e_Facciones.Caos _
                           Or .Faccion.Status = e_Facciones.Criminal _
                           Or .Faccion.Status = e_Facciones.concilio)
                
                Select Case .Stats.ELV
                    Case 1 To 12
                        .Char.FX = IIf(isCriminal, MeditationCriminalLevel1to12, MeditationLevel1to12)
                    Case 13 To 17
                        .Char.FX = IIf(isCriminal, MeditationCriminalLevel13to17, MeditationLevel13to17)
                    Case 18 To 24
                        .Char.FX = IIf(isCriminal, MeditationCriminalLevel18to24, MeditationLevel18to24)
                    Case 25 To 28
                        .Char.FX = IIf(isCriminal, MeditationCriminalLevel25to28, MeditationLevel25to28)
                    Case 29 To 32
                        .Char.FX = IIf(isCriminal, MeditationCriminalLevel29to32, MeditationLevel29to32)
                    Case 33 To 36
                        .Char.FX = IIf(isCriminal, MeditationCriminalLevel33to36, MeditationLevel33to36)
                    Case 37 To 39
                        .Char.FX = IIf(isCriminal, MeditationCriminalLevel37to39, MeditationLevel37to39)
                    Case 40 To 42
                        .Char.FX = IIf(isCriminal, MeditationCriminalLevel40to42, MeditationLevel40to42)
                    Case 43 To 44
                        .Char.FX = IIf(isCriminal, MeditationCriminalLevel43to44, MeditationLevel43to44)
                    Case 45 To 46
                        .Char.FX = IIf(isCriminal, MeditationCriminalLevel45to46, MeditationLevel45to46)
                    Case Else
                        .Char.FX = MeditationLevelMax
                End Select
            End If
        Else
            .Char.FX = 0
        End If
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, .Char.FX, .pos.x, .pos.y))
    End With
    Exit Sub
HandleMeditate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleMeditate", Erl)
End Sub

''
' Handles the "Resucitate" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleResucitate(ByVal UserIndex As Integer)
    On Error GoTo HandleResucitate_Err
    With UserList(UserIndex)
        'Se asegura que el target es un npc
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            ' Msg757=Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 757, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Validate NPC and make sure player is dead
        If (NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Revividor And (NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.ResucitadorNewbie Or Not EsNewbie( _
                UserIndex))) Or .flags.Muerto = 0 Then Exit Sub
        'Make sure it's close enough
        If Distancia(.pos, NpcList(.flags.TargetNPC.ArrayIndex).pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            'Msg8=Estás muy lejos.
            Exit Sub
        End If
        Call RevivirUsuario(UserIndex)
        UserList(UserIndex).Counters.timeFx = 3
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.Corazones, 100, False, , UserList( _
                UserIndex).pos.x, UserList(UserIndex).pos.y))
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(104, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        ' Msg585=¡Has sido resucitado!
        Call WriteLocaleMsg(UserIndex, 585, e_FontTypeNames.FONTTYPE_INFO)
    End With
    Exit Sub
HandleResucitate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleResucitate", Erl)
End Sub

Private Sub HandleHeal(ByVal UserIndex As Integer)
    On Error GoTo HandleHeal_Err
    With UserList(UserIndex)
        'Se asegura que el target es un npc
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            ' Msg757=Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 757, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If (NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Revividor And NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.ResucitadorNewbie) Or .flags.Muerto _
                <> 0 Then Exit Sub
        If Distancia(.pos, NpcList(.flags.TargetNPC.ArrayIndex).pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        .Stats.MinHp = .Stats.MaxHp
        Call WriteUpdateHP(UserIndex)
        'Msg496=¡¡Hás sido curado!!
        Call WriteLocaleMsg(UserIndex, 496, e_FontTypeNames.FONTTYPE_INFO)
    End With
    Exit Sub
HandleHeal_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleHeal", Erl)
End Sub

''
' Handles the "CommerceStart" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleCommerceStart(ByVal UserIndex As Integer)
    On Error GoTo HandleCommerceStart_Err
    With UserList(UserIndex)
        If IsInMapCarcelRestrictedArea(.pos) Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MSG_CANNOT_TRADE_IN_JAIL, vbNullString, e_FontTypeNames.FONTTYPE_INFO))
            Exit Sub
        End If
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            ''Msg77=¡¡Estás muerto!!.)
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Is it already in commerce mode??
        If .flags.Comerciando Then
            ' Msg759=Ya estás comerciando
            Call WriteLocaleMsg(UserIndex, 759, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Validate target NPC
        If IsValidNpcRef(.flags.TargetNPC) Then
            'VOS, como GM, NO podes COMERCIAR con NPCs. (excepto Admins)
            If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Admin)) = 0 Then
                ' Msg767=No podés vender items.
                Call WriteLocaleMsg(UserIndex, 767, e_FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            'Does the NPC want to trade??
            If NpcList(.flags.TargetNPC.ArrayIndex).Comercia = 0 Then
                If LenB(NpcList(.flags.TargetNPC.ArrayIndex).Desc) <> 0 Then
                    Call WriteLocaleChatOverHead(UserIndex, 1434, "", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite) ' Msg1434=No tengo ningún interés en comerciar.
                End If
                Exit Sub
            End If
            If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 3 Then
                Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
        ElseIf IsValidUserRef(.flags.TargetUser) Then
            ' **********************  Comercio con Usuarios  *********************
            'VOS, como GM, NO podes COMERCIAR con usuarios. (excepto  Admins)
            If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Admin)) = 0 Then
                ' Msg767=No podés vender items.
                Call WriteLocaleMsg(UserIndex, 767, e_FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            'NO podes COMERCIAR CON un GM. (excepto  Admins)
            If (UserList(.flags.TargetUser.ArrayIndex).flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Admin)) = 0 Then
                'Msg1165= No podés vender items a este usuario.
                Call WriteLocaleMsg(UserIndex, 1165, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'Is the other one dead??
            If UserList(.flags.TargetUser.ArrayIndex).flags.Muerto = 1 Then
                Call FinComerciarUsu(.flags.TargetUser.ArrayIndex, True)
                'Msg1166= ¡¡No podés comerciar con los muertos!!
                Call WriteLocaleMsg(UserIndex, 1166, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'Is it me??
            If .flags.TargetUser.ArrayIndex = UserIndex Then
                'Msg1167= No podés comerciar con vos mismo...
                Call WriteLocaleMsg(UserIndex, 1167, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'Check distance
            If .pos.Map <> UserList(.flags.TargetUser.ArrayIndex).pos.Map Or Distancia(UserList(.flags.TargetUser.ArrayIndex).pos, .pos) > 3 Then
                Call FinComerciarUsu(.flags.TargetUser.ArrayIndex, True)
                Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'Check if map is not safe
            If MapInfo(.pos.Map).Seguro = 0 Then
                Call FinComerciarUsu(.flags.TargetUser.ArrayIndex, True)
                'Msg1168= No se puede usar el comercio seguro en zona insegura.
                Call WriteLocaleMsg(UserIndex, 1168, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser.ArrayIndex).flags.Comerciando = True Then
                Call FinComerciarUsu(.flags.TargetUser.ArrayIndex, True)
                'Msg1169= No podés comerciar con el usuario en este momento.
                Call WriteLocaleMsg(UserIndex, 1169, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser.ArrayIndex).name
            .ComUsu.cant = 0
            .ComUsu.Objeto = 0
            .ComUsu.Acepto = False
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser.ArrayIndex)
        Else
            ' Msg760=Primero haz click izquierdo sobre el personaje.
            Call WriteLocaleMsg(UserIndex, 760, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleCommerceStart_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceStart", Erl)
End Sub

Private Sub HandleBankStart(ByVal UserIndex As Integer)
    On Error GoTo HandleBankStart_Err
    With UserList(UserIndex)
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.Comerciando Then
            ' Msg759=Ya estás comerciando
            Call WriteLocaleMsg(UserIndex, 759, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Validate target NPC
        If IsValidNpcRef(.flags.TargetNPC) Then
            If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 6 Then
                Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'If it's the banker....
            If NpcList(.flags.TargetNPC.ArrayIndex).npcType = e_NPCType.Banquero Then
                Call IniciarDeposito(UserIndex)
            End If
        Else
            ' Msg760=Primero haz click izquierdo sobre el personaje.
            Call WriteLocaleMsg(UserIndex, 760, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleBankStart_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankStart", Erl)
End Sub

Private Sub HandleEnlist(ByVal UserIndex As Integer)
    On Error GoTo HandleEnlist_Err
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
        'Validate target NPC
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            ' Msg761=Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 761, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        If Distancia(.pos, NpcList(.flags.TargetNPC.ArrayIndex).pos) > 4 Then
            'Msg1170= Debes acercarte más.
            Call WriteLocaleMsg(UserIndex, 1170, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If NpcList(.flags.TargetNPC.ArrayIndex).flags.Faccion = 0 Then
            Call EnlistarArmadaReal(UserIndex)
        Else
            Call EnlistarCaos(UserIndex)
        End If
    End With
    Exit Sub
HandleEnlist_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleEnlist", Erl)
End Sub

Private Sub HandleInformation(ByVal UserIndex As Integer)
    On Error GoTo HandleInformation_Err
    With UserList(UserIndex)
        'Validate target NPC
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            ' Msg761=Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 761, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        If Distancia(.pos, NpcList(.flags.TargetNPC.ArrayIndex).pos) > 4 Then
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If NpcList(.flags.TargetNPC.ArrayIndex).flags.Faccion = 0 Then
            If .Faccion.Status <> e_Facciones.Armada Or .Faccion.Status <> e_Facciones.consejo Then
                Call WriteLocaleChatOverHead(UserIndex, 1389, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1389=No perteneces a las tropas reales!!!
                Exit Sub
            End If
            Call WriteLocaleChatOverHead(UserIndex, 1390, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1390=Tu deber es combatir criminales, cada 100 criminales que derrotes te darí una recompensa.
        Else
            If .Faccion.Status <> e_Facciones.Caos Or .Faccion.Status <> e_Facciones.concilio Then
                Call WriteLocaleChatOverHead(UserIndex, 1391, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1391=No perteneces a la legión oscura!!!
                Exit Sub
            End If
            Call WriteLocaleChatOverHead(UserIndex, 1392, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1392=Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te darí una recompensa.
        End If
    End With
    Exit Sub
HandleInformation_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleInformation", Erl)
End Sub

''
' Handles the "Reward" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleReward(ByVal UserIndex As Integer)
    On Error GoTo HandleReward_Err
    With UserList(UserIndex)
        'Validate target NPC
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            ' Msg761=Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 761, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        If Distancia(.pos, NpcList(.flags.TargetNPC.ArrayIndex).pos) > 4 Then
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If NpcList(.flags.TargetNPC.ArrayIndex).flags.Faccion = 0 Then
            If .Faccion.Status <> e_Facciones.Armada And .Faccion.Status <> e_Facciones.consejo Then
                Call WriteLocaleChatOverHead(UserIndex, 1393, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1393=No perteneces a las tropas reales!!!
                Exit Sub
            End If
            Call RecompensaArmadaReal(UserIndex)
        Else
            If .Faccion.Status <> e_Facciones.Caos And .Faccion.Status <> e_Facciones.concilio Then
                Call WriteLocaleChatOverHead(UserIndex, 1394, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1394=No perteneces a la legión oscura!!!
                Exit Sub
            End If
            Call RecompensaCaos(UserIndex)
        End If
    End With
    Exit Sub
HandleReward_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleReward", Erl)
End Sub

''
' Handles the "GuildMessage" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildMessage(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim chat As String
        chat = reader.ReadString8()
        Dim PacketCounter As Long
        PacketCounter = reader.ReadInt32
        Dim Packet_ID As Long
        Packet_ID = PacketNames.GuildMessage
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "GuildMessage", PacketTimerThreshold( _
                Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        If LenB(chat) <> 0 Then
            '  Foto-denuncias - Push message
            Dim i As Integer
            For i = 1 To UBound(.flags.ChatHistory) - 1
                .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
            Next
            .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
            If .GuildIndex > 0 Then
                'HarThaoS: si es leade mando un 10 para el status del color(medio villero pero me dio paja)
                If LCase(GuildLeader(.GuildIndex)) = .name Then
                    Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.name & "> " & chat, 10))
                Else
                    .Counters.timeGuildChat = 1 + Ceil((3000 + 60 * Len(chat)) / 1000)
                    Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.name & "> " & chat, .Faccion.Status))
                    Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageChatOverHead("NOCONSOLA*< " & chat & " >", .Char.charindex, RGB(255, 255, 0), , .pos.x, .pos.y))
                End If
                'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
                ' Call SendData(SendTarget.ToAll, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "í< " & chat & " >í" & CStr(UserList(UserIndex).Char.CharIndex))
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildMessage", Erl)
End Sub

''
' Handles the "GuildOnline" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildOnline(ByVal UserIndex As Integer)
    On Error GoTo HandleGuildOnline_Err
    With UserList(UserIndex)
        Dim onlineList As String
        onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .GuildIndex)
        If .GuildIndex <> 0 Then
            'Msg1171= Compañeros de tu clan conectados: ¬1
            Call WriteLocaleMsg(UserIndex, 1171, e_FontTypeNames.FONTTYPE_INFO, onlineList)
        Else
            ' Msg762=No pertences a ningún clan.
            Call WriteLocaleMsg(UserIndex, 762, e_FontTypeNames.FONTTYPE_GUILDMSG)
        End If
    End With
    Exit Sub
HandleGuildOnline_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOnline", Erl)
End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleCouncilMessage(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim chat As String
        chat = reader.ReadString8()
        If LenB(chat) <> 0 Then
            '  Foto-denuncias - Push message
            Dim i As Long
            For i = 1 To UBound(.flags.ChatHistory) - 1
                .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
            Next
            .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
            If .Faccion.Status = e_Facciones.consejo Then
                Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageLocaleMsg(1812, .name & "¬" & chat, e_FontTypeNames.FONTTYPE_CONSEJO)) ' Msg1812=(Consejo) ¬1> ¬2
            ElseIf .Faccion.Status = e_Facciones.concilio Then
                Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageLocaleMsg(1813, .name & "¬" & chat, e_FontTypeNames.FONTTYPE_CONSEJOCAOS)) ' Msg1813=(Concilio) ¬1> ¬2)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCouncilMessage", Erl)
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleChangeDescription(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Description As String
        Description = reader.ReadString8()
        If .flags.Muerto = 1 Then
            ' Msg763=No podés cambiar la descripción estando muerto.
            Call WriteLocaleMsg(UserIndex, 763, e_FontTypeNames.FONTTYPE_INFOIAO)
        Else
            If Len(Description) > 128 Then
                ' Msg764=La descripción es muy larga.
                Call WriteLocaleMsg(UserIndex, 764, e_FontTypeNames.FONTTYPE_INFOIAO)
            ElseIf Not ValidDescription(Description) Then
                ' Msg765=La descripción tiene carácteres inválidos.
                Call WriteLocaleMsg(UserIndex, 765, e_FontTypeNames.FONTTYPE_INFOIAO)
            ElseIf Not ValidWordsDescription(Description) Then
                'Msg2000=La descripción contiene palabras que no están permitidas.
                Call WriteLocaleMsg(UserIndex, 2000, e_FontTypeNames.FONTTYPE_INFOIAO)
            Else
                .Desc = Trim$(Description)
                ' Msg766=La descripción a cambiado.
                Call WriteLocaleMsg(UserIndex, 766, e_FontTypeNames.FONTTYPE_INFOIAO)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeDescription", Erl)
End Sub

''
' Handles the "GuildVote" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildVote(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim vote     As String
        Dim errorStr As String
        vote = reader.ReadString8()
        'Msg1172= Elecciones de clan desactivadas por el momento.
        Call WriteLocaleMsg(UserIndex, 1172, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildVote", Erl)
End Sub

''
' Handles the "BankExtractGold" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)
    On Error GoTo HandleBankExtractGold_Err
    With UserList(UserIndex)
        Dim amount As Long
        amount = reader.ReadInt32()
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Validate target NPC
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            'Msg1173= Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 1173, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then Exit Sub
        If Distancia(.pos, NpcList(.flags.TargetNPC.ArrayIndex).pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If amount > 0 And amount <= .Stats.Banco Then
            .Stats.Banco = .Stats.Banco - amount
            .Stats.GLD = .Stats.GLD + amount
            Call WriteLocaleChatOverHead(UserIndex, 1418, .Stats.Banco, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite) ' Msg1418=Tenés ¬1 monedas de oro en tu cuenta.
            Call WriteUpdateGold(UserIndex)
            Call WriteUpdateBankGld(UserIndex)
        Else
            Call WriteLocaleChatOverHead(UserIndex, 1395, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1395=No tenés esa cantidad.
        End If
    End With
    Exit Sub
HandleBankExtractGold_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankExtractGold", Erl)
End Sub

Private Sub HandleLeaveFaction(ByVal UserIndex As Integer)
    On Error GoTo HandleLeaveFaction_Err
    With UserList(UserIndex)
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .Faccion.Status = e_Facciones.Ciudadano Then
            If .Faccion.Status = 1 Then
                Call VolverCriminal(UserIndex)
                'Msg1174= Ahora sos un criminal.
                Call WriteLocaleMsg(UserIndex, 1174, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        'Validate target NPC
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            If .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Then
                'Msg1175= Para salir del ejercito debes ir a visitar al rey.
                Call WriteLocaleMsg(UserIndex, 1175, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            ElseIf .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then
                'Msg1176= Para salir de la legion debes ir a visitar al diablo.
                Call WriteLocaleMsg(UserIndex, 1176, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Exit Sub
        End If
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType = e_NPCType.Enlistador Then
            'Quit the Royal Army?
            If .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Then
                If NpcList(.flags.TargetNPC.ArrayIndex).flags.Faccion = 0 Then
                    'Si tiene clan
                    If .GuildIndex > 0 Then
                        'Y no es leader
                        If Not PersonajeEsLeader(.Id) Then
                            'Me fijo de que alineación es el clan, si es ARMADA, lo hecho
                            If GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_ARMADA Then
                                Call m_EcharMiembroDeClan(UserIndex, .name)
                                'Msg1177= Has dejado el clan.
                                Call WriteLocaleMsg(UserIndex, 1177, e_FontTypeNames.FONTTYPE_INFO)
                            End If
                        Else
                            'Me fijo si está en un clan armada, en ese caso no lo dejo salir de la facción
                            If GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_ARMADA Then
                                Call WriteLocaleChatOverHead(UserIndex, 1396, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1396=Para dejar la facción primero deberás ceder el liderazgo del clan
                                Exit Sub
                            End If
                        End If
                    End If
                    Call ExpulsarFaccionReal(UserIndex)
                    Call WriteLocaleChatOverHead(UserIndex, 1397, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1397=Serás bienvenido a las fuerzas imperiales si deseas regresar.
                    Exit Sub
                Else
                    Call WriteLocaleChatOverHead(UserIndex, 1398, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1398=¡¡¡Sal de aquí bufón!!!
                End If
                'Quit the Chaos Legion??
            ElseIf .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then
                If NpcList(.flags.TargetNPC.ArrayIndex).flags.Faccion = 2 Then
                    'Si tiene clan
                    If .GuildIndex > 0 Then
                        'Y no es leader
                        If Not PersonajeEsLeader(.Id) Then
                            'Me fijo de que alineación es el clan, si es CAOS, lo hecho
                            If GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CAOTICA Then
                                Call m_EcharMiembroDeClan(UserIndex, .name)
                                'Msg1178= Has dejado el clan.
                                Call WriteLocaleMsg(UserIndex, 1178, e_FontTypeNames.FONTTYPE_INFO)
                            End If
                        Else
                            'Me fijo si está en un clan CAOS, en ese caso no lo dejo salir de la facción
                            If GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CAOTICA Then
                                Call WriteLocaleChatOverHead(UserIndex, 1399, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1399=Para dejar la facción primero deberás ceder el liderazgo del clan
                                Exit Sub
                            End If
                        End If
                    End If
                    Call ExpulsarFaccionCaos(UserIndex)
                    Call WriteLocaleChatOverHead(UserIndex, 1400, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1400=Ya volverás arrastrandote.
                Else
                    Call WriteLocaleChatOverHead(UserIndex, 1401, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1401=Sal de aquí maldito criminal
                End If
            Else
                Call WriteLocaleChatOverHead(UserIndex, 1402, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1402=¡No perteneces a ninguna facción!
            End If
        End If
    End With
    Exit Sub
HandleLeaveFaction_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleLeaveFaction", Erl)
End Sub

Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)
    On Error GoTo HandleBankDepositGold_Err
    With UserList(UserIndex)
        Dim amount As Long
        amount = reader.ReadInt32()
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Validate target NPC
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            'Msg1179= Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 1179, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then Exit Sub
        If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If amount > 0 And amount <= .Stats.GLD Then
            'substract first in case there is overflow we don't dup gold
            .Stats.GLD = .Stats.GLD - amount
            .Stats.Banco = .Stats.Banco + amount
            Call WriteLocaleChatOverHead(UserIndex, 1418, .Stats.Banco, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite) ' Msg1418=Tenés ¬1 monedas de oro en tu cuenta.
            Call WriteUpdateGold(UserIndex)
            Call WriteUpdateBankGld(UserIndex)
        Else
            Call WriteLocaleChatOverHead(UserIndex, 1419, "", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite) ' Msg1419=No tenés esa cantidad.
        End If
    End With
    Exit Sub
HandleBankDepositGold_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankDepositGold", Erl)
End Sub

' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildMemberList(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim guild       As String
        Dim memberCount As Integer
        Dim i           As Long
        Dim username    As String
        guild = reader.ReadString8()
        If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios) Then
            If (InStrB(guild, "\") <> 0) Then
                guild = Replace(guild, "\", "")
            End If
            If (InStrB(guild, "/") <> 0) Then
                guild = Replace(guild, "/", "")
            End If
            If Not modGuilds.YaExiste(guild) Then
                'Msg1180= No existe el clan: ¬1
                Call WriteLocaleMsg(UserIndex, 1180, e_FontTypeNames.FONTTYPE_INFO, guild)
                Exit Sub
            End If
            Dim MembersId() As Long
            MembersId = GetGuildMemberList(guild)
            For i = LBound(MembersId) To UBound(MembersId)
                Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1993, GetUserName(MembersId(i)) & "¬" & guild, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1993=¬1 <¬2>
            Next i
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildMemberList", Erl)
End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleOnlineRoyalArmy(ByVal UserIndex As Integer)
    On Error GoTo HandleOnlineRoyalArmy_Err
    With UserList(UserIndex)
        If .flags.Privilegios And e_PlayerType.User Then Exit Sub
        Dim i    As Long
        Dim list As String
        For i = 1 To LastUser
            If UserList(i).ConnectionDetails.ConnIDValida Then
                If UserList(i).Faccion.Status = e_Facciones.Armada Or UserList(i).Faccion.Status = e_Facciones.consejo Then
                    If UserList(i).flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios) Or .flags.Privilegios And (e_PlayerType.Dios Or _
                            e_PlayerType.Admin) Then
                        list = list & UserList(i).name & ", "
                    End If
                End If
            End If
        Next i
    End With
    If Len(list) > 0 Then
        'Msg1289= Armadas conectados: ¬1
        Call WriteLocaleMsg(UserIndex, 1289, e_FontTypeNames.FONTTYPE_INFO, Left$(list, Len(list) - 2))
    Else
        'Msg1182= No hay Armadas conectados
        Call WriteLocaleMsg(UserIndex, 1182, e_FontTypeNames.FONTTYPE_INFO)
    End If
    Exit Sub
HandleOnlineRoyalArmy_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnlineRoyalArmy", Erl)
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)
    On Error GoTo HandleOnlineChaosLegion_Err
    With UserList(UserIndex)
        If .flags.Privilegios And e_PlayerType.User Then Exit Sub
        Dim i    As Long
        Dim list As String
        For i = 1 To LastUser
            If UserList(i).ConnectionDetails.ConnIDValida Then
                If UserList(i).Faccion.Status = e_Facciones.Caos Or UserList(i).Faccion.Status = e_Facciones.concilio Then
                    If UserList(i).flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios) Or .flags.Privilegios And (e_PlayerType.Dios Or _
                            e_PlayerType.Admin) Then
                        list = list & UserList(i).name & ", "
                    End If
                End If
            End If
        Next i
    End With
    If Len(list) > 0 Then
        'Msg1290= Caos conectados: ¬1
        Call WriteLocaleMsg(UserIndex, 1290, e_FontTypeNames.FONTTYPE_INFO, Left$(list, Len(list) - 2))
    Else
        'Msg1184= No hay Caos conectados
        Call WriteLocaleMsg(UserIndex, 1184, e_FontTypeNames.FONTTYPE_INFO)
    End If
    Exit Sub
HandleOnlineChaosLegion_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnlineChaosLegion", Erl)
End Sub

''
' Handles the "Comment" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleComment(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim comment As String
        comment = reader.ReadString8()
        If Not .flags.Privilegios And e_PlayerType.User Then
            Call LogGM(.name, "Comentario: " & comment)
            'Msg1185= Comentario salvado...
            Call WriteLocaleMsg(UserIndex, 1185, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleComment", Erl)
End Sub

Private Sub HandleUseKey(ByVal UserIndex As Integer)
    On Error GoTo HandleUseKey_Err
    With UserList(UserIndex)
        Dim Slot As Byte
        Slot = reader.ReadInt8
        Call UsarLlave(UserIndex, Slot)
    End With
    Exit Sub
HandleUseKey_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleUseKey", Erl)
End Sub

Private Sub HandleMensajeUser(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim mensaje  As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        mensaje = reader.ReadString8()
        If EsGM(UserIndex) Then
            If LenB(username) = 0 Or LenB(mensaje) = 0 Then
                'Msg1186= Utilice /MENSAJEINFORMACION nick@mensaje
                Call WriteLocaleMsg(UserIndex, 1186, e_FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(username)
                If IsValidUserRef(tUser) Then
                    'Msg1187= Mensaje recibido de ¬1
                    Call WriteLocaleMsg(tUser.ArrayIndex, 1187, e_FontTypeNames.FONTTYPE_INFO, .name)
                    Call WriteConsoleMsg(tUser.ArrayIndex, mensaje, e_FontTypeNames.FONTTYPE_New_DONADOR)
                Else
                    If PersonajeExiste(username) Then
                        Call SetMessageInfoDatabase(username, "Mensaje recibido de " & .name & " [Game Master]: " & vbNewLine & mensaje & vbNewLine)
                    End If
                End If
                'Msg1188= Mensaje enviado a ¬1
                Call WriteLocaleMsg(UserIndex, 1188, e_FontTypeNames.FONTTYPE_INFO, username)
                Call LogGM(.name, "Envió mensaje como GM a " & username & ": " & mensaje)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleMensajeUser", Erl)
End Sub

Private Sub HandleTraerBoveda(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Call UpdateUserHechizos(True, UserIndex, 0)
        Call UpdateUserInv(True, UserIndex, 0)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleTraerBoveda", Erl)
End Sub


Private Sub HandleNotifyInventariohechizos(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim value     As Byte
        Dim hechiSel  As Byte
        Dim scrollSel As Byte
        value = reader.ReadInt8()
        hechiSel = reader.ReadInt8()
        scrollSel = reader.ReadInt8()
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleReviveChar", Erl)
End Sub


Private Sub HandlePerdonFaccion(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        Dim LoopC    As Byte
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            If UCase$(username) <> "YO" Then
                tUser = NameIndex(username)
            Else
                Call SetUserRef(tUser, UserIndex)
            End If
            If Not IsValidUserRef(tUser) Then
                ' Msg743=Usuario offline.
                Call WriteLocaleMsg(UserIndex, 743, e_FontTypeNames.FONTTYPE_INFO)
            End If
            If UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.Armada Or UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.Caos Or UserList( _
                    tUser.ArrayIndex).Faccion.Status = e_Facciones.consejo Or UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.concilio Then
                'Msg1189= No puedes perdonar a alguien que ya pertenece a una facción
                Call WriteLocaleMsg(UserIndex, 1189, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'Si es ciudadano aparte de quitarle las reenlistadas le saco los ciudadanos matados.
            If UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.Ciudadano Then
                If UserList(tUser.ArrayIndex).Faccion.ciudadanosMatados > 0 Or UserList(tUser.ArrayIndex).Faccion.Reenlistadas > 0 Then
                    UserList(tUser.ArrayIndex).Faccion.ciudadanosMatados = 0
                    UserList(tUser.ArrayIndex).Faccion.Reenlistadas = 0
                    UserList(tUser.ArrayIndex).Faccion.RecibioArmaduraReal = 0
                    'Msg1190= Has sido perdonado.
                    Call WriteLocaleMsg(tUser.ArrayIndex, 1190, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg1191= Has perdonado a ¬1
                    Call WriteLocaleMsg(UserIndex, 1191, e_FontTypeNames.FONTTYPE_INFO, UserList(tUser.ArrayIndex).name)
                Else
                    'Msg1192= No necesitas ser perdonado.
                    Call WriteLocaleMsg(tUser.ArrayIndex, 1192, e_FontTypeNames.FONTTYPE_INFO)
                End If
            ElseIf UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.Criminal Then
                If UserList(tUser.ArrayIndex).Faccion.Reenlistadas = 0 Then
                    'Msg1193= No necesitas ser perdonado.
                    Call WriteLocaleMsg(tUser.ArrayIndex, 1193, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                Else
                    UserList(tUser.ArrayIndex).Faccion.Reenlistadas = 0
                    UserList(tUser.ArrayIndex).Faccion.RecibioArmaduraCaos = 0
                    'Msg1194= Has sido perdonado.
                    Call WriteLocaleMsg(tUser.ArrayIndex, 1194, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg1195= Has perdonado a ¬1
                    Call WriteLocaleMsg(UserIndex, 1195, e_FontTypeNames.FONTTYPE_INFO, UserList(tUser.ArrayIndex).name)
                End If
            End If
        Else
            Call WriteLocaleMsg(UserIndex, 528, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandlePerdonFaccion", Erl)
End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildOnlineMembers(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim GuildName As String
        Dim tGuild    As Integer
        GuildName = reader.ReadString8()
        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")
        End If
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            tGuild = GuildIndex(GuildName)
            If tGuild > 0 Then
                'Msg1196= Clan ¬1
                Call WriteLocaleMsg(UserIndex, 1196, e_FontTypeNames.FONTTYPE_INFO, UCase$(GuildName))
            End If
        Else
            Call WriteLocaleMsg(UserIndex, 528, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOnlineMembers", Erl)
End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleRoyalArmyMessage(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Message As String
        Message = reader.ReadString8()
        'Solo dioses, admins y RMS
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Or .Faccion.Status = e_Facciones.consejo Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageLocaleMsg(1815, UserList(UserIndex).name & "¬" & Message, e_FontTypeNames.FONTTYPE_CONSEJO)) ' Msg1815=[ARMADA REAL] ¬1> ¬2
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRoyalArmyMessage", Erl)
End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleChaosLegionMessage(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Message As String
        Message = reader.ReadString8()
        'Solo dioses, admins y RMS
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Or .Faccion.Status = e_Facciones.concilio Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageLocaleMsg(1816, UserList(UserIndex).name & "¬" & Message, e_FontTypeNames.FONTTYPE_CONSEJOCAOS)) ' Msg1816=[FUERZAS DEL CAOS] ¬1> ¬2
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChaosLegionMessage", Erl)
End Sub

Private Sub HandleFactionMessage(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim currentTime  As Long
    Dim elapsedMs    As Double
    Dim Message      As String
    Dim factionLabel As String
    Dim fontType     As e_FontTypeNames
    Dim Target       As Byte
    With UserList(UserIndex)
        Message = reader.ReadString8()
        If LenB(Message) = 0 Then Exit Sub
        currentTime = GetTickCountRaw()
        elapsedMs = TicksElapsed(.Counters.MensajeGlobal, currentTime)
        'Si esta silenciado no le deja enviar mensaje
        If .flags.Silenciado = 1 Then
            Call WriteLocaleMsg(UserIndex, 110, e_FontTypeNames.FONTTYPE_VENENO, .flags.MinutosRestantes)
            Exit Sub
        End If
        'Previene spam de mensajes globales
        If elapsedMs < IntervaloMensajeGlobal Then
            ' Msg548=No puedes escribir mensajes globales tan rápido.
            Call WriteLocaleMsg(UserIndex, 548, e_FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        'Actualiza el tiempo del último mensaje
        .Counters.MensajeGlobal = currentTime
        'Determina la etiqueta y estilo según la facción
        Select Case .Faccion.Status
            Case e_Facciones.consejo
                factionLabel = "MENSAJE_CONSEJO"
                fontType = e_FontTypeNames.FONTTYPE_CONSEJO
                Target = SendTarget.ToRealYRMs
            Case e_Facciones.Armada
                factionLabel = "MENSAJE_ARMADA"
                fontType = e_FontTypeNames.FONTTYPE_CITIZEN_ARMADA
                Target = SendTarget.ToRealYRMs
            Case e_Facciones.concilio
                factionLabel = "MENSAJE_CONCILIO"
                fontType = e_FontTypeNames.FONTTYPE_CONSEJOCAOS
                Target = SendTarget.ToCaosYRMs
            Case e_Facciones.Caos
                factionLabel = "MENSAJE_LEGION"
                fontType = e_FontTypeNames.FONTTYPE_CRIMINAL_CAOS
                Target = SendTarget.ToCaosYRMs
            Case Else
                Exit Sub 'Si no pertenece a ninguna facción válida
        End Select
        'Envía el mensaje de facción
        Dim formattedMessage As String
        formattedMessage = " " & .name & "> " & Message
        Call SendData(Target, 0, PrepareFactionMessageConsole(factionLabel, formattedMessage, fontType))
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleFactionMessage", Erl)
End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleAcceptRoyalCouncilMember(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        Dim LoopC    As Byte
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            tUser = NameIndex(username)
            If Not IsValidUserRef(tUser) Then
                'Msg1197= Usuario offline
                Call WriteLocaleMsg(UserIndex, 1197, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                If UserList(tUser.ArrayIndex).GuildIndex > 0 Then
                    If GuildAlignmentIndex(UserList(tUser.ArrayIndex).GuildIndex) <> e_ALINEACION_GUILD.ALINEACION_ARMADA Then
                        'Msg1198= El miembro no puede ingresar al consejo porque forma parte de un clan que no es de la armada.
                        Call WriteLocaleMsg(UserIndex, 1198, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1643, username, e_FontTypeNames.FONTTYPE_CONSEJO)) 'Msg1643=¬1 fue aceptado en el honorable Consejo Real de Banderbill.
                With UserList(tUser.ArrayIndex)
                    .Faccion.Status = e_Facciones.consejo
                    Call WarpUserChar(tUser.ArrayIndex, .pos.Map, .pos.x, .pos.y, False)
                End With
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleAcceptRoyalCouncilMember", Erl)
End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleAcceptChaosCouncilMember(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        Dim LoopC    As Byte
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            tUser = NameIndex(username)
            If Not IsValidUserRef(tUser) Then
                'Msg1199= Usuario offline
                Call WriteLocaleMsg(UserIndex, 1199, e_FontTypeNames.FONTTYPE_INFO)
            Else
                If UserList(tUser.ArrayIndex).GuildIndex > 0 Then
                    If GuildAlignmentIndex(UserList(tUser.ArrayIndex).GuildIndex) <> e_ALINEACION_GUILD.ALINEACION_CAOTICA Then
                        'Msg1200= El miembro no puede ingresar al concilio porque forma parte de un clan que no es caótico.
                        Call WriteLocaleMsg(UserIndex, 1200, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1644, username, e_FontTypeNames.FONTTYPE_CONSEJOCAOS)) 'Msg1644=¬1 fue aceptado en el Consejo de la Legión Oscura.
                With UserList(tUser.ArrayIndex)
                    .Faccion.Status = e_Facciones.concilio
                    Call WarpUserChar(tUser.ArrayIndex, .pos.Map, .pos.x, .pos.y, False)
                End With
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleAcceptChaosCouncilMember", Erl)
End Sub

''
' Handles the "CouncilKick" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleCouncilKick(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            tUser = NameIndex(username)
            If Not IsValidUserRef(tUser) Then
                If PersonajeExiste(username) Then
                    'Msg1201= Usuario offline, echando de los consejos
                    Call WriteLocaleMsg(UserIndex, 1201, e_FontTypeNames.FONTTYPE_INFO)
                    Dim Status As Integer
                    Status = GetDBValue("user", "status", "name", username)
                    Call EcharConsejoDatabase(username, IIf(Status = 4, 2, 3))
                    'Msg1202= Usuario ¬1
                    Call WriteLocaleMsg(UserIndex, 1202, e_FontTypeNames.FONTTYPE_INFO, username)
                Else
                    'Msg1203= No existe el personaje.
                    Call WriteLocaleMsg(UserIndex, 1203, e_FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                With UserList(tUser.ArrayIndex)
                    If .Faccion.Status = e_Facciones.consejo Then
                        'Msg1204= Has sido echado del consejo de Banderbill
                        Call WriteLocaleMsg(tUser.ArrayIndex, 1204, e_FontTypeNames.FONTTYPE_INFO)
                        .Faccion.Status = e_Facciones.Armada
                        Call WarpUserChar(tUser.ArrayIndex, .pos.Map, .pos.x, .pos.y)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1645, username, e_FontTypeNames.FONTTYPE_CONSEJO)) 'Msg1645=¬1 fue expulsado del Consejo Real de Banderbill.
                    End If
                    If .Faccion.Status = e_Facciones.concilio Then
                        'Msg1205= Has sido echado del consejo de la Legión Oscura
                        Call WriteLocaleMsg(tUser.ArrayIndex, 1205, e_FontTypeNames.FONTTYPE_INFO)
                        .Faccion.Status = e_Facciones.Caos
                        Call WarpUserChar(tUser.ArrayIndex, .pos.Map, .pos.x, .pos.y)
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1646, username, e_FontTypeNames.FONTTYPE_CONSEJOCAOS)) 'Msg1646=¬1 fue expulsado del Consejo de la Legión Oscura.
                    End If
                    Call RefreshCharStatus(tUser.ArrayIndex)
                End With
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCouncilKick", Erl)
End Sub

''
' Handles the "GuildBan" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildBan(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim GuildName   As String
        Dim cantMembers As Integer
        Dim LoopC       As Long
        Dim member      As String
        Dim count       As Byte
        Dim tUser       As t_UserReference
        Dim tFile       As String
        GuildName = reader.ReadString8()
        If (Not .flags.Privilegios And e_PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            If Not FileExist(tFile) Then
                'Msg1206= No existe el clan: ¬1
                Call WriteLocaleMsg(UserIndex, 1206, e_FontTypeNames.FONTTYPE_INFO, GuildName)
            Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1647, .name & "¬" & UCase$(GuildName), e_FontTypeNames.FONTTYPE_FIGHT)) 'Msg1647=¬1 banned al clan ¬2.
                'baneamos a los miembros
                Call LogGM(.name, "BANCLAN a " & UCase$(GuildName))
                cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
                For LoopC = 1 To cantMembers
                    'member es la victima
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    Call Ban(member, "Administracion del servidor", "Clan Banned")
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1648, member & "¬" & GuildName, e_FontTypeNames.FONTTYPE_FIGHT)) 'Msg1648=¬1<¬2> ha sido expulsado del servidor.
                    tUser = NameIndex(member)
                    If IsValidUserRef(tUser) Then
                        'esta online
                        UserList(tUser.ArrayIndex).flags.Ban = 1
                        Call CloseSocket(tUser.ArrayIndex)
                    End If
                    Call SaveBanDatabase(member, .name & " - BAN AL CLAN: " & GuildName & ". " & Date & " " & Time, .name)
                Next LoopC
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildBan", Erl)
End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleChaosLegionKick(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "\", "")
            End If
            If (InStrB(username, "/") <> 0) Then
                username = Replace(username, "/", "")
            End If
            tUser = NameIndex(username)
            Call LogGM(.name, "ECHO DEL CAOS A: " & username)
            If IsValidUserRef(tUser) Then
                If UserList(tUser.ArrayIndex).GuildIndex > 0 Then
                    'Me fijo de que alineación es el clan, si es Legion, lo echo
                    If GuildAlignmentIndex(UserList(tUser.ArrayIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CAOTICA Then
                        Call m_EcharMiembroDeClan(UserIndex, UserList(tUser.ArrayIndex).Id)
                    End If
                End If
                    UserList(tUser.ArrayIndex).Faccion.Reenlistadas = 2
                    UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.Criminal
                    Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1992, username, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1992=¬1 expulsado de las fuerzas del caos y prohibida la reenlistada.
                    Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1991, .name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1991=¬1 te ha expulsado en forma definitiva de las fuerzas del caos.
            Else
                If PersonajeExiste(username) Then
                    'Msg1208= Usuario offline, echando de la facción
                    Call WriteLocaleMsg(UserIndex, 1208, e_FontTypeNames.FONTTYPE_INFO)
                    Dim Status As Integer
                    Status = GetDBValue("user", "status", "name", username)
                    If Status = e_Facciones.Caos Then
                        Call EcharLegionDatabase(username)
                        'Msg1209= Usuario ¬1
                        Call WriteLocaleMsg(UserIndex, 1209, e_FontTypeNames.FONTTYPE_INFO, username)
                    Else
                        'Msg1210= El personaje no pertenece a la legión.
                        Call WriteLocaleMsg(UserIndex, 1210, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    'Msg1211= No existe el personaje.
                    Call WriteLocaleMsg(UserIndex, 1211, e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChaosLegionKick", Erl)
End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    UserIndex The index of the user sending the message.
Private Sub HandleRoyalArmyKick(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()

        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "\", "")
            End If
            If (InStrB(username, "/") <> 0) Then
                username = Replace(username, "/", "")
            End If
            tUser = NameIndex(username)
            Call LogGM(.name, "ECHO DE LA REAL A: " & username)
            If IsValidUserRef(tUser) Then
                If UserList(tUser.ArrayIndex).GuildIndex > 0 Then
                    'Me fijo de que alineación es el clan, si es ARMADA, lo echo
                    If GuildAlignmentIndex(UserList(tUser.ArrayIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_ARMADA Then
                        Call m_EcharMiembroDeClan(UserIndex, UserList(tUser.ArrayIndex).Id)
                    End If
                End If
                UserList(tUser.ArrayIndex).Faccion.Reenlistadas = 2
                UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.Ciudadano
                Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1990, username, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1990=¬1 expulsado de las fuerzas reales y prohibida la reenlistada.
                Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1989, .name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1989=¬1 te ha expulsado en forma definitiva de las fuerzas reales.
            Else
                If PersonajeExiste(username) Then
                    'Msg1213= Usuario offline, echando de la facción
                    Call WriteLocaleMsg(UserIndex, 1213, e_FontTypeNames.FONTTYPE_INFO)
                    Dim Status As Integer
                    Status = GetDBValue("user", "status", "name", username)
                    If Status = e_Facciones.Armada Then
                        Call EcharArmadaDatabase(username)
                        'Msg1214= Usuario ¬1
                        Call WriteLocaleMsg(UserIndex, 1214, e_FontTypeNames.FONTTYPE_INFO, username)
                    Else
                        'Msg1215= El personaje no pertenece a la armada.
                        Call WriteLocaleMsg(UserIndex, 1215, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    'Msg1216= No existe el personaje.
                End If
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRoyalArmyKick", Erl)
End Sub

''
' Handles the "ChatColor" message.
'
' @param    UserIndex The index of the user sending the message.
Public Sub HandleChatColor(ByVal UserIndex As Integer)
    On Error GoTo HandleChatColor_Err
    'Change the user`s chat color
    With UserList(UserIndex)
        Dim Color As Long
        Color = RGB(reader.ReadInt8(), reader.ReadInt8(), reader.ReadInt8())
        If EsGM(UserIndex) Then
            .flags.ChatColor = Color
        End If
    End With
    Exit Sub
HandleChatColor_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChatColor", Erl)
End Sub

Public Sub HandleDonateGold(ByVal UserIndex As Integer)
    On Error GoTo handle
    With UserList(UserIndex)
        Dim Oro As Long
        Oro = reader.ReadInt32
        If Oro <= 0 Then Exit Sub
        'Se asegura que el target es un npc
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            'Msg1217= Primero tenés que seleccionar al sacerdote.
            Call WriteLocaleMsg(UserIndex, 1217, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim priest As t_Npc
        priest = NpcList(.flags.TargetNPC.ArrayIndex)
        'Validate NPC is an actual priest and the player is not dead
        If (priest.npcType <> e_NPCType.Revividor And (priest.npcType <> e_NPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 1 Then Exit Sub
        'Make sure it's close enough
        If Distancia(.pos, NpcList(.flags.TargetNPC.ArrayIndex).pos) > 3 Then
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .Faccion.Status = e_Facciones.Ciudadano Or .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Or .Faccion.Status = e_Facciones.concilio Or _
                .Faccion.Status = e_Facciones.Caos Then
            Call WriteLocaleChatOverHead(UserIndex, 1377, "", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite) ' Msg1377=No puedo aceptar tu donación en este momento...
            Exit Sub
        End If
        If .GuildIndex <> 0 Then
            If modGuilds.Alineacion(.GuildIndex) = 1 Then
                Call WriteLocaleChatOverHead(UserIndex, 1404, vbNullString, priest.Char.charindex, vbWhite)  ' Msg1404=Te encuentras en un clan criminal... no puedo aceptar tu donación.
                Exit Sub
            End If
        End If
        If .Stats.GLD < Oro Then
            'Msg1218= No tienes suficiente dinero.
            Call WriteLocaleMsg(UserIndex, 1218, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim Donacion As Long
        If .Faccion.ciudadanosMatados > 0 Then
            Donacion = .Faccion.ciudadanosMatados * SvrConfig.GetValue("GoldMult") * SvrConfig.GetValue("CostoPerdonPorCiudadano")
        Else
            Donacion = SvrConfig.GetValue("CostoPerdonPorCiudadano") / 2
        End If
        If Oro < Donacion Then
            Call WriteLocaleChatOverHead(UserIndex, 1405, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1405=Dios no puede perdonarte si eres una persona avara.
            Exit Sub
        End If
        .Stats.GLD = .Stats.GLD - Oro
        Call WriteUpdateGold(UserIndex)
        'Msg1219= Has donado ¬1
        Call WriteLocaleMsg(UserIndex, 1219, e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(Oro))
        Call WriteLocaleChatOverHead(UserIndex, 1406, vbNullString, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, vbYellow)  ' Msg1406=¡Gracias por tu generosa donación! Con estas palabras, te libero de todo tipo de pecados. ¡Que Dios te acompañe hijo mío!
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, 80, 100, False))
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(100, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        Call VolverCiudadano(UserIndex)
    End With
    Exit Sub
handle:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleDonateGold", Erl)
End Sub

Public Sub HandlePromedio(ByVal UserIndex As Integer)
    On Error GoTo handle
    With UserList(UserIndex)
        Call WriteLocaleMsg(UserIndex, 1988, e_FontTypeNames.FONTTYPE_INFOBOLD, .clase & "¬" & .raza & "¬" & .Stats.ELV)
        Dim Promedio As Double, Vida As Long
        Promedio = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(e_Atributos.Constitucion)) * 0.5
        Vida = 18 + ModRaza(.raza).Constitucion + Promedio * (.Stats.ELV - 1)
        'Msg1220= Vida esperada: ¬1
        Call WriteLocaleMsg(UserIndex, 1220, e_FontTypeNames.FONTTYPE_INFOBOLD, Vida & "¬" & Promedio)
        Promedio = CalcularPromedioVida(UserIndex)
        Dim Diff As Long, Color As e_FontTypeNames, Signo As String
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
        'Msg1221= Vida actual: ¬1
        Call WriteLocaleMsg(UserIndex, 1221, Color, .Stats.MaxHp & " (" & Signo & Abs(Diff) & ")" & "¬" & Round(Promedio, 2))
    End With
    Exit Sub
handle:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandlePromedio", Erl)
End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param UserIndex The index of the user sending the message
Public Sub HandleShowGuildMessages(ByVal UserIndex As Integer)
    'Allows admins to read guild messages
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim guild As String
        guild = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(UserIndex, guild)
            Call LogGM(.name, .name & " espia a " & guild)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

''
' Handle the "DoBackUp" message
'
' @param UserIndex The index of the user sending the message
Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
    On Error GoTo HandleDoBackUp_Err
    'Show guilds messages
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        Call LogGM(.name, .name & " ha hecho un backup")
        Call ES.DoBackUp 'Sino lo confunde con la id del paquete
    End With
    Exit Sub
HandleDoBackUp_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleDoBackUp", Erl)
End Sub

''
' Handle the "NavigateToggle" message
'
' @param UserIndex The index of the user sending the message
Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)
    On Error GoTo HandleNavigateToggle_Err
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, 528, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1
        End If
        'Tell the client that we are navigating.
        Call WriteNavigateToggle(UserIndex, .flags.Navegando)
    End With
    Exit Sub
HandleNavigateToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleNavigateToggle", Erl)
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param UserIndex The index of the user sending the message
Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)
    On Error GoTo HandleServerOpenToUsersToggle_Err
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        If ServerSoloGMs > 0 Then
            'Msg1222= Servidor habilitado para todos.
            Call WriteLocaleMsg(UserIndex, 1222, e_FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
        Else
            'Msg1223= Servidor restringido a administradores.
            Call WriteLocaleMsg(UserIndex, 1223, e_FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
        End If
    End With
    Exit Sub
HandleServerOpenToUsersToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleServerOpenToUsersToggle", Erl)
End Sub

Public Sub HandleParticipar(ByVal UserIndex As Integer)
    On Error GoTo HandleParticipar_Err
    Dim handle   As Integer
    Dim RoomId   As Integer
    Dim Password As String
    RoomId = reader.ReadInt16
    Password = reader.ReadString8
    With UserList(UserIndex)
        If RoomId = -1 Then
            If CurrentActiveEventType = CaptureTheFlag Then
                If Not InstanciaCaptura Is Nothing Then
                    Call InstanciaCaptura.inscribirse(UserIndex)
                    Exit Sub
                End If
            Else
                RoomId = GlobalLobbyIndex
            End If
        End If
        If LobbyList(RoomId).State = AcceptingPlayers Then
            If LobbyList(RoomId).IsPublic Then
                Dim addPlayerResult As t_response
                addPlayerResult = ModLobby.AddPlayerOrGroup(LobbyList(RoomId), UserIndex, Password)
                Call WriteLocaleMsg(UserIndex, addPlayerResult.Message, e_FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteLocaleMsg(UserIndex, MsgCantJoinPrivateLobby, e_FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If
    End With
    Exit Sub
HandleParticipar_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleParticipar", Erl)
End Sub

''
' Handle the "ResetFactions" message
'
' @param UserIndex The index of the user sending the message
Public Sub HandleResetFactions(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            Call LogGM(.name, "/RAJAR " & username)
            tUser = NameIndex(username)
            If IsValidUserRef(tUser) Then Call ResetFacciones(tUser.ArrayIndex)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleResetFactions", Erl)
End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param UserIndex The index of the user sending the message
Public Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username   As String
        Dim GuildIndex As Integer
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
            Call LogGM(.name, "/RAJARCLAN " & username)
            GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, username)
            If GuildIndex = 0 Then
                'Msg1224= No pertenece a ningún clan o es fundador.
                Call WriteLocaleMsg(UserIndex, 1224, e_FontTypeNames.FONTTYPE_INFO)
            Else
                'Msg1225= Expulsado.
                Call WriteLocaleMsg(UserIndex, 1225, e_FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageLocaleMsg(1817, username, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1817=¬1 ha sido expulsado del clan por los administradores del servidor.
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRemoveCharFromGuild", Erl)
End Sub

''
' Handle the "SystemMessage" message
'
' @param UserIndex The index of the user sending the message
Public Sub HandleSystemMessage(ByVal UserIndex As Integer)
    'Send a message to all the users
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Message As String
        Message = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
            Call LogGM(.name, "Mensaje de sistema:" & Message)
            Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(Message))
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSystemMessage", Erl)
End Sub

Private Sub HandleOfertaInicial(ByVal UserIndex As Integer)
    On Error GoTo HandleOfertaInicial_Err
    With UserList(UserIndex)
        Dim Oferta As Long
        Oferta = reader.ReadInt32()
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            'Msg1226= Primero tenés que hacer click sobre el subastador.
            Call WriteLocaleMsg(UserIndex, 1226, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Subastador Then
            'Msg1227= Primero tenés que hacer click sobre el subastador.
            Call WriteLocaleMsg(UserIndex, 1227, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 2 Then
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.Subastando = False Then
            Call WriteLocaleChatOverHead(UserIndex, 1407, vbNullString, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1407=Oye amigo, tu no podés decirme cual es la oferta inicial.
            Exit Sub
        End If
        If Subasta.HaySubastaActiva = False And .flags.Subastando = False Then
            'Msg1228= No hay ninguna subasta en curso.
            Call WriteLocaleMsg(UserIndex, 1228, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.Subastando = True Then
            UserList(UserIndex).Counters.TiempoParaSubastar = 0
            Subasta.OfertaInicial = Oferta
            Subasta.MejorOferta = 0
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1649, .name & "¬" & ObjData(Subasta.ObjSubastado).name & "¬" & Subasta.ObjSubastadoCantidad & "¬" & _
                    PonerPuntos(Subasta.OfertaInicial), e_FontTypeNames.FONTTYPE_SUBASTA)) 'Msg1649=¬1 está subastando: ¬2 (Cantidad: ¬3 ) - con un precio inicial de ¬4 monedas. Escribe /OFERTAR (cantidad) para participar.
            .flags.Subastando = False
            Subasta.HaySubastaActiva = True
            Subasta.Subastador = .name
            Subasta.MinutosDeSubasta = 5
            Subasta.TiempoRestanteSubasta = 300
            Call LogearEventoDeSubasta( _
                    "#################################################################################################################################################################################################")
            Call LogearEventoDeSubasta("El dia: " & Date & " a las " & Time)
            Call LogearEventoDeSubasta(.name & ": Esta subastando el item numero " & Subasta.ObjSubastado & " con una cantidad de " & Subasta.ObjSubastadoCantidad & _
                    " y con un precio inicial de " & PonerPuntos(Subasta.OfertaInicial) & " monedas.")
            frmMain.SubastaTimer.Enabled = True
            Call WarpUserChar(UserIndex, 14, 27, 64, True)
            'lalala toda la bola de los timerrr
        End If
    End With
    Exit Sub
HandleOfertaInicial_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleOfertaInicial", Erl)
End Sub

Private Sub HandleOfertaDeSubasta(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Oferta   As Long
        Dim ExOferta As t_UserReference
        Oferta = reader.ReadInt32()
        If Subasta.HaySubastaActiva = False Then
            'Msg1229= No hay ninguna subasta en curso.
            Call WriteLocaleMsg(UserIndex, 1229, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Oferta < Subasta.MejorOferta + 100 Then
            'Msg1230= Debe haber almenos una diferencia de 100 monedas a la ultima oferta!
            Call WriteLocaleMsg(UserIndex, 1230, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .name = Subasta.Subastador Then
            'Msg1231= No podés auto ofertar en tus subastas. La proxima vez iras a la carcel...
            Call WriteLocaleMsg(UserIndex, 1231, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .Stats.GLD >= Oferta Then
            'revisar que pasa si el usuario que oferto antes esta offline
            'Devolvemos el oro al usuario que oferto antes...(si es que hubo oferta)
            If Subasta.HuboOferta = True Then
                ExOferta = NameIndex(Subasta.Comprador)
                UserList(ExOferta.ArrayIndex).Stats.GLD = UserList(ExOferta.ArrayIndex).Stats.GLD + Subasta.MejorOferta
                Call WriteUpdateGold(ExOferta.ArrayIndex)
            End If
            Subasta.MejorOferta = Oferta
            Subasta.Comprador = .name
            .Stats.GLD = .Stats.GLD - Oferta
            Call WriteUpdateGold(UserIndex)
            If Subasta.TiempoRestanteSubasta < 60 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1650, .name & "¬" & PonerPuntos(Oferta) & "¬", e_FontTypeNames.FONTTYPE_SUBASTA)) 'Msg1650=Oferta mejorada por: ¬1 (Ofrece ¬2 monedas de oro) - Tiempo Extendido. Escribe /SUBASTA para mas información.
                Call LogearEventoDeSubasta(.name & ": Mejoro la oferta en el ultimo minuto ofreciendo " & PonerPuntos(Oferta) & " monedas.")
                Subasta.TiempoRestanteSubasta = Subasta.TiempoRestanteSubasta + 30
            Else
                Call LogearEventoDeSubasta(.name & ": Mejoro la oferta ofreciendo " & PonerPuntos(Oferta) & " monedas.")
                Subasta.HuboOferta = True
                Subasta.PosibleCancelo = False
            End If
        Else
            'Msg1232= No posees esa cantidad de oro.
            Call WriteLocaleMsg(UserIndex, 1232, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Public Sub HandleDuel(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim Players         As String
    Dim Bet             As Long
    Dim PocionesMaximas As Integer
    Dim CaenItems       As Boolean
    With UserList(UserIndex)
        Players = reader.ReadString8
        Bet = reader.ReadInt32
        PocionesMaximas = reader.ReadInt16
        CaenItems = reader.ReadBool
        'Msg1233= No puedes realizar un reto en este momento.
        Call WriteLocaleMsg(UserIndex, 1233, e_FontTypeNames.FONTTYPE_INFO)
        'Exit Sub
        Call CrearReto(UserIndex, Players, Bet, PocionesMaximas, CaenItems)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleDuel", Erl)
End Sub

Private Sub HandleAcceptDuel(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim Offerer As String
    With UserList(UserIndex)
        Offerer = reader.ReadString8
        Call AceptarReto(UserIndex, Offerer)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleAcceptDuel", Erl)
End Sub

Private Sub HandleCancelDuel(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        reader.ReadInt16
        If .flags.SolicitudReto.Estado <> e_SolicitudRetoEstado.Libre Then
            Call CancelarSolicitudReto(UserIndex, .name & " ha cancelado la solicitud.")
        ElseIf IsValidUserRef(.flags.AceptoReto) Then
            Call CancelarSolicitudReto(.flags.AceptoReto.ArrayIndex, .name & " ha cancelado su admisión.")
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

Private Sub HandleTransFerGold(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim Cantidad As Long
        Dim tUser    As t_UserReference
        Cantidad = reader.ReadInt32()
        username = reader.ReadString8()
        '  Chequeos de seguridad... Estos chequeos ya se hacen en el cliente, pero si no se hacen se puede duplicar oro...
        ' Cantidad válida?
        If Cantidad <= 0 Then Exit Sub
        ' Tiene el oro?
        If .Stats.Banco < Cantidad Then Exit Sub
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            ''Msg77=¡¡Estás muerto!!.)
            Exit Sub
        End If
        'Validate target NPC
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            'Msg1234= Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.
            Call WriteLocaleMsg(UserIndex, 1234, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then Exit Sub
        If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        tUser = NameIndex(username)
        ' Enviar a vos mismo?
        If tUser.ArrayIndex = UserIndex Then
            Call WriteLocaleChatOverHead(UserIndex, 1408, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1408=¡No puedo enviarte oro a vos mismo!
            Exit Sub
        End If
        If Not EsGM(UserIndex) Then
            If Not IsValidUserRef(tUser) Then
                Dim nowRaw As Long
                nowRaw = GetTickCountRaw()
                If TicksElapsed(.Counters.LastTransferGold, nowRaw) >= 10000 Then
                    If PersonajeExiste(username) Then
                        If Not AddOroBancoDatabase(username, Cantidad) Then
                            Call WriteLocaleChatOverHead(UserIndex, 1409, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1409=Error al realizar la operación.
                            Exit Sub
                        Else
                            UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(Cantidad) 'Quitamos el oro al usuario
                        End If
                        Call LogBankTransfer(.name, username, Cantidad, False)
                        .Counters.LastTransferGold = nowRaw
                    Else
                        Call WriteLocaleChatOverHead(UserIndex, 1410, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1410=El usuario no existe.
                        Exit Sub
                    End If
                Else
                    Call WriteLocaleChatOverHead(UserIndex, 1411, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1411=Espera un momento.
                    Exit Sub
                End If
            Else
                UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(Cantidad) 'Quitamos el oro al usuario
                UserList(tUser.ArrayIndex).Stats.Banco = UserList(tUser.ArrayIndex).Stats.Banco + val(Cantidad) 'Se lo damos al otro.
                Call LogBankTransfer(.name, username, Cantidad, True)
            End If
            Call WriteLocaleChatOverHead(UserIndex, 1435, "", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite) ' Msg1435=¡El envío se ha realizado con éxito! Gracias por utilizar los servicios de Finanzas Goliath
        Else
            Call WriteLocaleChatOverHead(UserIndex, 1413, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1413=Los administradores no pueden transferir oro.
            Call LogGM(.name, "Quizo transferirle oro a: " & username)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Private Sub HandleMoveItem(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim SlotViejo As Byte
        Dim SlotNuevo As Byte
        SlotViejo = reader.ReadInt8()
        SlotNuevo = reader.ReadInt8()
        Dim Objeto               As t_Obj
        Dim AppliedElementalTags As Boolean
        Dim tmpElementalTags     As Long
        Dim Equipado             As Boolean
        Dim Equipado2            As Boolean
        Dim Equipado3            As Boolean
        Dim ObjCania             As t_Obj
        'HarThaoS: Si es un hilo de pesca y lo estoy arrastrando en una caña rota borro del slot viejo y en el nuevo pongo la caña correspondiente
        If SlotViejo > getMaxInventorySlots(UserIndex) Or SlotNuevo > getMaxInventorySlots(UserIndex) Or SlotViejo <= 0 Or SlotNuevo <= 0 Then Exit Sub
        If .invent.Object(SlotViejo).ObjIndex = OBJ_FISHING_LINE Then
            Select Case .invent.Object(SlotNuevo).ObjIndex
                Case OBJ_BROKEN_FISHING_ROD_BASIC
                    ObjCania.ObjIndex = OBJ_FISHING_ROD_BASIC
                Case OBJ_BROKEN_FISHING_ROD_COMMON
                    ObjCania.ObjIndex = OBJ_FISHING_ROD_COMMON
                Case OBJ_BROKEN_FISHING_ROD_FINE
                    ObjCania.ObjIndex = OBJ_FISHING_ROD_FINE
                Case OBJ_BROKEN_FISHING_ROD_ELITE
                    ObjCania.ObjIndex = OBJ_FISHING_ROD_ELITE
            End Select
            ObjCania.amount = 1
            'si el objeto que estaba pisando era una caña rota.
            If ObjCania.ObjIndex > 0 Then
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, SlotViejo, 1)
                Call UpdateUserInv(False, UserIndex, SlotViejo)
                Call QuitarUserInvItem(UserIndex, SlotNuevo, 1)
                Call UpdateUserInv(False, UserIndex, SlotNuevo)
                Call MeterItemEnInventario(UserIndex, ObjCania)
                Exit Sub
            End If
        End If
        If IsFeatureEnabled("elemental_tags") Then
            AppliedElementalTags = False
            If CanElementalTagBeApplied(UserIndex, SlotNuevo, SlotViejo) Then
                AppliedElementalTags = True
                Call RemoveItemFromInventory(UserIndex, SlotViejo)
            End If
        End If
        If (SlotViejo > .CurrentInventorySlots) Or (SlotNuevo > .CurrentInventorySlots) Then
            'Msg1235= Espacio no desbloqueado.
            Call WriteLocaleMsg(UserIndex, 1235, e_FontTypeNames.FONTTYPE_INFO)
        Else
            If .invent.Object(SlotNuevo).ObjIndex = .invent.Object(SlotViejo).ObjIndex And .invent.Object(SlotNuevo).ElementalTags = .invent.Object(SlotViejo).ElementalTags Then
                .invent.Object(SlotNuevo).amount = .invent.Object(SlotNuevo).amount + .invent.Object(SlotViejo).amount
                Dim Excedente As Integer
                Excedente = .invent.Object(SlotNuevo).amount - GetMaxInvOBJ()
                If Excedente > 0 Then
                    .invent.Object(SlotViejo).amount = Excedente
                    .invent.Object(SlotNuevo).amount = GetMaxInvOBJ()
                Else
                    If .invent.Object(SlotViejo).Equipped = 1 Then
                        .invent.Object(SlotNuevo).Equipped = 1
                    End If
                    .invent.Object(SlotViejo).ObjIndex = 0
                    .invent.Object(SlotViejo).amount = 0
                    .invent.Object(SlotViejo).Equipped = 0
                    'Cambiamos si alguno es un anillo
                    If .invent.EquippedRingAccesorySlot = SlotViejo Then
                        .invent.EquippedRingAccesorySlot = SlotNuevo
                    End If
                    If .invent.EquippedRingAccesorySlot = SlotViejo Then
                        .invent.EquippedRingAccesorySlot = SlotNuevo
                    End If
                    'Cambiamos si alguno es un armor
                    If .invent.EquippedArmorSlot = SlotViejo Then
                        .invent.EquippedArmorSlot = SlotNuevo
                    End If
                    'Cambiamos si alguno es un barco
                    If .invent.EquippedShipSlot = SlotViejo Then
                        .invent.EquippedShipSlot = SlotNuevo
                    End If
                    'Cambiamos si alguno es una montura
                    If .invent.EquippedSaddleSlot = SlotViejo Then
                        .invent.EquippedSaddleSlot = SlotNuevo
                    End If
                    'Cambiamos si alguno es un casco
                    If .invent.EquippedHelmetSlot = SlotViejo Then
                        .invent.EquippedHelmetSlot = SlotNuevo
                    End If
                    'Cambiamos si alguno es un escudo
                    If .invent.EquippedShieldSlot = SlotViejo Then
                        .invent.EquippedShieldSlot = SlotNuevo
                    End If
                    'Cambiamos si alguno es munición
                    If .invent.EquippedMunitionSlot = SlotViejo Then
                        .invent.EquippedMunitionSlot = SlotNuevo
                    End If
                    'Cambiamos si alguno es un arma
                    If .invent.EquippedWeaponSlot = SlotViejo Then
                        .invent.EquippedWeaponSlot = SlotNuevo
                    End If
                    'Cambiamos si alguno es un magico
                    If .invent.EquippedAmuletAccesorySlot = SlotViejo Then
                        .invent.EquippedAmuletAccesorySlot = SlotNuevo
                    End If
                    'Cambiamos si alguno es una herramienta
                    If .invent.EquippedWorkingToolSlot = SlotViejo Then
                        .invent.EquippedWorkingToolSlot = SlotNuevo
                    End If
                End If
            Else
                If .invent.Object(SlotNuevo).ObjIndex <> 0 And Not AppliedElementalTags Then
                    Objeto.amount = .invent.Object(SlotViejo).amount
                    Objeto.ObjIndex = .invent.Object(SlotViejo).ObjIndex
                    tmpElementalTags = .invent.Object(SlotViejo).ElementalTags
                    If .invent.Object(SlotViejo).Equipped = 1 Then
                        Equipado = True
                    End If
                    If .invent.Object(SlotNuevo).Equipped = 1 Then
                        Equipado2 = True
                    End If
                    '  If .Invent.Object(SlotNuevo).Equipped = 1 And .Invent.Object(SlotViejo).Equipped = 1 Then
                    '     Equipado3 = True
                    ' End If
                    .invent.Object(SlotViejo).ObjIndex = .invent.Object(SlotNuevo).ObjIndex
                    .invent.Object(SlotViejo).amount = .invent.Object(SlotNuevo).amount
                    .invent.Object(SlotViejo).ElementalTags = .invent.Object(SlotNuevo).ElementalTags
                    .invent.Object(SlotNuevo).ObjIndex = Objeto.ObjIndex
                    .invent.Object(SlotNuevo).amount = Objeto.amount
                    .invent.Object(SlotNuevo).ElementalTags = tmpElementalTags
                    If Equipado Then
                        .invent.Object(SlotNuevo).Equipped = 1
                    Else
                        .invent.Object(SlotNuevo).Equipped = 0
                    End If
                    If Equipado2 Then
                        .invent.Object(SlotViejo).Equipped = 1
                    Else
                        .invent.Object(SlotViejo).Equipped = 0
                    End If
                End If
                'Cambiamos si alguno es un anillo
                If .invent.EquippedRingAccesorySlot = SlotViejo Then
                    .invent.EquippedRingAccesorySlot = SlotNuevo
                ElseIf .invent.EquippedRingAccesorySlot = SlotNuevo Then
                    .invent.EquippedRingAccesorySlot = SlotViejo
                End If
                'Cambiamos si alguno es un armor
                If .invent.EquippedArmorSlot = SlotViejo Then
                    .invent.EquippedArmorSlot = SlotNuevo
                ElseIf .invent.EquippedArmorSlot = SlotNuevo Then
                    .invent.EquippedArmorSlot = SlotViejo
                End If
                'Cambiamos si alguno es un barco
                If .invent.EquippedShipSlot = SlotViejo Then
                    .invent.EquippedShipSlot = SlotNuevo
                ElseIf .invent.EquippedShipSlot = SlotNuevo Then
                    .invent.EquippedShipSlot = SlotViejo
                End If
                'Cambiamos si alguno es una montura
                If .invent.EquippedSaddleSlot = SlotViejo Then
                    .invent.EquippedSaddleSlot = SlotNuevo
                ElseIf .invent.EquippedSaddleSlot = SlotNuevo Then
                    .invent.EquippedSaddleSlot = SlotViejo
                End If
                'Cambiamos si alguno es un casco
                If .invent.EquippedHelmetSlot = SlotViejo Then
                    .invent.EquippedHelmetSlot = SlotNuevo
                ElseIf .invent.EquippedHelmetSlot = SlotNuevo Then
                    .invent.EquippedHelmetSlot = SlotViejo
                End If
                'Cambiamos si alguno es un escudo
                If .invent.EquippedShieldSlot = SlotViejo Then
                    .invent.EquippedShieldSlot = SlotNuevo
                ElseIf .invent.EquippedShieldSlot = SlotNuevo Then
                    .invent.EquippedShieldSlot = SlotViejo
                End If
                'Cambiamos si alguno es munición
                If .invent.EquippedMunitionSlot = SlotViejo Then
                    .invent.EquippedMunitionSlot = SlotNuevo
                ElseIf .invent.EquippedMunitionSlot = SlotNuevo Then
                    .invent.EquippedMunitionSlot = SlotViejo
                End If
                'Cambiamos si alguno es un arma
                If .invent.EquippedWeaponSlot = SlotViejo Then
                    .invent.EquippedWeaponSlot = SlotNuevo
                ElseIf .invent.EquippedWeaponSlot = SlotNuevo Then
                    .invent.EquippedWeaponSlot = SlotViejo
                End If
                'Cambiamos si alguno es un magico
                If .invent.EquippedAmuletAccesorySlot = SlotViejo Then
                    .invent.EquippedAmuletAccesorySlot = SlotNuevo
                ElseIf .invent.EquippedAmuletAccesorySlot = SlotNuevo Then
                    .invent.EquippedAmuletAccesorySlot = SlotViejo
                End If
                'Cambiamos si alguno es una herramienta
                If .invent.EquippedWorkingToolSlot = SlotViejo Then
                    .invent.EquippedWorkingToolSlot = SlotNuevo
                ElseIf .invent.EquippedWorkingToolSlot = SlotNuevo Then
                    .invent.EquippedWorkingToolSlot = SlotViejo
                End If
                If Objeto.ObjIndex = 0 And Not AppliedElementalTags Then
                    .invent.Object(SlotNuevo).ObjIndex = .invent.Object(SlotViejo).ObjIndex
                    .invent.Object(SlotNuevo).amount = .invent.Object(SlotViejo).amount
                    .invent.Object(SlotNuevo).Equipped = .invent.Object(SlotViejo).Equipped
                    .invent.Object(SlotNuevo).ElementalTags = .invent.Object(SlotViejo).ElementalTags
                    .invent.Object(SlotViejo).ObjIndex = 0
                    .invent.Object(SlotViejo).amount = 0
                    .invent.Object(SlotViejo).Equipped = 0
                    .invent.Object(SlotViejo).ElementalTags = 0
                End If
            End If
            Call UpdateUserInv(False, UserIndex, SlotViejo)
            Call UpdateUserInv(False, UserIndex, SlotNuevo)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleMoveItem", Erl)
End Sub

Private Sub HandleBovedaMoveItem(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim SlotViejo As Byte
        Dim SlotNuevo As Byte
        SlotViejo = reader.ReadInt8()
        SlotNuevo = reader.ReadInt8()
        Dim Objeto    As t_Obj
        Dim Equipado  As Boolean
        Dim Equipado2 As Boolean
        Dim Equipado3 As Boolean
        If SlotViejo > MAX_BANCOINVENTORY_SLOTS Or SlotNuevo > MAX_BANCOINVENTORY_SLOTS Or SlotViejo <= 0 Or SlotNuevo <= 0 Then Exit Sub
        Objeto.ObjIndex = UserList(UserIndex).BancoInvent.Object(SlotViejo).ObjIndex
        Objeto.amount = UserList(UserIndex).BancoInvent.Object(SlotViejo).amount
        Objeto.ElementalTags = UserList(UserIndex).BancoInvent.Object(SlotViejo).ElementalTags
        UserList(UserIndex).BancoInvent.Object(SlotViejo).ObjIndex = UserList(UserIndex).BancoInvent.Object(SlotNuevo).ObjIndex
        UserList(UserIndex).BancoInvent.Object(SlotViejo).amount = UserList(UserIndex).BancoInvent.Object(SlotNuevo).amount
        UserList(UserIndex).BancoInvent.Object(SlotViejo).ElementalTags = UserList(UserIndex).BancoInvent.Object(SlotNuevo).ElementalTags
        UserList(UserIndex).BancoInvent.Object(SlotNuevo).ObjIndex = Objeto.ObjIndex
        UserList(UserIndex).BancoInvent.Object(SlotNuevo).amount = Objeto.amount
        UserList(UserIndex).BancoInvent.Object(SlotNuevo).ElementalTags = Objeto.ElementalTags
        'Actualizamos el banco
        Call UpdateBanUserInv(False, UserIndex, SlotViejo, "HandleBovedaMoveItem - slot viejo")
        Call UpdateBanUserInv(False, UserIndex, SlotNuevo, "HandleBovedaMoveItem - slot nuevo")
    End With
    Exit Sub
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleBovedaMoveItem", Erl)
End Sub

Private Sub HandleQuieroFundarClan(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        If UserList(UserIndex).flags.Privilegios And e_PlayerType.Consejero Then Exit Sub
        If UserList(UserIndex).GuildIndex > 0 Then
            'Msg1236= Ya perteneces a un clan, no podés fundar otro.
            Call WriteLocaleMsg(UserIndex, 1236, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).Stats.ELV < 23 Or UserList(UserIndex).Stats.UserSkills(e_Skill.liderazgo) < 50 Then
            'Msg1237= Para fundar un clan debes ser Nivel 23, tener 50 en liderazgo y tener en tu inventario las 4 Gemas de Fundación: Gema Verde, Gema Roja, Gema Azul y Gema Polar.
            Call WriteLocaleMsg(UserIndex, 1237, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not TieneObjetos(407, 1, UserIndex) Or Not TieneObjetos(408, 1, UserIndex) Or Not TieneObjetos(409, 1, UserIndex) Or Not TieneObjetos(412, 1, UserIndex) Then
            'Msg1238= Para fundar un clan debes tener en tu inventario las 4 Gemas de Fundación: Gema Verde, Gema Roja, Gema Azul y Gema Polar.
            Call WriteLocaleMsg(UserIndex, 1238, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Msg1239= Servidor » ¡Comenzamos a fundar el clan! Ingresa todos los datos solicitados.
        Call WriteLocaleMsg(UserIndex, 1239, e_FontTypeNames.FONTTYPE_INFO)
        Call WriteShowFundarClanForm(UserIndex)
    End With
    Exit Sub
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuieroFundarClan", Erl)
End Sub

Private Sub HandleLlamadadeClan(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim refError   As String
        Dim clan_nivel As Byte
        If .GuildIndex <> 0 Then
            clan_nivel = modGuilds.NivelDeClan(.GuildIndex)
            If clan_nivel >= RequiredGuildLevelCallSupport Then
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageLocaleMsg(1818, .name & "¬" & get_map_name(.pos.Map) & "¬" & .pos.Map & "¬" & .pos.x & "¬" & _
                        .pos.y, e_FontTypeNames.FONTTYPE_GUILD)) ' Msg1818=Clan> [¬1] solicita apoyo de su clan en ¬2 (¬3-¬4-¬5). Puedes ver su ubicación en el mapa del mundo.
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave("43", NO_3D_SOUND, NO_3D_SOUND))
                Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageUbicacionLlamada(.pos.Map, .pos.x, .pos.y))
            Else
                'Msg1240= Servidor » El nivel de tu clan debe ser ¬ o mayor para utilizar esta opción.
                Call WriteLocaleMsg(UserIndex, 1240, e_FontTypeNames.FONTTYPE_INFO, RequiredGuildLevelCallSupport)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleLlamadadeClan", Erl)
End Sub

Private Sub HandleCasamiento(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        tUser = NameIndex(username)
        If Not IsValidUserRef(tUser) Then
            ' Msg743=Usuario offline.
            Call WriteLocaleMsg(UserIndex, 743, e_FontTypeNames.FONTTYPE_INFO)
        End If
        If IsValidNpcRef(.flags.TargetNPC) Then
            If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Revividor Then
                ' Msg744=Primero haz click sobre un sacerdote.
                Call WriteLocaleMsg(UserIndex, 744, e_FontTypeNames.FONTTYPE_INFO)
            Else
                If Distancia(.pos, NpcList(.flags.TargetNPC.ArrayIndex).pos) > 10 Then
                    Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                Else
                    If tUser.ArrayIndex = UserIndex Then
                        ' Msg745=No podés casarte contigo mismo.
                        Call WriteLocaleMsg(UserIndex, 745, e_FontTypeNames.FONTTYPE_INFO)
                    ElseIf .flags.Casado = 1 Then
                        ' Msg746=¡Ya estás casado! Debes divorciarte de tu actual pareja para casarte nuevamente.
                        Call WriteLocaleMsg(UserIndex, 746, e_FontTypeNames.FONTTYPE_INFO)
                    ElseIf UserList(tUser.ArrayIndex).flags.Casado = 1 Then
                        ' Msg747=Tu pareja debe divorciarse antes de tomar tu mano en matrimonio.
                        Call WriteLocaleMsg(UserIndex, 747, e_FontTypeNames.FONTTYPE_INFO)
                    Else
                        If UserList(tUser.ArrayIndex).flags.Candidato.ArrayIndex = UserIndex Then
                            UserList(tUser.ArrayIndex).flags.Casado = 1
                            UserList(tUser.ArrayIndex).flags.SpouseId = UserList(UserIndex).Id
                            .flags.Casado = 1
                            .flags.SpouseId = UserList(tUser.ArrayIndex).Id
                            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(e_SoundEffects.Casamiento_sound, NO_3D_SOUND, NO_3D_SOUND))
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1651, get_map_name(.pos.Map) & "¬" & UserList(UserIndex).name & "¬" & UserList( _
                                    tUser.ArrayIndex).name, e_FontTypeNames.FONTTYPE_WARNING)) 'Msg1651=El sacerdote de ¬1 celebra el casamiento entre ¬2 y ¬3.
                            Call WriteLocaleChatOverHead(UserIndex, 1414, vbNullString, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1414=Los declaro unidos en legal matrimonio ¡Felicidades!
                            Call WriteLocaleChatOverHead(tUser.ArrayIndex, 1415, vbNullString, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)  ' Msg1415=Los declaro unidos en legal matrimonio ¡Felicidades!
                        Else
                            Call WriteLocaleChatOverHead(UserIndex, 1420, username, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite) ' Msg1420=La solicitud de casamiento a sido enviada a ¬1.
                            Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1956, .name, e_FontTypeNames.FONTTYPE_TALK)) ' Msg1956=¬1 desea casarse contigo, para permitirlo haz click en el sacerdote y escribe /PROPONER ¬1.
                            .flags.Candidato = tUser
                        End If
                    End If
                End If
            End If
        Else
            ' Msg748=Primero haz click sobre el sacerdote.
            Call WriteLocaleMsg(UserIndex, 748, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCasamiento", Erl)
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
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleComenzarTorneo", Erl)
End Sub

Private Sub HandleBusquedaTesoro(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Tipo As Byte
        Tipo = reader.ReadInt8()
        If (.flags.Privilegios And Not (e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.User)) Then
            Select Case Tipo
                Case 0
                    If Not BusquedaTesoroActiva And BusquedaRegaloActiva = False And BusquedaNpcActiva = False Then
                        Call PerderTesoro
                    Else
                        If BusquedaTesoroActiva Then
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1652, get_map_name(TesoroNumMapa) & "¬" & TesoroNumMapa, e_FontTypeNames.FONTTYPE_TALK)) 'Msg1652=Eventos> Todavia nadie fue capaz de encontar el tesoro, recorda que se encuentra en ¬1(¬2). ¿Quien sera el valiente que lo encuentre?
                            'Msg1241= Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: ¬1
                            Call WriteLocaleMsg(UserIndex, 1241, e_FontTypeNames.FONTTYPE_INFO, TesoroNumMapa)
                        Else
                            ' Msg734=Ya hay una busqueda del tesoro activa.
                            Call WriteLocaleMsg(UserIndex, 734, e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                Case 1
                    If Not BusquedaRegaloActiva And BusquedaTesoroActiva = False And BusquedaNpcActiva = False Then
                        Call PerderRegalo
                    Else
                        If BusquedaRegaloActiva Then
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1653, get_map_name(RegaloNumMapa) & "¬" & RegaloNumMapa, e_FontTypeNames.FONTTYPE_TALK)) 'Msg1653=Eventos> Ningún valiente fue capaz de encontrar el item misterioso, recuerda que se encuentra en ¬1(¬2). ¡Ten cuidado!
                            'Msg1242= Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: ¬1
                            Call WriteLocaleMsg(UserIndex, 1242, e_FontTypeNames.FONTTYPE_INFO, RegaloNumMapa)
                        Else
                            ' Msg734=Ya hay una busqueda del tesoro activa.
                            Call WriteLocaleMsg(UserIndex, 734, e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                Case 2
                    If Not BusquedaNpcActiva And BusquedaTesoroActiva = False And BusquedaRegaloActiva = False Then
                        Dim pos As t_WorldPos
                        pos.Map = TesoroNPCMapa(RandomNumber(1, UBound(TesoroNPCMapa)))
                        pos.y = 50
                        pos.x = 50
                        npc_index_evento = SpawnNpc(TesoroNPC(RandomNumber(1, UBound(TesoroNPC))), pos, True, False, True)
                        BusquedaNpcActiva = True
                    Else
                        If BusquedaNpcActiva Then
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1654, NpcList(npc_index_evento).pos.Map, e_FontTypeNames.FONTTYPE_TALK)) 'Msg1654=Eventos> Todavía nadie logró matar el NPC que se encuentra en el mapa ¬1.
                            'Msg1243= Ya hay una busqueda de npc activo. El tesoro se encuentra en: ¬1
                            Call WriteLocaleMsg(UserIndex, 1243, e_FontTypeNames.FONTTYPE_INFO, NpcList(npc_index_evento).pos.Map)
                        Else
                            ' Msg734=Ya hay una busqueda del tesoro activa.
                            Call WriteLocaleMsg(UserIndex, 734, e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
            End Select
        Else
            ' Msg735=Servidor » No estas habilitado para hacer Eventos.
            Call WriteLocaleMsg(UserIndex, 735, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleBusquedaTesoro", Erl)
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
    Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Private Sub HandleCompletarAccion(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Accion As Byte
        Accion = reader.ReadInt8()
        If .Accion.AccionPendiente = True Then
            If .Accion.TipoAccion = Accion Then
                Call CompletarAccionFin(UserIndex)
            Else
                ' Msg749=Servidor » La acción que solicitas no se corresponde.
                Call WriteLocaleMsg(UserIndex, 749, e_FontTypeNames.FONTTYPE_SERVER)
            End If
        Else
            ' Msg750=Servidor » Tu no tenias ninguna acción pendiente.
            Call WriteLocaleMsg(UserIndex, 750, e_FontTypeNames.FONTTYPE_SERVER)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Private Sub HandleInvitarGrupo(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
        Else
            If .Grupo.CantidadMiembros <= UBound(.Grupo.Miembros) Then
                Call WriteWorkRequestTarget(UserIndex, e_Skill.Grupo)
            Else
                ' Msg751=¡No podés invitar a más personas!
                Call WriteLocaleMsg(UserIndex, 751, e_FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
    Exit Sub
HandleInvitarGrupo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleInvitarGrupo", Erl)
End Sub

Private Sub HandleMarcaDeClan(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo HandleMarcaDeClan_Err
    With UserList(UserIndex)
        'Exit sub para anular marca de clan
        Exit Sub
        If UserList(UserIndex).GuildIndex = 0 Then
            Exit Sub
        End If
        If .flags.Muerto = 1 Then
            ''Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim clan_nivel As Byte
        clan_nivel = modGuilds.NivelDeClan(UserList(UserIndex).GuildIndex)
        If clan_nivel > 20 Then
            ' Msg721=Servidor » El nivel de tu clan debe ser 3 para utilizar esta opción.
            Call WriteLocaleMsg(UserIndex, 721, e_FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub
        End If
        Call WriteWorkRequestTarget(UserIndex, e_Skill.MarcaDeClan)
    End With
    Exit Sub
HandleMarcaDeClan_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleMarcaDeClan", Erl)
End Sub

Private Sub HandleResponderPregunta(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim respuesta As Boolean
        Dim DeDonde   As String
        respuesta = reader.ReadBool()
        Dim Log As String
        Log = "Repuesta "
        UserList(UserIndex).flags.RespondiendoPregunta = False
        If respuesta Then
            Select Case UserList(UserIndex).flags.pregunta
                Case 1
                    Log = "Repuesta Afirmativa 1"
                    If UserList(UserIndex).Grupo.EnGrupo Then
                        Call WriteLocaleMsg(UserIndex, MsgYouAreAlreadyInGroup, e_FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub
                    End If
                    If IsValidUserRef(UserList(UserIndex).Grupo.PropuestaDe) Then
                        If UserList(UserList(UserIndex).Grupo.PropuestaDe.ArrayIndex).Grupo.Lider.ArrayIndex <> UserList(UserIndex).Grupo.PropuestaDe.ArrayIndex Then
                            ' Msg722=¡El lider del grupo ha cambiado, imposible unirse!
                            Call WriteLocaleMsg(UserIndex, 722, e_FontTypeNames.FONTTYPE_INFOIAO)
                        Else
                            Log = "Repuesta Afirmativa 1-1 "
                            If Not IsValidUserRef(UserList(UserList(UserIndex).Grupo.PropuestaDe.ArrayIndex).Grupo.Lider) Then
                                ' Msg723=¡El grupo ya no existe!
                                Call WriteLocaleMsg(UserIndex, 723, e_FontTypeNames.FONTTYPE_INFOIAO)
                            Else
                                Log = "Repuesta Afirmativa 1-2 "
                                If UserList(UserList(UserIndex).Grupo.PropuestaDe.ArrayIndex).Grupo.CantidadMiembros = 1 Then
                                    Call GroupCreateSuccess(UserList(UserIndex).Grupo.PropuestaDe.ArrayIndex)
                                    Log = "Repuesta Afirmativa 1-3 "
                                End If
                                Call AddUserToGRoup(UserIndex, UserList(UserIndex).Grupo.PropuestaDe.ArrayIndex)
                            End If
                        End If
                    Else
                        ' Msg724=Servidor » Solicitud de grupo invalida, reintente...
                        Call WriteLocaleMsg(UserIndex, 724, e_FontTypeNames.FONTTYPE_SERVER)
                    End If
                    'unirlo
                Case 2
                    Log = "Repuesta Afirmativa 2"
                    ' Msg725=¡Ahora sos un ciudadano!
                    Call WriteLocaleMsg(UserIndex, 725, e_FontTypeNames.FONTTYPE_INFOIAO)
                    Call VolverCiudadano(UserIndex)
                Case 3
                    Log = "Repuesta Afirmativa 3"
                    UserList(UserIndex).Hogar = UserList(UserIndex).PosibleHogar
                    Select Case UserList(UserIndex).Hogar
                        Case e_Ciudad.cUllathorpe
                            DeDonde = "Ullathorpe"
                        Case e_Ciudad.cNix
                            DeDonde = "Nix"
                        Case e_Ciudad.cBanderbill
                            DeDonde = "Banderbill"
                        Case e_Ciudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                            DeDonde = "Lindos"
                        Case e_Ciudad.cArghal
                            DeDonde = " Arghal"
                        Case e_Ciudad.cForgat
                            DeDonde = " Forgat"
                        Case e_Ciudad.cArkhein
                            DeDonde = " Arkhein"
                        Case e_Ciudad.cEldoria
                            DeDonde = " Eldoria"
                        Case e_Ciudad.cPenthar
                            DeDonde = " Penthar"
                        Case Else
                            DeDonde = "Ullathorpe"
                    End Select
                    If IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
                        Call WriteLocaleChatOverHead(UserIndex, 1421, UserList(UserIndex).name & "¬" & DeDonde, NpcList(UserList( _
                                UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite) ' Msg1421=¡Gracias ¬1! Ahora perteneces a la ciudad de ¬2.
                    Else
                        'Msg1244= ¡Gracias ¬1
                        Call WriteLocaleMsg(UserIndex, 1244, e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).name)
                    End If
                Case 4
                    Log = "Repuesta Afirmativa 4"
                    If IsValidUserRef(UserList(UserIndex).flags.TargetUser) Then
                        Dim TargetIndex As Integer
                        TargetIndex = UserList(UserIndex).flags.TargetUser.ArrayIndex
                        ' Ensure the target index is within bounds
                        If TargetIndex >= LBound(UserList) And TargetIndex <= UBound(UserList) Then
                            UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                            UserList(UserIndex).ComUsu.DestNick = UserList(TargetIndex).name
                            UserList(UserIndex).ComUsu.cant = 0
                            UserList(UserIndex).ComUsu.Objeto = 0
                            UserList(UserIndex).ComUsu.Acepto = False
                            ' Routine to start trading with another user
                            Call IniciarComercioConUsuario(UserIndex, TargetIndex)
                        Else
                            ' Invalid index; send error message
                            ' Msg726=Servidor » Solicitud de comercio invalida, reintente...
                            Call WriteLocaleMsg(UserIndex, 726, e_FontTypeNames.FONTTYPE_SERVER)
                        End If
                    Else
                        ' Invalid reference; send error message
                        ' Msg726=Servidor » Solicitud de comercio invalida, reintente...
                        Call WriteLocaleMsg(UserIndex, 726, e_FontTypeNames.FONTTYPE_SERVER)
                    End If
                Case 5
                    Dim i As Integer, j As Integer
                    With UserList(UserIndex)
                        For i = 1 To MAX_INVENTORY_SLOTS
                            For j = 1 To UBound(PecesEspeciales)
                                If .invent.Object(i).ObjIndex = PecesEspeciales(j).ObjIndex Then
                                    .Stats.PuntosPesca = .Stats.PuntosPesca + (ObjData(.invent.Object(i).ObjIndex).PuntosPesca * .invent.Object(i).amount)
                                    .Stats.GLD = .Stats.GLD + (ObjData(.invent.Object(i).ObjIndex).Valor * .invent.Object(i).amount * SvrConfig.GetValue( _
                                            "SpecialFishGoldMultiplier"))
                                    Call WriteUpdateGold(UserIndex)
                                    If IsFeatureEnabled("gain_exp_while_working") Then
                                        .Stats.Exp = .Stats.Exp + (ObjData(.invent.Object(i).ObjIndex).Valor * .invent.Object(i).amount * SvrConfig.GetValue( _
                                                "SpecialFishExpMultiplier"))
                                        Call WriteUpdateExp(UserIndex)
                                        Call CheckUserLevel(UserIndex)
                                    End If
                                    Call QuitarUserInvItem(UserIndex, i, .invent.Object(i).amount)
                                    Call UpdateUserInv(False, UserIndex, i)
                                End If
                            Next j
                        Next i
                        Dim charindexstr As Integer
                        charindexstr = str(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex)
                        If charindexstr > 0 Then
                            Call WriteLocaleChatOverHead(UserIndex, 1422, .Stats.PuntosPesca, charindexstr, &HFFFF00) ' Msg1422=¡Felicitaciones! Ahora tienes un total de ¬1 puntos de pesca.
                        End If
                        .flags.pregunta = 0
                    End With
                Case Else
                    ' Msg727=No tienes preguntas pendientes.
                    Call WriteLocaleMsg(UserIndex, 727, e_FontTypeNames.FONTTYPE_INFOIAO)
            End Select
        Else
            Log = "Repuesta negativa"
            Select Case UserList(UserIndex).flags.pregunta
                Case 1
                    Log = "Repuesta negativa 1"
                    If IsValidUserRef(UserList(UserIndex).Grupo.PropuestaDe) Then
                        'Msg1245= El usuario no esta interesado en formar parte del grupo.
                        Call WriteLocaleMsg(UserList(UserIndex).Grupo.PropuestaDe.ArrayIndex, "1245", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                    Call SetUserRef(UserList(UserIndex).Grupo.PropuestaDe, 0)
                    'Msg1246= Has rechazado la propuesta.
                    Call WriteLocaleMsg(UserIndex, 1246, e_FontTypeNames.FONTTYPE_INFO)
                Case 2
                    Log = "Repuesta negativa 2"
                    'Msg1247= ¡Continuas siendo neutral!
                    Call WriteLocaleMsg(UserIndex, 1247, e_FontTypeNames.FONTTYPE_INFO)
                    Call VolverCriminal(UserIndex)
                Case 3
                    Log = "Repuesta negativa 3"
                    Select Case UserList(UserIndex).PosibleHogar
                        Case e_Ciudad.cUllathorpe
                            DeDonde = "Ullathorpe"
                        Case e_Ciudad.cNix
                            DeDonde = "Nix"
                        Case e_Ciudad.cBanderbill
                            DeDonde = "Banderbill"
                        Case e_Ciudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                            DeDonde = "Lindos"
                        Case e_Ciudad.cArghal
                            DeDonde = " Arghal"
                        Case e_Ciudad.cForgat
                            DeDonde = " Forgat"
                        Case e_Ciudad.cArkhein
                            DeDonde = " Arkhein"
                        Case e_Ciudad.cEldoria
                            DeDonde = " Eldoria"
                        Case e_Ciudad.cPenthar
                            DeDonde = " Penthar"
                        Case Else
                            DeDonde = "Ullathorpe"
                    End Select
                    If IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
                        Call WriteLocaleChatOverHead(UserIndex, 1423, UserList(UserIndex).name & "¬" & DeDonde, NpcList(UserList( _
                                UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite) ' Msg1423=¡No hay problema ¬1! Sos bienvenido en ¬2 cuando gustes.
                    End If
                    UserList(UserIndex).PosibleHogar = UserList(UserIndex).Hogar
                Case 4
                    Log = "Repuesta negativa 4"
                    If IsValidUserRef(UserList(UserIndex).flags.TargetUser) Then
                        'Msg1248= El usuario no desea comerciar en este momento.
                        Call WriteLocaleMsg(UserList(UserIndex).flags.TargetUser.ArrayIndex, "1248", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Case 5
                    Log = "Repuesta negativa 5"
                Case Else
                    ' Msg727=No tienes preguntas pendientes.
                    Call WriteLocaleMsg(UserIndex, 727, e_FontTypeNames.FONTTYPE_INFOIAO)
            End Select
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleResponderPregunta", Erl)
End Sub

Private Sub HandleRequestGrupo(ByVal UserIndex As Integer)
    On Error GoTo hErr
    'Author: Pablo Mercavides
    Call WriteDatosGrupo(UserIndex)
    Exit Sub
hErr:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestGrupo", Erl)
End Sub

Private Sub HandleAbandonarGrupo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo HandleAbandonarGrupo_Err
    With UserList(UserIndex)
        Call reader.ReadInt16
        If UserList(UserIndex).Grupo.Lider.ArrayIndex = UserIndex Then
            Call FinalizarGrupo(UserIndex)
        Else
            Call SalirDeGrupo(UserIndex)
        End If
    End With
    Exit Sub
HandleAbandonarGrupo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleAbandonarGrupo", Erl)
End Sub

Private Sub HandleHecharDeGrupo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo HandleHecharDeGrupo_Err
    With UserList(UserIndex)
        Dim Indice As Byte
        Indice = reader.ReadInt8()
        Call EcharMiembro(UserIndex, Indice)
    End With
    Exit Sub
HandleHecharDeGrupo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleHecharDeGrupo", Erl)
End Sub

Private Sub HandleMacroPos(ByVal UserIndex As Integer)
    On Error GoTo HandleMacroPos_Err
    With UserList(UserIndex)
        .ChatCombate = reader.ReadInt8()
        .ChatGlobal = reader.ReadInt8()
        .ShowNothingInterestingMessage = reader.ReadInt8()
    End With
    Exit Sub
HandleMacroPos_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleMacroPos", Erl)
End Sub

Private Sub HandleSubastaInfo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo HandleSubastaInfo_Err
    With UserList(UserIndex)
        If Subasta.HaySubastaActiva Then
            'Msg1249= Subastador: ¬1
            Call WriteLocaleMsg(UserIndex, 1249, e_FontTypeNames.FONTTYPE_INFO, Subasta.Subastador)
            'Msg1250= Objeto: ¬1
            Call WriteLocaleMsg(UserIndex, 1250, e_FontTypeNames.FONTTYPE_INFO, ObjData(Subasta.ObjSubastado).name)
            If Subasta.HuboOferta Then
                'Msg1251= Mejor oferta: ¬1
                Call WriteLocaleMsg(UserIndex, 1251, e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(Subasta.MejorOferta))
                'Msg1252= Podes realizar una oferta escribiendo /OFERTAR ¬1
                Call WriteLocaleMsg(UserIndex, 1252, e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(Subasta.MejorOferta + 100))
            Else
                'Msg1253= Oferta inicial: ¬1
                Call WriteLocaleMsg(UserIndex, 1253, e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(Subasta.OfertaInicial))
                'Msg1254= Podes realizar una oferta escribiendo /OFERTAR ¬1
                Call WriteLocaleMsg(UserIndex, 1254, e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(Subasta.OfertaInicial + 100))
            End If
            'Msg1255= Tiempo Restante de subasta:  ¬1
            Call WriteLocaleMsg(UserIndex, 1255, e_FontTypeNames.FONTTYPE_INFO, SumarTiempo(Subasta.TiempoRestanteSubasta))
        Else
            ' Msg728=No hay ninguna subasta activa en este momento.
            Call WriteLocaleMsg(UserIndex, 728, e_FontTypeNames.FONTTYPE_SUBASTA)
        End If
    End With
    Exit Sub
HandleSubastaInfo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSubastaInfo", Erl)
End Sub

Private Sub HandleCancelarExit(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo HandleCancelarExit_Err
    Call CancelExit(UserIndex)
    Exit Sub
HandleCancelarExit_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCancelarExit", Erl)
End Sub

Private Sub HandleEventoInfo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo HandleEventoInfo_Err
    With UserList(UserIndex)
        If EventoActivo Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1957, PublicidadEvento & "¬" & TiempoRestanteEvento, e_FontTypeNames.FONTTYPE_New_Eventos)) ' Msg1957=¬1. Tiempo restante: ¬2 minuto(s).
        Else
            ' Msg729=Eventos> Actualmente no hay ningún evento en curso.
            Call WriteLocaleMsg(UserIndex, 729, e_FontTypeNames.FONTTYPE_New_Eventos)
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
            'Msg1256= Eventos> El proximo evento ¬1
            Call WriteLocaleMsg(UserIndex, 1256, e_FontTypeNames.FONTTYPE_INFO, DescribirEvento(HoraProximo))
        Else
            ' Msg730=Eventos> No hay eventos próximos.
            Call WriteLocaleMsg(UserIndex, 730, e_FontTypeNames.FONTTYPE_New_Eventos)
        End If
    End With
    Exit Sub
HandleEventoInfo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleEventoInfo", Erl)
End Sub

Private Sub HandleCrearEvento(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Tipo           As Byte
        Dim Duracion       As Byte
        Dim multiplicacion As Byte
        Tipo = reader.ReadInt8()
        Duracion = reader.ReadInt8()
        multiplicacion = reader.ReadInt8()
        If multiplicacion > 5 Then 'no superar este multiplicador
            multiplicacion = 2
        End If
        '/ dejar solo Administradores
        If .flags.Privilegios >= e_PlayerType.Admin Then
            If EventoActivo = False Then
                If LenB(Tipo) = 0 Or LenB(Duracion) = 0 Or LenB(multiplicacion) = 0 Then
                    ' Msg731=Utilice /CREAREVENTO TIPO@DURACION@MULTIPLICACION.
                    Call WriteLocaleMsg(UserIndex, 731, e_FontTypeNames.FONTTYPE_New_Eventos)
                Else
                    Call ForzarEvento(Tipo, Duracion, multiplicacion, UserList(UserIndex).name)
                End If
            Else
                ' Msg732=Ya hay un evento en curso. Finalicelo con /FINEVENTO primero.
                Call WriteLocaleMsg(UserIndex, 732, e_FontTypeNames.FONTTYPE_New_Eventos)
            End If
        Else
            ' Msg733=Servidor » Solo Administradores pueder crear estos eventos.
            Call WriteLocaleMsg(UserIndex, 733, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Private Sub HandleCompletarViaje(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Destino As Byte
        Dim costo   As Long
        Destino = reader.ReadInt8()
        costo = reader.ReadInt32()
        '  WTF el costo lo decide el cliente... Desactivo....
        Exit Sub
        If costo <= 0 Then Exit Sub
        Dim DeDonde As t_CityWorldPos
        If UserList(UserIndex).Stats.GLD < costo Then
            'Msg1257= No tienes suficiente dinero.
            Call WriteLocaleMsg(UserIndex, 1257, e_FontTypeNames.FONTTYPE_INFO)
        Else
            Select Case Destino
                Case e_Ciudad.cUllathorpe
                    DeDonde = CityUllathorpe
                Case e_Ciudad.cNix
                    DeDonde = CityNix
                Case e_Ciudad.cBanderbill
                    DeDonde = CityBanderbill
                Case e_Ciudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
                    DeDonde = CityLindos
                Case e_Ciudad.cArghal
                    DeDonde = CityArghal
                Case e_Ciudad.cForgat
                    DeDonde = CityForgat
                Case e_Ciudad.cArkhein
                    DeDonde = CityArkhein
                Case e_Ciudad.cEldoria
                    DeDonde = CityEldoria
                Case e_Ciudad.cPenthar
                    DeDonde = CityPenthar
                Case Else
                    DeDonde = CityUllathorpe
            End Select
            If DeDonde.NecesitaNave > 0 Then
                If UserList(UserIndex).Stats.UserSkills(e_Skill.Navegacion) < 80 Then
                    'Msg1258= Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.
                    Call WriteLocaleMsg(UserIndex, 1258, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg1259= Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.
                    Call WriteLocaleMsg(UserIndex, 1259, e_FontTypeNames.FONTTYPE_INFO)
                Else
                    If IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
                        If NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundClose <> 0 Then
                            Call WritePlayWave(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundClose, NO_3D_SOUND, NO_3D_SOUND, , 1)
                        End If
                    End If
                    Call WarpToLegalPos(UserIndex, DeDonde.MapaViaje, DeDonde.ViajeX, DeDonde.ViajeY, True)
                    'Msg1260= Has viajado por varios días, te sientes exhausto!
                    Call WriteLocaleMsg(UserIndex, 1260, e_FontTypeNames.FONTTYPE_INFO)
                    UserList(UserIndex).Stats.MinAGU = 0
                    UserList(UserIndex).Stats.MinHam = 0
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - costo
                    Call WriteUpdateHungerAndThirst(UserIndex)
                    Call WriteUpdateUserStats(UserIndex)
                End If
            Else
                Dim Map As Integer
                Dim x   As Byte
                Dim y   As Byte
                Map = DeDonde.MapaViaje
                x = DeDonde.ViajeX
                y = DeDonde.ViajeY
                If IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
                    If NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundClose <> 0 Then
                        Call WritePlayWave(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundClose, NO_3D_SOUND, NO_3D_SOUND, , 1)
                    End If
                End If
                Call WarpUserChar(UserIndex, Map, x, y, True)
                'Msg1261= Has viajado por varios días, te sientes exhausto!
                Call WriteLocaleMsg(UserIndex, 1261, e_FontTypeNames.FONTTYPE_INFO)
                UserList(UserIndex).Stats.MinAGU = 0
                UserList(UserIndex).Stats.MinHam = 0
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - costo
                Call WriteUpdateHungerAndThirst(UserIndex)
                Call WriteUpdateUserStats(UserIndex)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCompletarViaje", Erl)
End Sub

Public Sub HandleQuest(ByVal UserIndex As Integer)
    On Error GoTo HandleQuest_Err
    If Not IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then Exit Sub
    Dim NpcIndex As Integer
    Dim tmpByte  As Byte
    NpcIndex = UserList(UserIndex).flags.TargetNPC.ArrayIndex
    'Esta el personaje en la distancia correcta?
    If Distancia(UserList(UserIndex).pos, NpcList(NpcIndex).pos) > 5 Then
        ' Msg8=Estas demasiado lejos.
        Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    'El NPC hace quests?
    If NpcList(NpcIndex).NumQuest = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareLocalizedChatOverHead(2085, NpcList(NpcIndex).Char.charindex, vbWhite))
        Exit Sub
    End If
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareLocalizedChatOverHead(2086, NpcList(NpcIndex).Char.charindex, vbWhite))
    Exit Sub
HandleQuest_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuest", Erl)
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
    Dim tmpIndex  As Integer
    Indice = reader.ReadInt8
    If Not IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) And UserList(UserIndex).flags.QuestOpenByObj = False Then Exit Sub
    NpcIndex = UserList(UserIndex).flags.TargetNPC.ArrayIndex
    Dim tmpQuest As t_Quest
    'npc or item quest
    If NpcIndex > 0 Then
        'npc handled quest
        tmpQuest = QuestList(NpcList(NpcIndex).QuestNumber(Indice))
        tmpIndex = NpcList(NpcIndex).QuestNumber(Indice)
    Else
        'item handled quest
        tmpIndex = UserList(UserIndex).flags.QuestNumber
        tmpQuest = QuestList(tmpIndex)
    End If
    If Not ModQuest.CanUserAcceptQuest(UserIndex, NpcIndex, tmpIndex, tmpQuest) Then
        Exit Sub
    End If
    QuestSlot = FreeQuestSlot(UserIndex)
    If QuestSlot = 0 Then
        Call WriteLocaleChatOverHead(UserIndex, 1417, vbNullString, NpcList(NpcIndex).Char.charindex, vbYellow)  ' Msg1417=Debes completar las misiones en curso para poder aceptar más misiones.
        Exit Sub
    End If
    'Agregamos la quest.
    With UserList(UserIndex).QuestStats.Quests(QuestSlot)
        .QuestIndex = tmpIndex
        If QuestList(.QuestIndex).RequiredNPCs Then ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
        If QuestList(.QuestIndex).RequiredTargetNPCs Then ReDim .NPCsTarget(1 To QuestList(.QuestIndex).RequiredTargetNPCs)
        UserList(UserIndex).flags.ModificoQuests = True
        'Msg1264= Has aceptado la misión ¬1
        Call WriteLocaleMsg(UserIndex, 1264, e_FontTypeNames.FONTTYPE_INFOIAO, .QuestIndex)
        If NpcIndex > 0 Then
            If (FinishQuestCheck(UserIndex, .QuestIndex, QuestSlot)) Then
                Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 3)
            Else
                Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 4)
            End If
        Else
            Call RemoveItemFromInventory(UserIndex, UserList(UserIndex).flags.QuestItemSlot)
        End If
    End With
    Exit Sub
HandleQuestAccept_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuestAccept", Erl)
End Sub

Public Sub HandleQuestDetailsRequest(ByVal UserIndex As Integer)
    On Error GoTo HandleQuestDetailsRequest_Err
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el paquete QuestInfoRequest.
    'Last modified: 30/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim QuestSlot As Byte
    QuestSlot = reader.ReadInt8
    If QuestSlot <= MAXUSERQUESTS And QuestSlot > 0 Then
        If UserList(UserIndex).QuestStats.Quests(QuestSlot).QuestIndex > 0 Then
            Call WriteQuestDetails(UserIndex, UserList(UserIndex).QuestStats.Quests(QuestSlot).QuestIndex, QuestSlot)
        End If
    End If
    Exit Sub
HandleQuestDetailsRequest_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuestDetailsRequest", Erl)
End Sub
 
Public Sub HandleQuestAbandon(ByVal UserIndex As Integer)
    On Error GoTo HandleQuestAbandon_Err
    With UserList(UserIndex)
        Dim Slot As Byte
        Slot = reader.ReadInt8
        If Slot > MAXUSERQUESTS Then Exit Sub
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
                            Call QuitarObjetos(ObjIndex, GetMaxInvOBJ(), UserIndex)
                        End If
                    End If
                Next i
            End If
        End With
        'Le avisamos que abandono la quest
        'Msg2115=Has abandonado la misión ¬1.
        Call WriteLocaleMsg(UserIndex, 2115, e_FontTypeNames.FONTTYPE_INFOIAO, QuestList(UserList(UserIndex).QuestStats.Quests(Slot).QuestIndex).nombre)
        'Borramos la quest.
        Call CleanQuestSlot(UserIndex, Slot)
        'Ordenamos la lista de quests del usuario.
        Call ArrangeUserQuests(UserIndex)
        'Enviamos la lista de quests actualizada.
        Call WriteQuestListSend(UserIndex)
    End With
    Exit Sub
HandleQuestAbandon_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuestAbandon", Erl)
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
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuestListRequest", Erl)
End Sub

''
' Handles the "Consulta" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleConsulta(ByVal UserIndex As Integer)
    'Habilita/Deshabilita el modo consulta.
    'Agrego validaciones.
    'No se hace visible en los clientes si estaba navegando (porque ya lo estaba).
    Dim UserConsulta As t_UserReference
    With UserList(UserIndex)
        Dim Nick As String
        Nick = reader.ReadString8
        ' Comando exclusivo para gms
        If Not EsGM(UserIndex) Then Exit Sub
        If Len(Nick) <> 0 Then
            UserConsulta = NameIndex(Nick)
            'Se asegura que el target exista
            If Not IsValidUserRef(UserConsulta) Then
                'Msg1265= El usuario se encuentra offline.
                Call WriteLocaleMsg(UserIndex, 1265, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call SetUserRef(UserConsulta, .flags.TargetUser.ArrayIndex)
            'Se asegura que el target exista
            If IsValidUserRef(UserConsulta) Then
                'Msg1266= Primero tienes que seleccionar un usuario, haz click izquierdo sobre el.
                Call WriteLocaleMsg(UserIndex, 1266, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        ' No podes ponerte a vos mismo en modo consulta.
        If UserConsulta.ArrayIndex = UserIndex Then Exit Sub
        ' No podes estra en consulta con otro gm
        If EsGM(UserConsulta.ArrayIndex) Then
            'Msg1267= No puedes iniciar el modo consulta con otro administrador.
            Call WriteLocaleMsg(UserIndex, 1267, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        ' Si ya estaba en consulta, termina la consulta
        If UserList(UserConsulta.ArrayIndex).flags.EnConsulta Then
            'Msg1268= Has terminado el modo consulta con ¬1
            Call WriteLocaleMsg(UserIndex, 1268, e_FontTypeNames.FONTTYPE_INFO, UserList(UserConsulta.ArrayIndex).name)
            'Msg1269= Has terminado el modo consulta.
            Call WriteLocaleMsg(UserConsulta.ArrayIndex, 1269, e_FontTypeNames.FONTTYPE_INFO)
            Call LogGM(.name, "Termino consulta con " & UserList(UserConsulta.ArrayIndex).name)
            UserList(UserConsulta.ArrayIndex).flags.EnConsulta = False
            ' Sino la inicia
        Else
            'Msg1270= Has iniciado el modo consulta con ¬1
            Call WriteLocaleMsg(UserIndex, 1270, e_FontTypeNames.FONTTYPE_INFO, UserList(UserConsulta.ArrayIndex).name)
            'Msg1271= Has iniciado el modo consulta.
            Call WriteLocaleMsg(UserConsulta.ArrayIndex, 1271, e_FontTypeNames.FONTTYPE_INFO)
            Call LogGM(.name, "Inicio consulta con " & UserList(UserConsulta.ArrayIndex).name)
            With UserList(UserConsulta.ArrayIndex)
                If Not EstaPCarea(UserIndex, UserConsulta.ArrayIndex) Then
                    Dim x As Byte
                    Dim y As Byte
                    x = .pos.x
                    y = .pos.y
                    Call FindLegalPos(UserIndex, .pos.Map, x, y)
                    Call WarpUserChar(UserIndex, .pos.Map, x, y, True)
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
                    .Counters.DisabledInvisibility = 0
                    If UserList(UserConsulta.ArrayIndex).flags.Navegando = 0 Then
                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                    End If
                End If
            End With
        End If
        Call SetModoConsulta(UserConsulta.ArrayIndex)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleConsulta", Erl)
End Sub

Private Sub HandleGetMapInfo(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If EsGM(UserIndex) Then
            Dim response As String
            response = "[Info de mapa " & .pos.Map & "]" & vbNewLine
            response = response & "Nombre = " & MapInfo(.pos.Map).map_name & vbNewLine
            response = response & "Seguro = " & MapInfo(.pos.Map).Seguro & vbNewLine
            response = response & "Newbie = " & MapInfo(.pos.Map).Newbie & vbNewLine
            response = response & "Nivel = " & MapInfo(.pos.Map).MinLevel & "/" & MapInfo(.pos.Map).MaxLevel & vbNewLine
            response = response & "SinInviOcul = " & MapInfo(.pos.Map).SinInviOcul & vbNewLine
            response = response & "SinMagia = " & MapInfo(.pos.Map).SinMagia & vbNewLine
            response = response & "SoloClanes = " & MapInfo(.pos.Map).SoloClanes & vbNewLine
            response = response & "NoPKs = " & MapInfo(.pos.Map).NoPKs & vbNewLine
            response = response & "NoCiudadanos = " & MapInfo(.pos.Map).NoCiudadanos & vbNewLine
            response = response & "Salida = " & MapInfo(.pos.Map).Salida.Map & "-" & MapInfo(.pos.Map).Salida.x & "-" & MapInfo(.pos.Map).Salida.y & vbNewLine
            response = response & "Terreno = " & MapInfo(.pos.Map).terrain & vbNewLine
            response = response & "NoCiudadanos = " & MapInfo(.pos.Map).NoCiudadanos & vbNewLine
            response = response & "Zona = " & MapInfo(.pos.Map).zone & vbNewLine
            Call WriteConsoleMsg(UserIndex, response, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
End Sub

Private Sub HandleSeguroResu(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .flags.SeguroResu = Not .flags.SeguroResu
        Call WriteSeguroResu(UserIndex, .flags.SeguroResu)
    End With
End Sub

Private Sub HandleLegionarySecure(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .flags.LegionarySecure = Not .flags.LegionarySecure
        Call WriteLegionarySecure(UserIndex, .flags.LegionarySecure)
    End With
End Sub

Private Sub HandleCuentaExtractItem(ByVal UserIndex As Integer)
    On Error GoTo HandleCuentaExtractItem_Err
    With UserList(UserIndex)
        Dim Slot        As Byte
        Dim slotdestino As Byte
        Dim amount      As Integer
        Slot = reader.ReadInt8()
        amount = reader.ReadInt16()
        slotdestino = reader.ReadInt8()
        If .flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Exit Sub
        End If
        If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then
            Exit Sub
        End If
    End With
    Exit Sub
HandleCuentaExtractItem_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCuentaExtractItem", Erl)
End Sub

Private Sub HandleCuentaDeposit(ByVal UserIndex As Integer)
    On Error GoTo HandleCuentaDeposit_Err
    With UserList(UserIndex)
        Dim Slot        As Byte
        Dim slotdestino As Byte
        Dim amount      As Integer
        Slot = reader.ReadInt8()
        amount = reader.ReadInt16()
        slotdestino = reader.ReadInt8()
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'íEl target es un NPC valido?
        If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
        'íEl NPC puede comerciar?
        If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then
            Exit Sub
        End If
        If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End With
    Exit Sub
HandleCuentaDeposit_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCuentaDeposit", Erl)
End Sub

Private Sub HandleCommerceSendChatMessage(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim chatMessage As String
        chatMessage = "[" & UserList(UserIndex).name & "] " & reader.ReadString8
        'El mensaje se lo envío al destino
        If Not IsValidUserRef(UserList(UserIndex).ComUsu.DestUsu) Then Exit Sub
        Call WriteCommerceRecieveChatMessage(UserList(UserIndex).ComUsu.DestUsu.ArrayIndex, chatMessage)
        'y tambien a mi mismo
        Call WriteCommerceRecieveChatMessage(UserIndex, chatMessage)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceSendChatMessage", Erl)
End Sub

Private Sub HandleLogMacroClickHechizo(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim tipoMacro As Byte
        Dim mensaje   As String
        Dim clicks    As Long
        Dim Motivo    As String
        tipoMacro = reader.ReadInt8
        clicks = reader.ReadInt32
        mensaje = "Control AntiCheat--> El usuario " & .name & "| está utilizando "
        Select Case tipoMacro
            Case tMacro.Coordenadas
                Motivo = "macro de COORDENADAS"
            Case tMacro.dobleclick
                Motivo = "macro de DOBLE CLICK (CANTIDAD DE CLICKS: " & clicks & ")"
            Case tMacro.inasistidoPosFija
                Dim spellID As Integer
                spellID = .Stats.UserHechizos(.flags.Hechizo)
                If Not IsUnassistedSpellAllowed(spellID) Then
                    Motivo = "macro de INASISTIDO"
                End If
            Case tMacro.borrarCartel
                Motivo = "macro de CARTELEO"
        End Select
        If Motivo <> "" Then
            Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageConsoleMsg(mensaje & Motivo & ".", e_FontTypeNames.FONTTYPE_INFO))
        End If
    End With
End Sub

Private Function IsUnassistedSpellAllowed(ByVal spellID As Integer) As Boolean
    Select Case spellID
        Case SPELL_UNASSISTED_DARDO, SPELL_UNASSISTED_RUGIDO_SALVAJE, SPELL_UNASSISTED_RUGIDO_SALVAJE, SPELL_UNASSISTED_FULGOR_IGNEO, SPELL_UNASSISTED_LATIDO_IGNEO, SPELL_UNASSISTED_ECO_IGNEO, SPELL_UNASSISTED_DESTELLO_MALVA, _
            SPELL_UNASSISTED_FRACTURA_GLACIAL, SPELL_UNASSISTED_ALIENTO_CARMESI, SPELL_UNASSISTED_ENERGIA_ANCESTRAL
            IsUnassistedSpellAllowed = True
        Case Else
            IsUnassistedSpellAllowed = False
    End Select
End Function

Private Sub HandleHome(ByVal UserIndex As Integer)
    On Error GoTo HandleHome_Err
    'Add the UCase$ to prevent problems.
    With UserList(UserIndex)
        If IsInMapCarcelRestrictedArea(UserList(UserIndex).pos) Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MSG_CANNOT_USE_HOME_IN_JAIL, vbNullString, e_FontTypeNames.FONTTYPE_INFO))
            Exit Sub
        End If
        If .flags.Muerto = 0 Then
            'Msg1272= Debes estar muerto para utilizar este comando.
            Call WriteLocaleMsg(UserIndex, 1272, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Si el mapa tiene alguna restriccion (newbie, dungeon, etc...), no lo dejamos viajar.
        If MapInfo(.pos.Map).zone = "NEWBIE" Or MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
            'Msg1273= No pueder viajar a tu hogar desde este mapa.
            Call WriteLocaleMsg(UserIndex, 1273, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Si es un mapa comun y no esta en cana
        If .Counters.Pena <> 0 Then
            'Msg1274= No puedes usar este comando en prisión.
            Call WriteLocaleMsg(UserIndex, 1274, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.EnReto Then
            'Msg1275= No podés regresar desde un reto. Usa /ABANDONAR para admitir la derrota y volver a la ciudad.
            Call WriteLocaleMsg(UserIndex, 1275, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.Traveling = 0 Then
            If .pos.Map <> Ciudades(.Hogar).Map Then
                Call goHome(UserIndex)
            Else
                'Msg1276= Ya te encuentras en tu hogar.
                Call WriteLocaleMsg(UserIndex, 1276, e_FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            .flags.Traveling = 0
            .Counters.goHome = 0
            'Msg1277= Ya hay un viaje en curso.
            Call WriteLocaleMsg(UserIndex, 1277, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleHome_Err:
    Call TraceError(Err.Number, Err.Description, "Hogar.HandleHome", Erl)
End Sub

Private Sub HandleAddItemCrafting(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim InvSlot As Byte, CraftSlot As Byte
        InvSlot = reader.ReadInt8
        CraftSlot = reader.ReadInt8
        If .flags.Crafteando = 0 Then Exit Sub
        If InvSlot < 1 Or InvSlot > .CurrentInventorySlots Then Exit Sub
        If .invent.Object(InvSlot).ObjIndex = 0 Then Exit Sub
        If CraftSlot < 1 Then
            For CraftSlot = 1 To MAX_SLOTS_CRAFTEO
                If .CraftInventory(CraftSlot) = 0 Then
                    Exit For
                End If
            Next
        End If
        If CraftSlot > MAX_SLOTS_CRAFTEO Then
            Exit Sub
        End If
        If .CraftInventory(CraftSlot) <> 0 Then Exit Sub
        .CraftInventory(CraftSlot) = .invent.Object(InvSlot).ObjIndex
        Call QuitarUserInvItem(UserIndex, InvSlot, 1)
        Call UpdateUserInv(False, UserIndex, InvSlot)
        Call WriteCraftingItem(UserIndex, CraftSlot, .CraftInventory(CraftSlot))
        Dim Result As clsCrafteo
        Set Result = CheckCraftingResult(UserIndex)
        If Not Result Is .CraftResult Then
            Set .CraftResult = Result
            If Not .CraftResult Is Nothing Then
                Call WriteCraftingResult(UserIndex, .CraftResult.Resultado, CalculateCraftProb(UserIndex, .CraftResult.Probabilidad), .CraftResult.precio)
            Else
                Call WriteCraftingResult(UserIndex, 0)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleAddItemCrafting", Erl)
End Sub

Private Sub HandleRemoveItemCrafting(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim InvSlot As Byte, CraftSlot As Byte
        CraftSlot = reader.ReadInt8
        InvSlot = reader.ReadInt8
        If .flags.Crafteando = 0 Then Exit Sub
        If CraftSlot < 1 Or CraftSlot > MAX_SLOTS_CRAFTEO Then Exit Sub
        If .CraftInventory(CraftSlot) = 0 Then Exit Sub
        If InvSlot < 1 Then
            Dim TmpObj As t_Obj
            TmpObj.ObjIndex = .CraftInventory(CraftSlot)
            TmpObj.amount = 1
            If Not MeterItemEnInventario(UserIndex, TmpObj) Then Exit Sub
        ElseIf InvSlot <= .CurrentInventorySlots Then
            If .invent.Object(InvSlot).ObjIndex = 0 Then
                .invent.Object(InvSlot).ObjIndex = .CraftInventory(CraftSlot)
            ElseIf .invent.Object(InvSlot).ObjIndex <> .CraftInventory(CraftSlot) Then
                Exit Sub
            End If
            .invent.Object(InvSlot).amount = .invent.Object(InvSlot).amount + 1
            Call UpdateUserInv(False, UserIndex, InvSlot)
        End If
        .CraftInventory(CraftSlot) = 0
        Call WriteCraftingItem(UserIndex, CraftSlot, 0)
        Dim Result As clsCrafteo
        Set Result = CheckCraftingResult(UserIndex)
        If Not Result Is .CraftResult Then
            Set .CraftResult = Result
            If Not .CraftResult Is Nothing Then
                Call WriteCraftingResult(UserIndex, .CraftResult.Resultado, CalculateCraftProb(UserIndex, .CraftResult.Probabilidad), .CraftResult.precio)
            Else
                Call WriteCraftingResult(UserIndex, 0)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRemoveItemCrafting", Erl)
End Sub

Private Sub HandleAddCatalyst(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Slot As Byte
        Slot = reader.ReadInt8
        If .flags.Crafteando = 0 Then Exit Sub
        If Slot < 1 Or Slot > .CurrentInventorySlots Then Exit Sub
        If .invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        If ObjData(.invent.Object(Slot).ObjIndex).CatalizadorTipo = 0 Then Exit Sub
        If .CraftCatalyst.ObjIndex <> 0 Then Exit Sub
        .CraftCatalyst.ObjIndex = .invent.Object(Slot).ObjIndex
        .CraftCatalyst.amount = .invent.Object(Slot).amount
        Call QuitarUserInvItem(UserIndex, Slot, GetMaxInvOBJ())
        Call UpdateUserInv(False, UserIndex, Slot)
        If .CraftResult Is Nothing Then
            Call WriteCraftingCatalyst(UserIndex, .CraftCatalyst.ObjIndex, .CraftCatalyst.amount, 0)
        Else
            Call WriteCraftingCatalyst(UserIndex, .CraftCatalyst.ObjIndex, .CraftCatalyst.amount, CalculateCraftProb(UserIndex, .CraftResult.Probabilidad))
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleAddCatalyst", Erl)
End Sub

Private Sub HandleRemoveCatalyst(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Slot As Byte
        Slot = reader.ReadInt8
        If .flags.Crafteando = 0 Then Exit Sub
        If .CraftCatalyst.ObjIndex = 0 Then Exit Sub
        If Slot < 1 Then
            If Not MeterItemEnInventario(UserIndex, .CraftCatalyst) Then Exit Sub
        ElseIf Slot <= .CurrentInventorySlots Then
            If .invent.Object(Slot).ObjIndex = 0 Then
                .invent.Object(Slot).ObjIndex = .CraftCatalyst.ObjIndex
            ElseIf .invent.Object(Slot).ObjIndex <> .CraftCatalyst.ObjIndex Then
                Exit Sub
            End If
            .invent.Object(Slot).amount = .invent.Object(Slot).amount + .CraftCatalyst.amount
            Call UpdateUserInv(False, UserIndex, Slot)
        End If
        .CraftCatalyst.ObjIndex = 0
        .CraftCatalyst.amount = 0
        If .CraftResult Is Nothing Then
            Call WriteCraftingCatalyst(UserIndex, 0, 0, 0)
        Else
            Call WriteCraftingCatalyst(UserIndex, 0, 0, CalculateCraftProb(UserIndex, .CraftResult.Probabilidad))
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRemoveCatalyst", Erl)
End Sub

Sub HandleCraftItem(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    If UserList(UserIndex).flags.Crafteando = 0 Then Exit Sub
    Call DoCraftItem(UserIndex)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftItem", Erl)
End Sub

Private Sub HandleCloseCrafting(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    If UserList(UserIndex).flags.Crafteando = 0 Then Exit Sub
    Call ReturnCraftingItems(UserIndex)
    UserList(UserIndex).flags.Crafteando = 0
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCloseCrafting", Erl)
End Sub

Private Sub HandleMoveCraftItem(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Drag As Byte, Drop As Byte
        Drag = reader.ReadInt8
        Drop = reader.ReadInt8
        If .flags.Crafteando = 0 Then Exit Sub
        If Drag < 1 Or Drag > MAX_SLOTS_CRAFTEO Then Exit Sub
        If Drop < 1 Or Drop > MAX_SLOTS_CRAFTEO Then Exit Sub
        If Drag = Drop Then Exit Sub
        If .CraftInventory(Drag) = 0 Then Exit Sub
        If .CraftInventory(Drag) = .CraftInventory(Drop) Then Exit Sub
        Dim aux As Integer
        aux = .CraftInventory(Drop)
        .CraftInventory(Drop) = .CraftInventory(Drag)
        .CraftInventory(Drag) = aux
        Call WriteCraftingItem(UserIndex, Drag, .CraftInventory(Drag))
        Call WriteCraftingItem(UserIndex, Drop, .CraftInventory(Drop))
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleMoveCraftItem", Erl)
End Sub

Private Sub HandlePetLeaveAll(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim AlmenosUna As Boolean, i As Integer
        For i = 1 To MAXMASCOTAS
            If IsValidNpcRef(.MascotasIndex(i)) Then
                If NpcList(.MascotasIndex(i).ArrayIndex).flags.NPCActive Then
                    Call QuitarNPC(.MascotasIndex(i).ArrayIndex, e_DeleteSource.ePetLeave)
                    AlmenosUna = True
                End If
            End If
        Next i
        If AlmenosUna Then
            .flags.ModificoMascotas = True
            'Msg1278= Liberaste a tus mascotas.
            Call WriteLocaleMsg(UserIndex, 1278, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandlePetLeaveAll", Erl)
End Sub

Private Sub HandleResetChar(ByVal UserIndex As Integer)
    On Error GoTo HandleResetChar_Err:
    Dim Nick As String: Nick = reader.ReadString8()
    #If DEBUGGING = 1 Then
        If UserList(UserIndex).flags.Privilegios And e_PlayerType.Admin Then
            Dim User As t_UserReference
            User = NameIndex(Nick)
            If Not IsValidUserRef(User) Then
                'Msg1279= Usuario offline o inexistente.
                Call WriteLocaleMsg(UserIndex, 1279, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            With UserList(User.ArrayIndex)
                .Stats.ELV = 1
                .Stats.Exp = 0
                .Stats.SkillPts = 10
                Dim i As Integer
                For i = 1 To NUMSKILLS
                    .Stats.UserSkills(i) = 0
                Next
                .Stats.MaxAGU = 100
                .Stats.MinAGU = 100
                .Stats.MaxHam = 100
                .Stats.MinHam = 100
                .Stats.MaxHit = 2
                .Stats.MinHIT = 1
                .Stats.MaxMAN = UserMod.GetMaxMana(UserIndex)
                .Stats.MaxSta = UserMod.GetMaxStamina(UserIndex)
                .Stats.MaxHp = .Stats.UserAtributos(e_Atributos.Constitucion)
                .Stats.MinHp = .Stats.MaxHp
                .Stats.MinMAN = .Stats.MaxMAN
                .Stats.MinSta = .Stats.MaxSta
                .flags.ModificoSkills = True
                Call WriteUpdateUserStats(User.ArrayIndex)
                Call WriteLevelUp(User.ArrayIndex, .Stats.SkillPts)
            End With
            'Msg1280= Personaje reseteado a nivel 1.
            Call WriteLocaleMsg(UserIndex, 1280, e_FontTypeNames.FONTTYPE_INFO)
        End If
    #End If
    Exit Sub
HandleResetChar_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleResetChar", Erl)
End Sub

Private Sub HandleResetearPersonaje(ByVal UserIndex As Integer)
    On Error GoTo HandleResetearPersonaje_Err:
    ' Call resetPj(UserIndex)
    Exit Sub
HandleResetearPersonaje_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleResetearPersonaje", Erl)
End Sub

Private Sub HandleRomperCania(ByVal UserIndex As Integer)
    On Error GoTo HandleRomperCania_Err:
    Dim LoopC    As Integer
    Dim obj      As t_Obj
    Dim caniaOld As Integer
    Dim shouldBreak As Boolean

    With UserList(UserIndex)
        obj.ObjIndex = .invent.EquippedWorkingToolObjIndex
        caniaOld = .invent.EquippedWorkingToolObjIndex
        obj.amount = 1
        shouldBreak = (RandomNumber(1, 3) = 1)
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            'Rastreo la caña que está usando en el inventario y se la rompo
            If .invent.Object(LoopC).ObjIndex = .invent.EquippedWorkingToolObjIndex Then
                If caniaOld = OBJ_FISHING_NET_BASIC Or caniaOld = OBJ_FISHING_NET_ELITE Then
                    If shouldBreak Then
                        'Le quito una red
                        Call QuitarUserInvItem(UserIndex, LoopC, 1)
                        Call UpdateUserInv(False, UserIndex, LoopC)
                        Call WriteLocaleMsg(UserIndex, MSG_REMOVE_NET_LOST, e_FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteLocaleMsg(UserIndex, MSG_REMOVE_NET_ALMOST_LOST, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    If shouldBreak Then
                        'Le quito una caña
                        Call QuitarUserInvItem(UserIndex, LoopC, 1)
                        Call UpdateUserInv(False, UserIndex, LoopC)

                        Select Case caniaOld
                            Case OBJ_FISHING_ROD_BASIC
                                obj.ObjIndex = OBJ_BROKEN_FISHING_ROD_BASIC
                            Case OBJ_FISHING_ROD_COMMON
                                obj.ObjIndex = OBJ_BROKEN_FISHING_ROD_COMMON
                            Case OBJ_FISHING_ROD_FINE
                                obj.ObjIndex = OBJ_BROKEN_FISHING_ROD_FINE
                            Case OBJ_FISHING_ROD_ELITE
                                obj.ObjIndex = OBJ_BROKEN_FISHING_ROD_ELITE
                        End Select
                        Call MeterItemEnInventario(UserIndex, obj)
                    Else
                        Call WriteLocaleMsg(UserIndex, MSG_REMOVE_ALMOST_YOUR_FISHING, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                Exit Sub
            End If
        Next LoopC
    End With
    'UserList(UserIndex).Invent.EquippedWorkingToolObjIndex
HandleRomperCania_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRomperCania", Erl)
End Sub

Private Sub HandleFinalizarPescaEspecial(ByVal UserIndex As Integer)
    On Error GoTo HandleFinalizarPescaEspecial_Err:
    Call EntregarPezEspecial(UserIndex)
    Exit Sub
HandleFinalizarPescaEspecial_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleFinalizarPescaEspecial", Erl)
End Sub

Private Sub HandleRepeatMacro(ByVal UserIndex As Integer)
    On Error GoTo HandleRepeatMacro_Err:
    'Call LogMacroCliente("El usuario " & UserList(UserIndex).name & " iteró el paquete click o u." & GetTickCount)
    Exit Sub
HandleRepeatMacro_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRepeatMacro", Erl)
End Sub

Private Sub HandleBuyShopItem(ByVal UserIndex As Integer)
    On Error GoTo HandleBuyShopItem_Err:
    Dim obj_to_buy As Long
    obj_to_buy = reader.ReadInt32
    Call ModShopAO20.init_transaction(obj_to_buy, UserIndex)
    Exit Sub
HandleBuyShopItem_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleBuyShopItem", Erl)
End Sub

Private Sub HandlePublicarPersonajeMAO(ByVal UserIndex As Integer)
    On Error GoTo HandlePublicarPersonajeMAO_Err:
    Dim Valor As Long
    Valor = reader.ReadInt32
    If Valor <= MinimumPriceMao Then
        'Msg1281= El valor de venta del personaje debe ser mayor que $¬1
        Call WriteLocaleMsg(UserIndex, 1281, e_FontTypeNames.FONTTYPE_INFO, MinimumPriceMao)
        Exit Sub
    End If
    With UserList(UserIndex)
        ' Para recibir el ID del user
        Dim RS As ADODB.Recordset
        Set RS = Query("select is_locked_in_mao from user where id = ?;", .Id)
        If EsGM(UserIndex) Then
            'Msg1282= No podes vender un gm.
            Call WriteLocaleMsg(UserIndex, 1282, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If CBool(RS!is_locked_in_mao) Then
            'Msg1283= El personaje ya está publicado.
            Call WriteLocaleMsg(UserIndex, 1283, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .Stats.ELV < MinimumLevelMao Then
            'Msg1284= No puedes publicar un personaje menor a nivel ¬1
            Call WriteLocaleMsg(UserIndex, 1284, e_FontTypeNames.FONTTYPE_INFO, MinimumLevelMao)
            Exit Sub
        End If
        If .Stats.GLD < GoldPriceMao Then
            'Msg1291= El costo para vender tu personajes es de ¬1 monedas de oro, no tienes esa cantidad.
            Call WriteLocaleMsg(UserIndex, 1291, e_FontTypeNames.FONTTYPE_INFOBOLD, GoldPriceMao)
            Exit Sub
        Else
            .Stats.GLD = .Stats.GLD - GoldPriceMao
            Call WriteUpdateGold(UserIndex)
        End If
        Call Execute("update user set price_in_mao = ?, is_locked_in_mao = 1 where id = ?;", Valor, .Id)
        Call modNetwork.Kick(UserList(UserIndex).ConnectionDetails.ConnID, "El personaje fue publicado.")
    End With
    Exit Sub
HandlePublicarPersonajeMAO_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandlePublicarPersonajeMAO", Erl)
End Sub

Private Sub HandleDeleteItem(ByVal UserIndex As Integer)
    
Dim isSkin As Boolean
Dim Slot As Byte

    On Error GoTo HandleDeleteItem_Err:
    
    isSkin = reader.ReadBool
    Slot = reader.ReadInt8()
    
    With UserList(UserIndex)
        
        If Not isSkin Then
            If Slot > getMaxInventorySlots(UserIndex) Or Slot <= 0 Then Exit Sub
            If MapInfo(.pos.Map).Seguro = 0 Or EsMapaEvento(.pos.Map) Then
                'Msg1285= Solo puedes eliminar items en zona segura.
                Call WriteLocaleMsg(UserIndex, 1285, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If .flags.Muerto = 1 Then
                'Msg1286= No puede eliminar items cuando estas muerto.
                Call WriteLocaleMsg(UserIndex, 1286, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If .invent.Object(Slot).Equipped = 0 Then
                .invent.Object(Slot).amount = 0
                .invent.Object(Slot).Equipped = 0
                .invent.Object(Slot).ObjIndex = 0
                Call UpdateUserInv(False, UserIndex, Slot)
                'Msg1287= Objeto eliminado correctamente.
                Call WriteLocaleMsg(UserIndex, 1287, e_FontTypeNames.FONTTYPE_INFO)
            Else
                'Msg1288= No puedes eliminar un objeto estando equipado.
                Call WriteLocaleMsg(UserIndex, 1288, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            If Slot > MAX_SKINSINVENTORY_SLOTS Or Slot <= 0 Then Exit Sub
            
            If MapInfo(.pos.Map).Seguro = 0 Or EsMapaEvento(.pos.Map) Then
                'Msg1285= Solo puedes eliminar items en zona segura.
                Call WriteLocaleMsg(UserIndex, "1285", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .flags.Muerto = 1 Then
                'Msg1286= No puede eliminar items cuando estas muerto.
                Call WriteLocaleMsg(UserIndex, "1286", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .Invent_Skins.Object(Slot).Equipped = 0 Then
                Call LogShopTransactions("PJ ID: " & .Id & " Nick: " & .name & " -> Borró el Skin: " & ObjData(.Invent_Skins.Object(Slot).ObjIndex).name & " Tipo: " & ObjData(.Invent_Skins.Object(Slot).ObjIndex).ObjType & " Valor: " & ObjData(.Invent_Skins.Object(Slot).ObjIndex).Valor)
                Call DesequiparSkin(UserIndex, Slot)
                'Msg1287= Objeto eliminado correctamente.
                .Invent_Skins.Object(Slot).Deleted = True
                Call SaveUser(UserIndex, False)
                Call WriteChangeSkinSlot(UserIndex, 0, Slot)
                Call WriteLocaleMsg(UserIndex, "1287", e_FontTypeNames.FONTTYPE_INFO)
            Else
                'Msg1288= No puedes eliminar un objeto estando equipado.
                Call WriteLocaleMsg(UserIndex, "1288", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
    End With
    
    Exit Sub
HandleDeleteItem_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleDeleteItem", Erl)
End Sub

Public Sub HandleActionOnGroupFrame(ByVal UserIndex As Integer)
    On Error GoTo HandleActionOnGroupFrame_Err:
    Dim TargetGroupMember As Byte
    TargetGroupMember = reader.ReadInt8
    With UserList(UserIndex)
        If Not .Grupo.EnGrupo Then Exit Sub
        If Not IsFeatureEnabled("target_group_frames") Then Exit Sub
        If UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros < TargetGroupMember Then Exit Sub
        If Not IsValidUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(TargetGroupMember)) Then Exit Sub
        If UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(TargetGroupMember).ArrayIndex = UserIndex Then Exit Sub
        If UserMod.IsStun(.flags, .Counters) Then Exit Sub
        If .flags.Muerto = 1 Or .flags.Descansar Then Exit Sub
        Dim targetUserIndex As Integer
        targetUserIndex = UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(TargetGroupMember).ArrayIndex
        If Abs(.pos.Map <> UserList(targetUserIndex).pos.Map) Then Exit Sub
        If Abs(.pos.x - UserList(targetUserIndex).pos.x) > RANGO_VISION_X Or Abs(.pos.y - UserList(targetUserIndex).pos.y) > RANGO_VISION_Y Then Exit Sub
        If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
        If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
        If Not IntervaloPermiteLanzarSpell(UserIndex) Then Exit Sub
        If .flags.Meditando Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
        End If
        .flags.TargetUser = UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(TargetGroupMember)
        If .flags.Hechizo > 0 Then
            If Not IsSet(Hechizos(UserList(UserIndex).Stats.UserHechizos(.flags.Hechizo)).SpellRequirementMask, e_SpellRequirementMask.eIsBindable) Then
                Call WriteLocaleMsg(UserIndex, MsgBindableHotkeysOnly, e_FontTypeNames.FONTTYPE_INFO)
                Call WriteWorkRequestTarget(UserIndex, 0)
                Exit Sub
            End If
            .Counters.controlHechizos.HechizosTotales = .Counters.controlHechizos.HechizosTotales + 1
            Call LanzarHechizo(.flags.Hechizo, UserIndex)
            Call WriteWorkRequestTarget(UserIndex, 0)
            .flags.Hechizo = 0
        Else
            ' Msg587=¡Primero selecciona el hechizo que quieres lanzar!
            Call WriteLocaleMsg(UserIndex, 587, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleActionOnGroupFrame_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleActionOnGroupFrame UserId:" & UserIndex, Erl)
End Sub

Public Sub HandleSetHotkeySlot(ByVal UserIndex As Integer)
    On Error GoTo HandleSetHotkeySlot_Err:
    With UserList(UserIndex)
        Dim SlotIndex     As Byte
        Dim TargetIndex   As Integer
        Dim LastKnownSlot As Integer
        Dim HkType        As Byte
        SlotIndex = reader.ReadInt8
        TargetIndex = reader.ReadInt16
        LastKnownSlot = reader.ReadInt16
        HkType = reader.ReadInt8
        .HotkeyList(SlotIndex).Index = TargetIndex
        .HotkeyList(SlotIndex).LastKnownSlot = LastKnownSlot
        .HotkeyList(SlotIndex).Type = HkType
    End With
    Exit Sub
HandleSetHotkeySlot_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSetHotkeySlot", Erl)
End Sub

Public Sub HandleUseHKeySlot(ByVal UserIndex As Integer)
    On Error GoTo HandleUseHKeySlot_Err:
    Dim SlotIndex As Byte
    SlotIndex = reader.ReadInt8
    If Not IsFeatureEnabled("hotokey-enabled") Then Exit Sub
    Dim CurrentSlotIndex As Integer
    Dim i                As Integer
    With UserList(UserIndex)
        If .HotkeyList(SlotIndex).Index > 0 Then
            If .HotkeyList(SlotIndex).Type = Item Then
            ElseIf .HotkeyList(SlotIndex).Type = Spell Then
                If .HotkeyList(SlotIndex).LastKnownSlot > 0 And .HotkeyList(SlotIndex).LastKnownSlot < UBound(.Stats.UserHechizos) Then
                    If .Stats.UserHechizos(.HotkeyList(SlotIndex).LastKnownSlot) = .HotkeyList(SlotIndex).Index Then
                        CurrentSlotIndex = .HotkeyList(SlotIndex).LastKnownSlot
                    End If
                End If
                If CurrentSlotIndex = 0 Then
                    For i = LBound(.Stats.UserHechizos) To UBound(.Stats.UserHechizos)
                        If .Stats.UserHechizos(i) = .HotkeyList(SlotIndex).Index Then
                            CurrentSlotIndex = i
                            Exit For
                        End If
                    Next i
                End If
                If CurrentSlotIndex > 0 Then
                    If .Stats.UserHechizos(CurrentSlotIndex) > 0 Then
                        If IsSet(Hechizos(UserList(UserIndex).Stats.UserHechizos(CurrentSlotIndex)).SpellRequirementMask, e_SpellRequirementMask.eIsBindable) Then
                            Call UseSpellSlot(UserIndex, CurrentSlotIndex)
                        End If
                    End If
                End If
            End If
        End If
    End With
    Exit Sub
HandleUseHKeySlot_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleUseHKeySlot", Erl)
End Sub

Public Sub HandleAntiCheatMessage(ByVal UserIndex As Integer)
    On Error GoTo AntiCheatMessage_Err:
    Dim data() As Byte
    Call reader.ReadSafeArrayInt8(data)
    Call HandleAntiCheatServerMessage(UserIndex, data)
    Exit Sub
AntiCheatMessage_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.AntiCheatMessage", Erl)
End Sub

Public Sub HendleRequestLobbyList(ByVal UserIndex As Integer)
    On Error GoTo HendleRequestLobbyList_Err:
    Call WriteUpdateLobbyList(UserIndex)
    Exit Sub
HendleRequestLobbyList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HendleRequestLobbyList", Erl)
End Sub
Public Function IsInMapCarcelRestrictedArea(ByRef position As t_WorldPos) As Boolean
    If position.Map <> MAP_HOME_IN_JAIL Then Exit Function

    If position.x >= 33 And position.x <= 62 And position.y >= 32 And position.y <= 62 Then
        IsInMapCarcelRestrictedArea = True
    End If
End Function

Public Function HandleStartAutomatedAction(ByVal UserIndex As Integer)
    On Error GoTo HandleStartAutomatedAction_Err
    Dim x     As Byte
    Dim y     As Byte
    Dim skill As e_Skill
    x = reader.ReadInt8()
    y = reader.ReadInt8()
    skill = reader.ReadInt8()
    Select Case skill
        Case e_Skill.Pescar
            If Not CanUserFish(UserIndex, x, y) Then
                Exit Function
            End If
        Case e_Skill.Talar
            If Not CanUserExtractResource(UserIndex, e_OBJType.otTrees, x, y) Then
                Exit Function
            End If
        Case e_Skill.Mineria
            If Not CanUserExtractResource(UserIndex, e_OBJType.otOreDeposit, x, y) Then
                Exit Function
            End If
        Case Else
    End Select
    Call StartAutomatedAction(x, y, skill, UserIndex)
    Exit Function
HandleStartAutomatedAction_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleStartAutomatedAction", Erl)
End Function
