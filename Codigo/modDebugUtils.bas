Attribute VB_Name = "modDebugUtils"
' Argentum 20 Game Server
'
'    Copyright (C) 2025 Noland Studios LTD
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
'    Copyright (C) 2002 Mrquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
Option Explicit

Public Function PacketID_to_string(ByVal PacketId As ClientPacketID) As String
    Select Case PacketId
        Case ClientPacketID.eLoginExistingChar
            PacketID_to_string = "eLoginExistingChar"
        Case ClientPacketID.eLoginNewChar
            PacketID_to_string = "eLoginNewChar"
        Case ClientPacketID.eWalk
            PacketID_to_string = "eWalk"
        Case ClientPacketID.eAttack
            PacketID_to_string = "eAttack"
        Case ClientPacketID.eTalk
            PacketID_to_string = "eTalk"
        Case ClientPacketID.eYell
            PacketID_to_string = "eYell"
        Case ClientPacketID.eWhisper
            PacketID_to_string = "eWhisper"
        Case ClientPacketID.eRequestPositionUpdate
            PacketID_to_string = "eRequestPositionUpdate"
        Case ClientPacketID.ePickUp
            PacketID_to_string = "ePickUp"
        Case ClientPacketID.eSafeToggle
            PacketID_to_string = "eSafeToggle"
        Case ClientPacketID.ePartySafeToggle
            PacketID_to_string = "ePartySafeToggle"
        Case ClientPacketID.eRequestGuildLeaderInfo
            PacketID_to_string = "eRequestGuildLeaderInfo"
        Case ClientPacketID.eRequestAtributes
            PacketID_to_string = "eRequestAtributes"
        Case ClientPacketID.eRequestSkills
            PacketID_to_string = "eRequestSkills"
        Case ClientPacketID.eRequestMiniStats
            PacketID_to_string = "eRequestMiniStats"
        Case ClientPacketID.eCommerceEnd
            PacketID_to_string = "eCommerceEnd"
        Case ClientPacketID.eUserCommerceEnd
            PacketID_to_string = "eUserCommerceEnd"
        Case ClientPacketID.eBankEnd
            PacketID_to_string = "eBankEnd"
        Case ClientPacketID.eUserCommerceOk
            PacketID_to_string = "eUserCommerceOk"
        Case ClientPacketID.eUserCommerceReject
            PacketID_to_string = "eUserCommerceReject"
        Case ClientPacketID.eDrop
            PacketID_to_string = "eDrop"
        Case ClientPacketID.eCastSpell
            PacketID_to_string = "eCastSpell"
        Case ClientPacketID.eLeftClick
            PacketID_to_string = "eLeftClick"
        Case ClientPacketID.eDoubleClick
            PacketID_to_string = "eDoubleClick"
        Case ClientPacketID.eWork
            PacketID_to_string = "eWork"
        Case ClientPacketID.eUseSpellMacro
            PacketID_to_string = "eUseSpellMacro"
        Case ClientPacketID.eUseItem
            PacketID_to_string = "eUseItem"
        Case ClientPacketID.eUseItemU
            PacketID_to_string = "eUseItemU"
        Case ClientPacketID.eCraftBlacksmith
            PacketID_to_string = "eCraftBlacksmith"
        Case ClientPacketID.eCraftCarpenter
            PacketID_to_string = "eCraftCarpenter"
        Case ClientPacketID.eWorkLeftClick
            PacketID_to_string = "eWorkLeftClick"
        Case ClientPacketID.eStartAutomatedAction
            PacketID_to_string = "eStartAutomatedAction"
        Case ClientPacketID.eCreateNewGuild
            PacketID_to_string = "eCreateNewGuild"
        Case ClientPacketID.eSpellInfo
            PacketID_to_string = "eSpellInfo"
        Case ClientPacketID.eEquipItem
            PacketID_to_string = "eEquipItem"
        Case ClientPacketID.eChangeHeading
            PacketID_to_string = "eChangeHeading"
        Case ClientPacketID.eModifySkills
            PacketID_to_string = "eModifySkills"
        Case ClientPacketID.eTrain
            PacketID_to_string = "eTrain"
        Case ClientPacketID.eCommerceBuy
            PacketID_to_string = "eCommerceBuy"
        Case ClientPacketID.eBankExtractItem
            PacketID_to_string = "eBankExtractItem"
        Case ClientPacketID.eCommerceSell
            PacketID_to_string = "eCommerceSell"
        Case ClientPacketID.eBankDeposit
            PacketID_to_string = "eBankDeposit"
        Case ClientPacketID.eForumPost
            PacketID_to_string = "eForumPost"
        Case ClientPacketID.eMoveSpell
            PacketID_to_string = "eMoveSpell"
        Case ClientPacketID.eClanCodexUpdate
            PacketID_to_string = "eClanCodexUpdate"
        Case ClientPacketID.eUserCommerceOffer
            PacketID_to_string = "eUserCommerceOffer"
        Case ClientPacketID.eGuildAcceptPeace
            PacketID_to_string = "eGuildAcceptPeace"
        Case ClientPacketID.eGuildRejectAlliance
            PacketID_to_string = "eGuildRejectAlliance"
        Case ClientPacketID.eGuildRejectPeace
            PacketID_to_string = "eGuildRejectPeace"
        Case ClientPacketID.eGuildAcceptAlliance
            PacketID_to_string = "eGuildAcceptAlliance"
        Case ClientPacketID.eGuildOfferPeace
            PacketID_to_string = "eGuildOfferPeace"
        Case ClientPacketID.eGuildOfferAlliance
            PacketID_to_string = "eGuildOfferAlliance"
        Case ClientPacketID.eGuildAllianceDetails
            PacketID_to_string = "eGuildAllianceDetails"
        Case ClientPacketID.eGuildPeaceDetails
            PacketID_to_string = "eGuildPeaceDetails"
        Case ClientPacketID.eGuildRequestJoinerInfo
            PacketID_to_string = "eGuildRequestJoinerInfo"
        Case ClientPacketID.eGuildAlliancePropList
            PacketID_to_string = "eGuildAlliancePropList"
        Case ClientPacketID.eGuildPeacePropList
            PacketID_to_string = "eGuildPeacePropList"
        Case ClientPacketID.eGuildDeclareWar
            PacketID_to_string = "eGuildDeclareWar"
        Case ClientPacketID.eGuildNewWebsite
            PacketID_to_string = "eGuildNewWebsite"
        Case ClientPacketID.eGuildAcceptNewMember
            PacketID_to_string = "eGuildAcceptNewMember"
        Case ClientPacketID.eGuildRejectNewMember
            PacketID_to_string = "eGuildRejectNewMember"
        Case ClientPacketID.eGuildKickMember
            PacketID_to_string = "eGuildKickMember"
        Case ClientPacketID.eGuildUpdateNews
            PacketID_to_string = "eGuildUpdateNews"
        Case ClientPacketID.eGuildMemberInfo
            PacketID_to_string = "eGuildMemberInfo"
        Case ClientPacketID.eGuildOpenElections
            PacketID_to_string = "eGuildOpenElections"
        Case ClientPacketID.eGuildRequestMembership
            PacketID_to_string = "eGuildRequestMembership"
        Case ClientPacketID.eGuildRequestDetails
            PacketID_to_string = "eGuildRequestDetails"
        Case ClientPacketID.eOnline
            PacketID_to_string = "eOnline"
        Case ClientPacketID.eQuit
            PacketID_to_string = "eQuit"
        Case ClientPacketID.eGuildLeave
            PacketID_to_string = "eGuildLeave"
        Case ClientPacketID.eRequestAccountState
            PacketID_to_string = "eRequestAccountState"
        Case ClientPacketID.ePetStand
            PacketID_to_string = "ePetStand"
        Case ClientPacketID.ePetFollow
            PacketID_to_string = "ePetFollow"
        Case ClientPacketID.ePetFollowAll
            PacketID_to_string = "ePetFollowAll"
        Case ClientPacketID.ePetLeave
            PacketID_to_string = "ePetLeave"
        Case ClientPacketID.eGrupoMsg
            PacketID_to_string = "eGrupoMsg"
        Case ClientPacketID.eTrainList
            PacketID_to_string = "eTrainList"
        Case ClientPacketID.eRest
            PacketID_to_string = "eRest"
        Case ClientPacketID.eMeditate
            PacketID_to_string = "eMeditate"
        Case ClientPacketID.eResucitate
            PacketID_to_string = "eResucitate"
        Case ClientPacketID.eHeal
            PacketID_to_string = "eHeal"
        Case ClientPacketID.eHelp
            PacketID_to_string = "eHelp"
        Case ClientPacketID.eRequestStats
            PacketID_to_string = "eRequestStats"
        Case ClientPacketID.eCommerceStart
            PacketID_to_string = "eCommerceStart"
        Case ClientPacketID.eBankStart
            PacketID_to_string = "eBankStart"
        Case ClientPacketID.eEnlist
            PacketID_to_string = "eEnlist"
        Case ClientPacketID.eInformation
            PacketID_to_string = "eInformation"
        Case ClientPacketID.eReward
            PacketID_to_string = "eReward"
        Case ClientPacketID.eRequestMOTD
            PacketID_to_string = "eRequestMOTD"
        Case ClientPacketID.eUpTime
            PacketID_to_string = "eUpTime"
        Case ClientPacketID.eGuildMessage
            PacketID_to_string = "eGuildMessage"
        Case ClientPacketID.eGuildOnline
            PacketID_to_string = "eGuildOnline"
        Case ClientPacketID.eCouncilMessage
            PacketID_to_string = "eCouncilMessage"
        Case ClientPacketID.eRoleMasterRequest
            PacketID_to_string = "eRoleMasterRequest"
        Case ClientPacketID.eChangeDescription
            PacketID_to_string = "eChangeDescription"
        Case ClientPacketID.eGuildVote
            PacketID_to_string = "eGuildVote"
        Case ClientPacketID.epunishments
            PacketID_to_string = "epunishments"
        Case ClientPacketID.eGamble
            PacketID_to_string = "eGamble"
        Case ClientPacketID.eMapPriceEntrance
            PacketID_to_string = "eMapPriceEntrance"
        Case ClientPacketID.eLeaveFaction
            PacketID_to_string = "eLeaveFaction"
        Case ClientPacketID.eBankExtractGold
            PacketID_to_string = "eBankExtractGold"
        Case ClientPacketID.eBankDepositGold
            PacketID_to_string = "eBankDepositGold"
        Case ClientPacketID.eDenounce
            PacketID_to_string = "eDenounce"
        Case ClientPacketID.eGMMessage
            PacketID_to_string = "eGMMessage"
        Case ClientPacketID.eshowName
            PacketID_to_string = "eshowName"
        Case ClientPacketID.eOnlineRoyalArmy
            PacketID_to_string = "eOnlineRoyalArmy"
        Case ClientPacketID.eOnlineChaosLegion
            PacketID_to_string = "eOnlineChaosLegion"
        Case ClientPacketID.eGoNearby
            PacketID_to_string = "eGoNearby"
        Case ClientPacketID.ecomment
            PacketID_to_string = "ecomment"
        Case ClientPacketID.eWhere
            PacketID_to_string = "eWhere"
        Case ClientPacketID.eCreaturesInMap
            PacketID_to_string = "eCreaturesInMap"
        Case ClientPacketID.eWarpMeToTarget
            PacketID_to_string = "eWarpMeToTarget"
        Case ClientPacketID.eWarpChar
            PacketID_to_string = "eWarpChar"
        Case ClientPacketID.eSilence
            PacketID_to_string = "eSilence"
        Case ClientPacketID.eSOSShowList
            PacketID_to_string = "eSOSShowList"
        Case ClientPacketID.eSOSRemove
            PacketID_to_string = "eSOSRemove"
        Case ClientPacketID.eGoToChar
            PacketID_to_string = "eGoToChar"
        Case ClientPacketID.einvisible
            PacketID_to_string = "einvisible"
        Case ClientPacketID.eGMPanel
            PacketID_to_string = "eGMPanel"
        Case ClientPacketID.eRequestUserList
            PacketID_to_string = "eRequestUserList"
        Case ClientPacketID.eWorking
            PacketID_to_string = "eWorking"
        Case ClientPacketID.eHiding
            PacketID_to_string = "eHiding"
        Case ClientPacketID.eJail
            PacketID_to_string = "eJail"
        Case ClientPacketID.eKillNPC
            PacketID_to_string = "eKillNPC"
        Case ClientPacketID.eWarnUser
            PacketID_to_string = "eWarnUser"
        Case ClientPacketID.eEditChar
            PacketID_to_string = "eEditChar"
        Case ClientPacketID.eRequestCharInfo
            PacketID_to_string = "eRequestCharInfo"
        Case ClientPacketID.eRequestCharStats
            PacketID_to_string = "eRequestCharStats"
        Case ClientPacketID.eRequestCharGold
            PacketID_to_string = "eRequestCharGold"
        Case ClientPacketID.eRequestCharInventory
            PacketID_to_string = "eRequestCharInventory"
        Case ClientPacketID.eRequestCharBank
            PacketID_to_string = "eRequestCharBank"
        Case ClientPacketID.eRequestCharSkills
            PacketID_to_string = "eRequestCharSkills"
        Case ClientPacketID.eReviveChar
            PacketID_to_string = "eReviveChar"
        Case ClientPacketID.eNotifyInventarioHechizos
            PacketID_to_string = "eNotifyInventarioHechizos"
        Case ClientPacketID.eOnlineGM
            PacketID_to_string = "eOnlineGM"
        Case ClientPacketID.eOnlineMap
            PacketID_to_string = "eOnlineMap"
        Case ClientPacketID.eForgive
            PacketID_to_string = "eForgive"
        Case ClientPacketID.ePerdonFaccion
            PacketID_to_string = "ePerdonFaccion"
        Case ClientPacketID.eStartEvent
            PacketID_to_string = "eStartEvent"
        Case ClientPacketID.eCancelarEvento
            PacketID_to_string = "eCancelarEvento"
        Case ClientPacketID.eKick
            PacketID_to_string = "eKick"
        Case ClientPacketID.eExecute
            PacketID_to_string = "eExecute"
        Case ClientPacketID.eBanChar
            PacketID_to_string = "eBanChar"
        Case ClientPacketID.eUnbanChar
            PacketID_to_string = "eUnbanChar"
        Case ClientPacketID.eNPCFollow
            PacketID_to_string = "eNPCFollow"
        Case ClientPacketID.eSummonChar
            PacketID_to_string = "eSummonChar"
        Case ClientPacketID.eSpawnListRequest
            PacketID_to_string = "eSpawnListRequest"
        Case ClientPacketID.eSpawnCreature
            PacketID_to_string = "eSpawnCreature"
        Case ClientPacketID.eResetNPCInventory
            PacketID_to_string = "eResetNPCInventory"
        Case ClientPacketID.eCleanWorld
            PacketID_to_string = "eCleanWorld"
        Case ClientPacketID.eServerMessage
            PacketID_to_string = "eServerMessage"
        Case ClientPacketID.eNickToIP
            PacketID_to_string = "eNickToIP"
        Case ClientPacketID.eIPToNick
            PacketID_to_string = "eIPToNick"
        Case ClientPacketID.eGuildOnlineMembers
            PacketID_to_string = "eGuildOnlineMembers"
        Case ClientPacketID.eTeleportCreate
            PacketID_to_string = "eTeleportCreate"
        Case ClientPacketID.eTeleportDestroy
            PacketID_to_string = "eTeleportDestroy"
        Case ClientPacketID.eRainToggle
            PacketID_to_string = "eRainToggle"
        Case ClientPacketID.eSetCharDescription
            PacketID_to_string = "eSetCharDescription"
        Case ClientPacketID.eForceMIDIToMap
            PacketID_to_string = "eForceMIDIToMap"
        Case ClientPacketID.eForceWAVEToMap
            PacketID_to_string = "eForceWAVEToMap"
        Case ClientPacketID.eRoyalArmyMessage
            PacketID_to_string = "eRoyalArmyMessage"
        Case ClientPacketID.eChaosLegionMessage
            PacketID_to_string = "eChaosLegionMessage"
        Case ClientPacketID.eTalkAsNPC
            PacketID_to_string = "eTalkAsNPC"
        Case ClientPacketID.eDestroyAllItemsInArea
            PacketID_to_string = "eDestroyAllItemsInArea"
        Case ClientPacketID.eAcceptRoyalCouncilMember
            PacketID_to_string = "eAcceptRoyalCouncilMember"
        Case ClientPacketID.eAcceptChaosCouncilMember
            PacketID_to_string = "eAcceptChaosCouncilMember"
        Case ClientPacketID.eItemsInTheFloor
            PacketID_to_string = "eItemsInTheFloor"
        Case ClientPacketID.eMakeDumb
            PacketID_to_string = "eMakeDumb"
        Case ClientPacketID.eMakeDumbNoMore
            PacketID_to_string = "eMakeDumbNoMore"
        Case ClientPacketID.eCouncilKick
            PacketID_to_string = "eCouncilKick"
        Case ClientPacketID.eSetTrigger
            PacketID_to_string = "eSetTrigger"
        Case ClientPacketID.eAskTrigger
            PacketID_to_string = "eAskTrigger"
        Case ClientPacketID.eGuildMemberList
            PacketID_to_string = "eGuildMemberList"
        Case ClientPacketID.eGuildBan
            PacketID_to_string = "eGuildBan"
        Case ClientPacketID.eCreateItem
            PacketID_to_string = "eCreateItem"
        Case ClientPacketID.eDestroyItems
            PacketID_to_string = "eDestroyItems"
        Case ClientPacketID.eChaosLegionKick
            PacketID_to_string = "eChaosLegionKick"
        Case ClientPacketID.eRoyalArmyKick
            PacketID_to_string = "eRoyalArmyKick"
        Case ClientPacketID.eForceMIDIAll
            PacketID_to_string = "eForceMIDIAll"
        Case ClientPacketID.eForceWAVEAll
            PacketID_to_string = "eForceWAVEAll"
        Case ClientPacketID.eRemovePunishment
            PacketID_to_string = "eRemovePunishment"
        Case ClientPacketID.eTileBlockedToggle
            PacketID_to_string = "eTileBlockedToggle"
        Case ClientPacketID.eKillNPCNoRespawn
            PacketID_to_string = "eKillNPCNoRespawn"
        Case ClientPacketID.eKillAllNearbyNPCs
            PacketID_to_string = "eKillAllNearbyNPCs"
        Case ClientPacketID.eLastIP
            PacketID_to_string = "eLastIP"
        Case ClientPacketID.eChangeMOTD
            PacketID_to_string = "eChangeMOTD"
        Case ClientPacketID.eSetMOTD
            PacketID_to_string = "eSetMOTD"
        Case ClientPacketID.eSystemMessage
            PacketID_to_string = "eSystemMessage"
        Case ClientPacketID.eCreateNPC
            PacketID_to_string = "eCreateNPC"
        Case ClientPacketID.eCreateNPCWithRespawn
            PacketID_to_string = "eCreateNPCWithRespawn"
        Case ClientPacketID.eImperialArmour
            PacketID_to_string = "eImperialArmour"
        Case ClientPacketID.eChaosArmour
            PacketID_to_string = "eChaosArmour"
        Case ClientPacketID.eNavigateToggle
            PacketID_to_string = "eNavigateToggle"
        Case ClientPacketID.eServerOpenToUsersToggle
            PacketID_to_string = "eServerOpenToUsersToggle"
        Case ClientPacketID.eParticipar
            PacketID_to_string = "eParticipar"
        Case ClientPacketID.eTurnCriminal
            PacketID_to_string = "eTurnCriminal"
        Case ClientPacketID.eResetFactions
            PacketID_to_string = "eResetFactions"
        Case ClientPacketID.eRemoveCharFromGuild
            PacketID_to_string = "eRemoveCharFromGuild"
        Case ClientPacketID.eAlterName
            PacketID_to_string = "eAlterName"
        Case ClientPacketID.eDoBackUp
            PacketID_to_string = "eDoBackUp"
        Case ClientPacketID.eShowGuildMessages
            PacketID_to_string = "eShowGuildMessages"
        Case ClientPacketID.eChangeMapInfoPK
            PacketID_to_string = "eChangeMapInfoPK"
        Case ClientPacketID.eChangeMapInfoBackup
            PacketID_to_string = "eChangeMapInfoBackup"
        Case ClientPacketID.eChangeMapInfoRestricted
            PacketID_to_string = "eChangeMapInfoRestricted"
        Case ClientPacketID.eChangeMapInfoNoMagic
            PacketID_to_string = "eChangeMapInfoNoMagic"
        Case ClientPacketID.eChangeMapInfoNoInvi
            PacketID_to_string = "eChangeMapInfoNoInvi"
        Case ClientPacketID.eChangeMapInfoNoResu
            PacketID_to_string = "eChangeMapInfoNoResu"
        Case ClientPacketID.eChangeMapInfoLand
            PacketID_to_string = "eChangeMapInfoLand"
        Case ClientPacketID.eChangeMapInfoZone
            PacketID_to_string = "eChangeMapInfoZone"
        Case ClientPacketID.eChangeMapSetting
            PacketID_to_string = "eChangeMapSetting"
        Case ClientPacketID.eSaveChars
            PacketID_to_string = "eSaveChars"
        Case ClientPacketID.eCleanSOS
            PacketID_to_string = "eCleanSOS"
        Case ClientPacketID.eShowServerForm
            PacketID_to_string = "eShowServerForm"
        Case ClientPacketID.eKickAllChars
            PacketID_to_string = "eKickAllChars"
        Case ClientPacketID.eChatColor
            PacketID_to_string = "eChatColor"
        Case ClientPacketID.eIgnored
            PacketID_to_string = "eIgnored"
        Case ClientPacketID.eCheckSlot
            PacketID_to_string = "eCheckSlot"
        Case ClientPacketID.eSetSpeed
            PacketID_to_string = "eSetSpeed"
        Case ClientPacketID.eGlobalMessage
            PacketID_to_string = "eGlobalMessage"
        Case ClientPacketID.eGlobalOnOff
            PacketID_to_string = "eGlobalOnOff"
        Case ClientPacketID.eUseKey
            PacketID_to_string = "eUseKey"
        Case ClientPacketID.eDonateGold
            PacketID_to_string = "eDonateGold"
        Case ClientPacketID.ePromedio
            PacketID_to_string = "ePromedio"
        Case ClientPacketID.eGiveItem
            PacketID_to_string = "eGiveItem"
        Case ClientPacketID.eOfertaInicial
            PacketID_to_string = "eOfertaInicial"
        Case ClientPacketID.eOfertaDeSubasta
            PacketID_to_string = "eOfertaDeSubasta"
        Case ClientPacketID.eQuestionGM
            PacketID_to_string = "eQuestionGM"
        Case ClientPacketID.eCuentaRegresiva
            PacketID_to_string = "eCuentaRegresiva"
        Case ClientPacketID.ePossUser
            PacketID_to_string = "ePossUser"
        Case ClientPacketID.eDuel
            PacketID_to_string = "eDuel"
        Case ClientPacketID.eAcceptDuel
            PacketID_to_string = "eAcceptDuel"
        Case ClientPacketID.eCancelDuel
            PacketID_to_string = "eCancelDuel"
        Case ClientPacketID.eQuitDuel
            PacketID_to_string = "eQuitDuel"
        Case ClientPacketID.eNieveToggle
            PacketID_to_string = "eNieveToggle"
        Case ClientPacketID.eNieblaToggle
            PacketID_to_string = "eNieblaToggle"
        Case ClientPacketID.eTransFerGold
            PacketID_to_string = "eTransFerGold"
        Case ClientPacketID.eMoveitem
            PacketID_to_string = "eMoveitem"
        Case ClientPacketID.eGenio
            PacketID_to_string = "eGenio"
        Case ClientPacketID.eCasarse
            PacketID_to_string = "eCasarse"
        Case ClientPacketID.eCraftAlquimista
            PacketID_to_string = "eCraftAlquimista"
        Case ClientPacketID.eFlagTrabajar
            PacketID_to_string = "eFlagTrabajar"
        Case ClientPacketID.eCraftSastre
            PacketID_to_string = "eCraftSastre"
        Case ClientPacketID.eMensajeUser
            PacketID_to_string = "eMensajeUser"
        Case ClientPacketID.eTraerBoveda
            PacketID_to_string = "eTraerBoveda"
        Case ClientPacketID.eCompletarAccion
            PacketID_to_string = "eCompletarAccion"
        Case ClientPacketID.eInvitarGrupo
            PacketID_to_string = "eInvitarGrupo"
        Case ClientPacketID.eResponderPregunta
            PacketID_to_string = "eResponderPregunta"
        Case ClientPacketID.eRequestGrupo
            PacketID_to_string = "eRequestGrupo"
        Case ClientPacketID.eAbandonarGrupo
            PacketID_to_string = "eAbandonarGrupo"
        Case ClientPacketID.eHecharDeGrupo
            PacketID_to_string = "eHecharDeGrupo"
        Case ClientPacketID.eMacroPossent
            PacketID_to_string = "eMacroPossent"
        Case ClientPacketID.eSubastaInfo
            PacketID_to_string = "eSubastaInfo"
        Case ClientPacketID.eBanCuenta
            PacketID_to_string = "eBanCuenta"
        Case ClientPacketID.eUnbanCuenta
            PacketID_to_string = "eUnbanCuenta"
        Case ClientPacketID.eCerrarCliente
            PacketID_to_string = "eCerrarCliente"
        Case ClientPacketID.eEventoInfo
            PacketID_to_string = "eEventoInfo"
        Case ClientPacketID.eCrearEvento
            PacketID_to_string = "eCrearEvento"
        Case ClientPacketID.eBanTemporal
            PacketID_to_string = "eBanTemporal"
        Case ClientPacketID.eCancelarExit
            PacketID_to_string = "eCancelarExit"
        Case ClientPacketID.eCrearTorneo
            PacketID_to_string = "eCrearTorneo"
        Case ClientPacketID.eComenzarTorneo
            PacketID_to_string = "eComenzarTorneo"
        Case ClientPacketID.eCancelarTorneo
            PacketID_to_string = "eCancelarTorneo"
        Case ClientPacketID.eBusquedaTesoro
            PacketID_to_string = "eBusquedaTesoro"
        Case ClientPacketID.eCompletarViaje
            PacketID_to_string = "eCompletarViaje"
        Case ClientPacketID.eBovedaMoveItem
            PacketID_to_string = "eBovedaMoveItem"
        Case ClientPacketID.eQuieroFundarClan
            PacketID_to_string = "eQuieroFundarClan"
        Case ClientPacketID.ellamadadeclan
            PacketID_to_string = "ellamadadeclan"
        Case ClientPacketID.eMarcaDeClanPack
            PacketID_to_string = "eMarcaDeClanPack"
        Case ClientPacketID.eMarcaDeGMPack
            PacketID_to_string = "eMarcaDeGMPack"
        Case ClientPacketID.eQuest
            PacketID_to_string = "eQuest"
        Case ClientPacketID.eQuestAccept
            PacketID_to_string = "eQuestAccept"
        Case ClientPacketID.eQuestListRequest
            PacketID_to_string = "eQuestListRequest"
        Case ClientPacketID.eQuestDetailsRequest
            PacketID_to_string = "eQuestDetailsRequest"
        Case ClientPacketID.eQuestAbandon
            PacketID_to_string = "eQuestAbandon"
        Case ClientPacketID.eSeguroClan
            PacketID_to_string = "eSeguroClan"
        Case ClientPacketID.ehome
            PacketID_to_string = "ehome"
        Case ClientPacketID.eConsulta
            PacketID_to_string = "eConsulta"
        Case ClientPacketID.eGetMapInfo
            PacketID_to_string = "eGetMapInfo"
        Case ClientPacketID.eFinEvento
            PacketID_to_string = "eFinEvento"
        Case ClientPacketID.eSeguroResu
            PacketID_to_string = "eSeguroResu"
        Case ClientPacketID.eLegionarySecure
            PacketID_to_string = "eLegionarySecure"
        Case ClientPacketID.eCuentaExtractItem
            PacketID_to_string = "eCuentaExtractItem"
        Case ClientPacketID.eCuentaDeposit
            PacketID_to_string = "eCuentaDeposit"
        Case ClientPacketID.eCreateEvent
            PacketID_to_string = "eCreateEvent"
        Case ClientPacketID.eCommerceSendChatMessage
            PacketID_to_string = "eCommerceSendChatMessage"
        Case ClientPacketID.eLogMacroClickHechizo
            PacketID_to_string = "eLogMacroClickHechizo"
        Case ClientPacketID.eAddItemCrafting
            PacketID_to_string = "eAddItemCrafting"
        Case ClientPacketID.eRemoveItemCrafting
            PacketID_to_string = "eRemoveItemCrafting"
        Case ClientPacketID.eAddCatalyst
            PacketID_to_string = "eAddCatalyst"
        Case ClientPacketID.eRemoveCatalyst
            PacketID_to_string = "eRemoveCatalyst"
        Case ClientPacketID.eCraftItem
            PacketID_to_string = "eCraftItem"
        Case ClientPacketID.eCloseCrafting
            PacketID_to_string = "eCloseCrafting"
        Case ClientPacketID.eMoveCraftItem
            PacketID_to_string = "eMoveCraftItem"
        Case ClientPacketID.ePetLeaveAll
            PacketID_to_string = "ePetLeaveAll"
        Case ClientPacketID.eResetChar
            PacketID_to_string = "eResetChar"
        Case ClientPacketID.eResetearPersonaje
            PacketID_to_string = "eResetearPersonaje"
        Case ClientPacketID.eDeleteItem
            PacketID_to_string = "eDeleteItem"
        Case ClientPacketID.eFinalizarPescaEspecial
            PacketID_to_string = "eFinalizarPescaEspecial"
        Case ClientPacketID.eRomperCania
            PacketID_to_string = "eRomperCania"
        Case ClientPacketID.eRepeatMacro
            PacketID_to_string = "eRepeatMacro"
        Case ClientPacketID.eBuyShopItem
            PacketID_to_string = "eBuyShopItem"
        Case ClientPacketID.ePublicarPersonajeMAO
            PacketID_to_string = "ePublicarPersonajeMAO"
        Case ClientPacketID.eEventoFaccionario
            PacketID_to_string = "eEventoFaccionario"
        Case ClientPacketID.eRequestDebug
            PacketID_to_string = "eRequestDebug"
        Case ClientPacketID.eLobbyCommand
            PacketID_to_string = "eLobbyCommand"
        Case ClientPacketID.eFeatureToggle
            PacketID_to_string = "eFeatureToggle"
        Case ClientPacketID.eActionOnGroupFrame
            PacketID_to_string = "eActionOnGroupFrame"
        Case ClientPacketID.eSetHotkeySlot
            PacketID_to_string = "eSetHotkeySlot"
        Case ClientPacketID.eUseHKeySlot
            PacketID_to_string = "eUseHKeySlot"
        Case ClientPacketID.eAntiCheatMessage
            PacketID_to_string = "eAntiCheatMessage"
        Case ClientPacketID.eFactionMessage
            PacketID_to_string = "eFactionMessage"
        Case ClientPacketID.eCreateAccount
            PacketID_to_string = "eCreateAccount"
        Case ClientPacketID.eLoginAccount
            PacketID_to_string = "eLoginAccount"
        Case ClientPacketID.eDeleteCharacter
            PacketID_to_string = "eDeleteCharacter"
        Case Else
            PacketID_to_string = "Unknown ClientPacketID (" & CStr(PacketId) & ")"
    End Select
End Function
