Attribute VB_Name = "CharacterPersistence"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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
Option Base 0

Private Const SAVE_CHARACTER_TIME_LIMIT_MS As Long = 50

Private Function db_load_house_key(ByRef User As t_User) As Boolean
    db_load_house_key = False
    With User
        Debug.Assert .Stats.tipoUsuario = tAventurero Or .Stats.tipoUsuario = tHeroe Or .Stats.tipoUsuario = tLeyenda
        Dim RS As ADODB.Recordset
        Set RS = Query("SELECT key_obj FROM house_key WHERE account_id = ?", .AccountID)
        If Not RS Is Nothing Then
            Dim LoopC As Integer
            LoopC = 1
            While Not RS.EOF
                .Keys(LoopC) = RS!key_obj
                LoopC = LoopC + 1
                RS.MoveNext
                db_load_house_key = True
            Wend
        End If
    End With
End Function


Public Function GetCharacterNameByUserId(ByVal UserId As Long) As String
    On Error GoTo GetCharacterNameByUserId_Err

    Dim RS As ADODB.Recordset

    '------------------------------------------------------------------------------
    ' Returns the character name for a given UserId by querying the database.
    '
    ' IMPORTANT:
    ' - This is a DB lookup function (offline-safe).
    ' - The user does NOT need to be logged in.
    ' - The UserId may be stale or invalid (e.g. guild data, spouse links).
    ' - In those cases, this function returns vbNullString and logs the issue.
    '
    ' DO NOT use this function when you already have a logged UserIndex.
    ' In that case, prefer accessing UserList(UserIndex).Name directly.
    '------------------------------------------------------------------------------

    GetCharacterNameByUserId = vbNullString

    If UserId <= 0 Then
        Call LogDatabaseError("GetCharacterNameByUserId: invalid UserId=" & UserId)
        Exit Function
    End If

    Set RS = Query("SELECT name FROM user WHERE id = ? LIMIT 1;", UserId)
    If RS Is Nothing Then
        Call LogDatabaseError("GetCharacterNameByUserId: Query returned Nothing. UserId=" & UserId)
        Exit Function
    End If

    ' IMPORTANT: recordset may be empty if UserId does not exist
    If RS.EOF Or RS.BOF Then
        Call LogDatabaseError("GetCharacterNameByUserId: no DB row for UserId=" & UserId)
        Exit Function
    End If

    If IsNull(RS!Name) Then
        GetCharacterNameByUserId = vbNullString
    Else
        GetCharacterNameByUserId = CStr(RS!Name)
    End If

    Exit Function

GetCharacterNameByUserId_Err:
    Call LogDatabaseError("Error en GetCharacterNameByUserId: UserId=" & UserId & _
                          ". " & Err.Number & " - " & Err.Description & _
                          ". Línea: " & Erl)
End Function

Public Function LoadCharacterBank(ByVal UserIndex As Integer) As Boolean
    On Error GoTo LoadCharacterBank_Err
    With UserList(UserIndex)
        Dim RS      As ADODB.Recordset
        Dim counter As Long
        Set RS = Query("SELECT number, item_id, amount,elemental_tags FROM bank_item WHERE user_id = ?;", .Id)
        counter = 0
        If Not RS Is Nothing Then
            While Not RS.EOF
                With .BancoInvent.Object(RS!Number)
                    .ObjIndex = IIf(RS!item_id <= UBound(ObjData), RS!item_id, 0)
                    If .ObjIndex <> 0 Then
                        If LenB(ObjData(.ObjIndex).name) Then
                            counter = counter + 1
                            .amount = RS!amount
                            .ElementalTags = RS!elemental_tags
                        Else
                            .ObjIndex = 0
                        End If
                    End If
                End With
                RS.MoveNext
            Wend
            .BancoInvent.NroItems = counter
        End If
    End With
    LoadCharacterBank = True
    Exit Function
LoadCharacterBank_Err:
    Call LogDatabaseError("Error en LoadCharacterFromDB LoadCharacterBank: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description & ". Línea: " & Erl)
End Function

Public Function get_num_inv_slots_from_tier(ByVal t As e_TipoUsuario) As Integer
    'By default MAX_USERINVENTORY_SLOTS
    get_num_inv_slots_from_tier = MAX_USERINVENTORY_SLOTS
    Select Case t
        Case tLeyenda
            Const EXTRA_SLOTS_LEYENDA As Integer = 18
            get_num_inv_slots_from_tier = MAX_USERINVENTORY_SLOTS + EXTRA_SLOTS_LEYENDA
        Case tHeroe
            Const EXTRA_SLOTS_HEROE As Integer = 12
            get_num_inv_slots_from_tier = MAX_USERINVENTORY_SLOTS + EXTRA_SLOTS_HEROE
        Case tAventurero
            Const EXTRA_SLOTS_AVENTURERO As Integer = 6
            get_num_inv_slots_from_tier = MAX_USERINVENTORY_SLOTS + EXTRA_SLOTS_AVENTURERO
    End Select
End Function

Public Function LoadCharacterInventory(ByVal UserIndex As Integer) As Boolean
    On Error GoTo LoadCharacterInventory_Err
    With UserList(UserIndex)
        Dim RS                As ADODB.Recordset
        Dim counter           As Long
        Dim SQLQuery          As String
        Dim max_slots_to_load As Integer
        'Load all slots to avoid destroying items when user stops being patreon
        max_slots_to_load = get_num_inv_slots_from_tier(tLeyenda)
        SQLQuery = "SELECT number, item_id, is_equipped, amount, elemental_tags FROM inventory_item WHERE number <= " & max_slots_to_load & " AND user_id = ?;"
        Set RS = Query(SQLQuery, .Id)
        counter = 0
        If Not RS Is Nothing Then
            While Not RS.EOF
                Dim db_inv_slot As Integer
                db_inv_slot = RS!Number
                Debug.Assert db_inv_slot > 0 And db_inv_slot <= UBound(.invent.Object)
                If db_inv_slot > 0 And db_inv_slot <= max_slots_to_load Then
                    'Make sure the slot index is within array bounds and that we don't load slots more slots than required for the current tier
                    With .invent.Object(db_inv_slot)
                        .ObjIndex = IIf(RS!item_id <= UBound(ObjData), RS!item_id, 0)
                        If .ObjIndex <> 0 Then
                            If LenB(ObjData(.ObjIndex).name) Then
                                counter = counter + 1
                                .amount = RS!amount
                                .Equipped = False
                                .ElementalTags = RS!elemental_tags
                                If RS!is_equipped Then
                                    Call EquiparInvItem(UserIndex, RS!Number, True)
                                End If
                            Else
                                .ObjIndex = 0
                            End If
                        End If
                    End With
                End If
                RS.MoveNext
            Wend
            .invent.NroItems = counter
        End If
    End With
    LoadCharacterInventory = True
    Exit Function
LoadCharacterInventory_Err:
    Call LogDatabaseError("Error en LoadCharacterFromDB LoadCharacterInventory: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description & ". Línea: " & Erl)
End Function

Public Function LoadCharacterFromDB(ByVal UserIndex As Integer) As Boolean
    On Error GoTo ErrorHandler
#If LOGIN_STRESS_TEST = 1 Then
    LoadCharacterFromDB = True
    Exit Function
#End If
    
    Dim RS      As ADODB.Recordset
    Dim counter As Long
    LoadCharacterFromDB = False
    With UserList(UserIndex)
        Debug.Assert .id > 0
        ' Load main character data using the user name.
        Set RS = Query(QUERY_LOAD_MAINPJ, .id)
        If RS Is Nothing Then Exit Function
        Debug.Assert .AccountID > -1
        If CLng(RS!account_id) <> .AccountID Then
            Call CloseSocket(UserIndex)
            Exit Function
        End If

        .name = CStr(RS!name)

        ' Set up the Patreon tier early (needed by subsequent initialization).
        .Stats.tipoUsuario = GetPatronTierFromAccountID(.AccountID)
        .flags.is_donor = GetUserContributionStatusFromAccountID(.AccountID)
        If .flags.is_donor > 0 Then
            .Stats.tipoUsuario = e_TipoUsuario.tLeyenda
        End If
        ' Check for ban status.
        If RS!is_banned Then
            Dim BanNick As String, BaneoMotivo As String
            BanNick = RS!banned_by
            BaneoMotivo = RS!ban_reason
            If LenB(BanNick) = 0 Then BanNick = "*Error en la base de datos*"
            If LenB(BaneoMotivo) = 0 Then BaneoMotivo = "*No se registra el motivo del baneo.*"
            Call WriteShowMessageBox(UserIndex, 1755, BaneoMotivo & "¬" & BanNick) ' Msg1755=Se te ha prohibido la entrada al juego debido a ¬1. Esta decisión fue tomada por ¬2.
            Call CloseSocket(UserIndex)
            Exit Function
        End If
        ' Check if the character is locked/in a sale state.
        If RS!is_locked_in_mao Then
            Call WriteShowMessageBox(UserIndex, 1756, vbNullString) 'Msg1756=El personaje que estás intentando loguear se encuentra en venta, para desbloquearlo deberás hacerlo desde la página web.
            Call CloseSocket(UserIndex)
            Exit Function
        End If
        .Stats.shield = 0
        .InUse = True
        ' Set up core user fields.
        Call SetupUserBasicInfo(UserList(UserIndex), RS)
        Call SetupUserFlags(UserList(UserIndex), RS)
        Call SetupUserFactionInfo(UserList(UserIndex), RS)
        ' Refactored sections: load spells, pets, bank inventory, skills, quests, and completed quests.
        Call SetupUserSpells(UserList(UserIndex))
        Call SetupUserPets(UserList(UserIndex))
        Call SetupUserBankInventory(UserList(UserIndex))
        Call SetupUserSkills(UserList(UserIndex))
        Call SetupUserQuests(UserList(UserIndex))
        Call SetupUserQuestsDone(UserList(UserIndex))
        ' Load additional inventories.
        If Not LoadCharacterInventory(UserIndex) Then Exit Function
        If Not LoadCharacterBank(UserIndex) Then Exit Function
        Call LoadSkinsInventory(UserIndex)
        Call RegisterUserName(.Id, .name)
        Call Execute("update account set last_ip = ? where id = ?", .ConnectionDetails.IP, .AccountID)
        .Stats.Creditos = 0
        ' If the user is a patron-type, load the house key.
        If .Stats.tipoUsuario = tAventurero Or .Stats.tipoUsuario = tHeroe Or .Stats.tipoUsuario = tLeyenda Then
            Call db_load_house_key(UserList(UserIndex))
        End If
    End With
    Call InitUserPersistSnapshot(UserIndex)
    LoadCharacterFromDB = True
    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error en LoadCharacterFromDB: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description & ". Línea: " & Erl)
End Function

Private Sub SetupUserBasicInfo(ByRef User As t_User, ByRef RS As ADODB.Recordset)
    With User
        .Id = RS!Id
        .name = RS!name
        If IsNull(RS!alias) Then
            .Alias = vbNullString
        Else
            .Alias = RS!alias
        End If
        .Stats.ELV = RS!level
        .Stats.Exp = RS!Exp
        .genero = RS!genre_id
        .raza = RS!race_id
        .clase = RS!class_id
        .Hogar = RS!home_id
        .Desc = RS!Description
        .Stats.GLD = RS!gold
        .Stats.Banco = RS!bank_gold
        .Stats.SkillPts = RS!free_skillpoints
        .pos.Map = RS!pos_map
        .pos.x = RS!pos_X
        .pos.y = RS!pos_Y
        .MENSAJEINFORMACION = RS!message_info
        .OrigChar.body = RS!body_id
        .OrigChar.head = RS!head_id
        .OrigChar.originalhead = .OrigChar.head
        .OrigChar.WeaponAnim = RS!weapon_id
        .OrigChar.CascoAnim = RS!helmet_id
        .OrigChar.ShieldAnim = RS!shield_id
        .OrigChar.BackpackAnim = RS!backpack_id
        .OrigChar.Heading = RS!Heading
        .Stats.MaxHp = RS!max_hp
        .Stats.MinHp = RS!min_hp
        .Stats.MinMAN = RS!min_man
        .Stats.MinSta = RS!min_sta
        .Stats.MinHam = RS!min_ham
        .Stats.MaxHam = 100
        .Stats.MinAGU = RS!min_sed
        .Stats.MaxAGU = 100
        .Stats.NPCsMuertos = RS!killed_npcs
        .Stats.UsuariosMatados = RS!killed_users
        .Stats.PuntosPesca = RS!puntos_pesca
        .Stats.ELO = RS!ELO
        .Stats.JineteLevel = RS!jinete_level
        .Counters.Pena = RS!counter_pena
        .ChatGlobal = RS!chat_global
        .ChatCombate = RS!chat_combate
        .Stats.Advertencias = RS!warnings
        .GuildIndex = SanitizeNullValue(RS!Guild_Index, 0)
        .LastGuildRejection = SanitizeNullValue(RS!guild_rejected_because, vbNullString)
    End With
End Sub

Private Sub SetupUserFlags(ByRef User As t_User, ByRef RS As ADODB.Recordset)
    With User.flags
        .Desnudo = RS!is_naked
        .Envenenado = RS!is_poisoned
        .Incinerado = RS!is_incinerated
        .Escondido = False
        .Ban = RS!is_banned
        .Muerto = RS!is_dead
        .Navegando = RS!is_sailing
        .Paralizado = RS!is_paralyzed
        .VecesQueMoriste = RS!deaths
        .Montado = RS!is_mounted
        .SpouseId = RS!spouse
        .Casado = IIf(.SpouseId > 0, 1, 0)
        .Silenciado = RS!is_silenced
        .MinutosRestantes = RS!silence_minutes_left
        .SegundosPasados = RS!silence_elapsed_seconds
        .MascotasGuardadas = RS!pets_saved
        .ReturnPos.Map = RS!return_map
        .ReturnPos.x = RS!return_x
        .ReturnPos.y = RS!return_y
    End With
End Sub

Private Sub SetupUserFactionInfo(ByRef User As t_User, ByRef RS As ADODB.Recordset)
    With User.Faccion
        .ciudadanosMatados = RS!ciudadanos_matados
        .CriminalesMatados = RS!criminales_matados
        .RecibioArmaduraReal = RS!recibio_armadura_real
        .RecibioArmaduraCaos = RS!recibio_armadura_caos
        .RecompensasReal = RS!recompensas_real
        .RecompensasCaos = RS!recompensas_caos
        .Reenlistadas = RS!Reenlistadas
        .NivelIngreso = SanitizeNullValue(RS!nivel_ingreso, 0)
        .MatadosIngreso = SanitizeNullValue(RS!matados_ingreso, 0)
        .Status = RS!Status
        .FactionScore = RS!faction_score
    End With
End Sub

Private Sub SetupUserSpells(ByRef User As t_User)
    Dim RS As ADODB.Recordset
    Set RS = Query("SELECT number, spell_id FROM spell WHERE user_id = ?;", User.Id)
    If Not RS Is Nothing Then
        While Not RS.EOF
            User.Stats.UserHechizos(RS!Number) = RS!spell_id
            RS.MoveNext
        Wend
    End If
End Sub

Private Sub SetupUserPets(ByRef User As t_User)
    Dim RS As ADODB.Recordset
    Set RS = Query("SELECT number, pet_id FROM pet WHERE user_id = ?;", User.Id)
    If Not RS Is Nothing Then
        While Not RS.EOF
            User.MascotasType(RS!Number) = RS!pet_id
            If val(RS!pet_id) <> 0 Then
                User.NroMascotas = User.NroMascotas + 1
            End If
            RS.MoveNext
        Wend
    End If
End Sub

Private Sub SetupUserBankInventory(ByRef User As t_User)
    Dim RS      As ADODB.Recordset
    Dim counter As Long
    counter = 0
    Set RS = Query("SELECT number, item_id, amount, elemental_tags FROM bank_item WHERE user_id = ?;", User.Id)
    If Not RS Is Nothing Then
        While Not RS.EOF
            With User.BancoInvent.Object(RS!Number)
                .ObjIndex = RS!item_id
                If .ObjIndex <> 0 Then
                    If LenB(ObjData(.ObjIndex).name) > 0 Then
                        counter = counter + 1
                        .amount = RS!amount
                        .ElementalTags = RS!elemental_tags
                    Else
                        .ObjIndex = 0
                    End If
                End If
            End With
            RS.MoveNext
        Wend
        User.BancoInvent.NroItems = counter
    End If
End Sub

Private Sub SetupUserSkills(ByRef User As t_User)
    Dim RS As ADODB.Recordset
    Dim LoopC As Long
    Set RS = Query("SELECT number, value FROM skillpoint WHERE user_id = ?;", User.Id)
    If Not RS Is Nothing Then
        While Not RS.EOF
            User.Stats.UserSkills(RS!Number) = RS!value
            User.Stats.SkillDirty(RS!Number) = False
            RS.MoveNext
        Wend
    End If

    For LoopC = 1 To NUMSKILLS
        User.Stats.SkillDirty(LoopC) = False
    Next LoopC
End Sub

Private Sub SetupUserQuests(ByRef User As t_User)
    Dim RS    As ADODB.Recordset
    Dim LoopC As Byte
    Dim SlotC As Byte
    For SlotC = 1 To MAXUSERQUESTS
        User.QuestStats.Quests(SlotC).Dirty = False ' Loaded state starts clean.
    Next SlotC
    Set RS = Query("SELECT number, quest_id, npcs, npcstarget FROM quest WHERE user_id = ?;", User.Id)
    If Not RS Is Nothing Then
        While Not RS.EOF
            If Not IsNull(RS!Number) Then
                User.QuestStats.Quests(RS!Number).QuestIndex = RS!quest_id
                User.QuestStats.Quests(RS!Number).Dirty = False ' Loaded from DB.
                If User.QuestStats.Quests(RS!Number).QuestIndex > 0 Then
                    If QuestList(User.QuestStats.Quests(RS!Number).QuestIndex).RequiredNPCs Then
                        Dim NPCs() As String
                        NPCs = Split(RS!NPCs, "-")
                        ReDim User.QuestStats.Quests(RS!Number).NPCsKilled(1 To QuestList(User.QuestStats.Quests(RS!Number).QuestIndex).RequiredNPCs)
                        For LoopC = 1 To QuestList(User.QuestStats.Quests(RS!Number).QuestIndex).RequiredNPCs
                            User.QuestStats.Quests(RS!Number).NPCsKilled(LoopC) = val(NPCs(LoopC - 1))
                        Next LoopC
                    End If
                    If QuestList(User.QuestStats.Quests(RS!Number).QuestIndex).RequiredTargetNPCs Then
                        Dim NPCsTarget() As String
                        NPCsTarget = Split(RS!NPCsTarget, "-")
                        ReDim User.QuestStats.Quests(RS!Number).NPCsTarget(1 To QuestList(User.QuestStats.Quests(RS!Number).QuestIndex).RequiredTargetNPCs)
                        For LoopC = 1 To QuestList(User.QuestStats.Quests(RS!Number).QuestIndex).RequiredTargetNPCs
                            User.QuestStats.Quests(RS!Number).NPCsTarget(LoopC) = val(NPCsTarget(LoopC - 1))
                        Next LoopC
                    End If
                End If
            End If
            RS.MoveNext
        Wend
    End If
End Sub

Private Sub SetupUserQuestsDone(ByRef User As t_User)
    Dim RS    As ADODB.Recordset
    Dim LoopC As Byte
    Set RS = Query("SELECT quest_id FROM quest_done WHERE user_id = ?;", User.Id)
    If Not RS Is Nothing Then
        User.QuestStats.NumQuestsDone = RS.RecordCount
        If User.QuestStats.NumQuestsDone > 0 Then
            ReDim User.QuestStats.QuestsDone(1 To User.QuestStats.NumQuestsDone)
            LoopC = 1
            While Not RS.EOF
                User.QuestStats.QuestsDone(LoopC) = RS!quest_id
                LoopC = LoopC + 1
                RS.MoveNext
            Wend
        End If
    End If
End Sub

''' <summary>
''' Returns the maximum number of characters allowed for a given user tier.
''' </summary>
''' <param name="tier">The user tier.</param>
''' <returns>The maximum number of characters allowed.</returns>
Public Function MaxCharacterForTier(ByVal tier As e_TipoUsuario)
    #If DEBUGGING Then
        MaxCharacterForTier = 10
    #Else
        Select Case tier
            Case e_TipoUsuario.tAventurero
                MaxCharacterForTier = 3
            Case e_TipoUsuario.tHeroe
                MaxCharacterForTier = 5
            Case e_TipoUsuario.tLeyenda
                MaxCharacterForTier = 10
            Case Else
                MaxCharacterForTier = 1
        End Select
    #End If
End Function

Public Function GetPatronTierFromAccountID(ByVal account_id) As e_TipoUsuario
    On Error GoTo ErrorHandler_GetPatronTierFromAccountID
    GetPatronTierFromAccountID = e_TipoUsuario.tNormal
    Dim RS As ADODB.Recordset
    Set RS = Query("Select is_active_patron from account where id = ?", account_id)
    If Not RS Is Nothing Then
        Dim tipo_usuario_db As Long
        tipo_usuario_db = RS!is_active_patron
        Select Case tipo_usuario_db
            Case patron_tier_aventurero
                GetPatronTierFromAccountID = e_TipoUsuario.tAventurero
            Case patron_tier_heroe
                GetPatronTierFromAccountID = e_TipoUsuario.tHeroe
            Case patron_tier_leyenda
                GetPatronTierFromAccountID = e_TipoUsuario.tLeyenda
            Case Else
                GetPatronTierFromAccountID = e_TipoUsuario.tNormal
        End Select
    End If
    Exit Function
ErrorHandler_GetPatronTierFromAccountID:
    Call LogDatabaseError("Error en GetPatronTierFromAccountID: " & account_id & ". " & Err.Number & " - " & Err.Description & ". Línea: " & Erl)
End Function

Public Function GetUserContributionStatusFromAccountID(ByVal account_id) As Byte
    On Error GoTo ErrorHandler_GetUserContributionStatusFromAccountID
    Dim RS As ADODB.Recordset
    Set RS = Query("Select is_donor from account where id = ?", account_id)
    If Not RS Is Nothing Then
        GetUserContributionStatusFromAccountID = RS!is_donor
    End If
    Exit Function
ErrorHandler_GetUserContributionStatusFromAccountID:
    Call LogDatabaseError("Error en GetUserContributionStatusFromAccountID: " & account_id & ". " & Err.Number & " - " & Err.Description & ". Línea: " & Erl)
End Function

Public Sub LoadPatronCreditsFromDB(ByVal UserIndex As Integer)
    Dim RS As ADODB.Recordset
    With UserList(UserIndex)
        Set RS = Query("Select offline_patron_credits from account where id = ?;", .AccountID)
        If Not RS Is Nothing Then
            .Stats.Creditos = RS!offline_patron_credits
        Else
            .Stats.Creditos = 0
        End If
    End With
End Sub

Public Sub SaveCharacterDB(ByVal UserIndex As Integer)
    On Error GoTo ErrorHandler
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    Dim QueryBreakdown As String
    Call Builder.Clear
    With UserList(UserIndex)
        Debug.Assert .flags.UserLogged = True
        If Not .flags.UserLogged Then
            Call LogDatabaseError("Error trying to save an user not logged in SaveCharacterDB")
            Exit Sub
        End If
        Call SaveCharacterMainDB(UserList(UserIndex), QueryBreakdown)
        Call SaveCharacterSpellsDB(UserList(UserIndex), QueryBreakdown)
        Call SaveCharacterInventoryDB(UserList(UserIndex), QueryBreakdown)
        If HasBankChanged(UserIndex) Then
            Call SaveCharacterBankInventoryDB(UserList(UserIndex), QueryBreakdown)
        Else
            If LenB(QueryBreakdown) <> 0 Then QueryBreakdown = QueryBreakdown & "; "
            QueryBreakdown = QueryBreakdown & "save bank inventory: skipped"
        End If
        Call SaveCharacterSkillsDB(UserList(UserIndex), QueryBreakdown)
        Call SaveCharacterPetsDB(UserList(UserIndex), QueryBreakdown)
        ' ************************** User quests *********************************
        Call SaveCharacterQuestsDB(UserList(UserIndex), QueryBreakdown, Builder)
        Call SaveCharacterQuestsDoneDB(UserList(UserIndex), QueryBreakdown, Builder)
        Call SaveCharacterInventorySkinsDB(UserIndex, QueryBreakdown)
        Call InitUserPersistSnapshot(UserIndex)
        Call LogSaveCharacterDuration(PerformanceTimer, QueryBreakdown, .name, .Id)
    End With
    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error en SaveUserDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description)
End Sub

Private Sub SaveCharacterMainDB(ByRef U As t_User, ByRef QueryBreakdown As String)
    Dim QueryTimer As Long
    Dim Params() As Variant
    ReDim Params(61)
    Dim i As Integer
    Params(post_increment(i)) = U.Stats.ELV
    Params(post_increment(i)) = U.Stats.Exp
    Params(post_increment(i)) = U.Hogar
    Params(post_increment(i)) = U.Desc
    Params(post_increment(i)) = U.Stats.GLD
    Params(post_increment(i)) = U.Stats.Banco
    Params(post_increment(i)) = U.Stats.SkillPts
    Params(post_increment(i)) = U.flags.MascotasGuardadas
    Params(post_increment(i)) = U.pos.Map
    Params(post_increment(i)) = U.pos.x
    Params(post_increment(i)) = U.pos.y
    Params(post_increment(i)) = U.MENSAJEINFORMACION
    Params(post_increment(i)) = U.Char.body
    Params(post_increment(i)) = U.OrigChar.originalhead
    Params(post_increment(i)) = U.Char.WeaponAnim
    Params(post_increment(i)) = U.Char.CascoAnim
    Params(post_increment(i)) = U.Char.ShieldAnim
    Params(post_increment(i)) = U.Char.Heading
    Params(post_increment(i)) = U.Stats.MaxHp
    Params(post_increment(i)) = U.Stats.MinHp
    Params(post_increment(i)) = U.Stats.MinMAN
    Params(post_increment(i)) = U.Stats.MinSta
    Params(post_increment(i)) = U.Stats.MinHam
    Params(post_increment(i)) = U.Stats.MinAGU
    Params(post_increment(i)) = U.Stats.NPCsMuertos
    Params(post_increment(i)) = U.Stats.UsuariosMatados
    Params(post_increment(i)) = U.Stats.PuntosPesca
    Params(post_increment(i)) = U.Stats.ELO
    Params(post_increment(i)) = U.flags.Desnudo
    Params(post_increment(i)) = U.flags.Envenenado
    Params(post_increment(i)) = U.flags.Incinerado
    Params(post_increment(i)) = U.flags.Muerto
    Params(post_increment(i)) = U.flags.Navegando
    Params(post_increment(i)) = U.flags.Paralizado
    Params(post_increment(i)) = U.flags.Montado
    Params(post_increment(i)) = U.flags.Silenciado
    Params(post_increment(i)) = U.flags.MinutosRestantes
    Params(post_increment(i)) = U.flags.SegundosPasados
    Params(post_increment(i)) = U.flags.SpouseId
    Params(post_increment(i)) = U.Counters.Pena
    Params(post_increment(i)) = U.flags.VecesQueMoriste
    Params(post_increment(i)) = U.Faccion.ciudadanosMatados
    Params(post_increment(i)) = U.Faccion.CriminalesMatados
    Params(post_increment(i)) = U.Faccion.RecibioArmaduraReal
    Params(post_increment(i)) = U.Faccion.RecibioArmaduraCaos
    Params(post_increment(i)) = U.Faccion.RecompensasReal
    Params(post_increment(i)) = U.Faccion.FactionScore
    Params(post_increment(i)) = U.Faccion.RecompensasCaos
    Params(post_increment(i)) = U.Faccion.Reenlistadas
    Params(post_increment(i)) = U.Faccion.NivelIngreso
    Params(post_increment(i)) = U.Faccion.MatadosIngreso
    Params(post_increment(i)) = U.Faccion.Status
    Params(post_increment(i)) = U.GuildIndex
    Params(post_increment(i)) = U.ChatCombate
    Params(post_increment(i)) = U.ChatGlobal
    Params(post_increment(i)) = U.Stats.Advertencias
    Params(post_increment(i)) = U.flags.ReturnPos.Map
    Params(post_increment(i)) = U.flags.ReturnPos.x
    Params(post_increment(i)) = U.flags.ReturnPos.y
    Params(post_increment(i)) = U.Stats.JineteLevel
    Params(post_increment(i)) = U.Char.BackpackAnim
    ' WHERE block
    Params(post_increment(i)) = U.Id
    Debug.Assert i = UBound(Params) + 1
    QueryTimer = GetTickCountRaw()
    Call Execute(QUERY_UPDATE_MAINPJ, Params)
    Call AppendQueryDuration(QueryBreakdown, "update main", QueryTimer)
End Sub

Private Sub SaveCharacterSpellsDB(ByRef U As t_User, ByRef QueryBreakdown As String)
    Dim QueryTimer As Long
    Dim Params() As Variant
    Dim LoopC As Long
    Dim ParamC As Long
    ReDim Params(MAXUSERHECHIZOS * 3 - 1)
    ParamC = 0
    For LoopC = 1 To MAXUSERHECHIZOS
        Params(ParamC) = U.Id
        Params(ParamC + 1) = LoopC
        Params(ParamC + 2) = U.Stats.UserHechizos(LoopC)
        ParamC = ParamC + 3
    Next LoopC
    QueryTimer = GetTickCountRaw()
    Call Execute(QUERY_UPSERT_SPELLS, Params)
    Call AppendQueryDuration(QueryBreakdown, "upsert spells", QueryTimer)
    For LoopC = 1 To MAXUSERHECHIZOS
        U.Persist.LastSpells(LoopC) = U.Stats.UserHechizos(LoopC)
    Next LoopC
End Sub

Private Sub SaveCharacterInventoryDB(ByRef U As t_User, ByRef QueryBreakdown As String)
    Dim QueryTimer As Long
    Dim Params() As Variant
    Dim LoopC As Long
    Dim ParamC As Long
    ReDim Params(MAX_INVENTORY_SLOTS * 6 - 1)
    ParamC = 0
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Params(ParamC) = U.Id
        Params(ParamC + 1) = LoopC
        Params(ParamC + 2) = U.invent.Object(LoopC).ObjIndex
        Params(ParamC + 3) = U.invent.Object(LoopC).amount
        Params(ParamC + 4) = U.invent.Object(LoopC).Equipped
        Params(ParamC + 5) = U.invent.Object(LoopC).ElementalTags
        ParamC = ParamC + 6
    Next LoopC
    QueryTimer = GetTickCountRaw()
    Call Execute(QUERY_UPSERT_INVENTORY, Params)
    Call AppendQueryDuration(QueryBreakdown, "upsert inventory", QueryTimer)
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        U.Persist.LastInventory(LoopC) = U.invent.Object(LoopC)
    Next LoopC
End Sub

Private Sub SaveCharacterBankInventoryDB(ByRef U As t_User, ByRef QueryBreakdown As String)
    Dim QueryTimer As Long
    Dim Params() As Variant
    Dim LoopC As Long
    Dim ParamC As Long
    ReDim Params(MAX_BANCOINVENTORY_SLOTS * 5 - 1)
    ParamC = 0
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
        Params(ParamC) = U.Id
        Params(ParamC + 1) = LoopC
        Params(ParamC + 2) = U.BancoInvent.Object(LoopC).ObjIndex
        Params(ParamC + 3) = U.BancoInvent.Object(LoopC).amount
        Params(ParamC + 4) = U.BancoInvent.Object(LoopC).ElementalTags
        ParamC = ParamC + 5
    Next LoopC
    QueryTimer = GetTickCountRaw()
    Call ExecutePreparedBankSave(Params)
    Call AppendQueryDuration(QueryBreakdown, "save bank inventory", QueryTimer)
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
        U.Persist.LastBank(LoopC) = U.BancoInvent.Object(LoopC)
    Next LoopC
End Sub

Private Sub SaveCharacterSkillsDB(ByRef U As t_User, ByRef QueryBreakdown As String)
    Dim QueryTimer As Long
    Dim Params() As Variant
    Dim LoopC As Long
    Dim ParamC As Long
    Dim DirtyCount As Long

    For LoopC = 1 To NUMSKILLS
        If U.Stats.SkillDirty(LoopC) Then
            DirtyCount = DirtyCount + 1
        End If
    Next LoopC

    If DirtyCount = 0 Then
        If LenB(QueryBreakdown) <> 0 Then QueryBreakdown = QueryBreakdown & "; "
        QueryBreakdown = QueryBreakdown & "upsert skills: skipped"
        Exit Sub
    End If

    Call Builder.Clear
    Builder.Append "REPLACE INTO skillpoint (user_id, number, value) VALUES "
    For LoopC = 1 To DirtyCount
        Builder.Append "(?, ?, ?)"
        If LoopC < DirtyCount Then
            Builder.Append ", "
        End If
    Next LoopC

    ReDim Params(DirtyCount * 3 - 1)
    ParamC = 0
    For LoopC = 1 To NUMSKILLS
        If U.Stats.SkillDirty(LoopC) Then
            Params(ParamC) = U.Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = U.Stats.UserSkills(LoopC)
            ParamC = ParamC + 3
        End If
    Next LoopC

    QueryTimer = GetTickCountRaw()
    Call Execute(Builder.ToString(), Params)
    Call AppendQueryDuration(QueryBreakdown, "upsert skills", QueryTimer)
    Call Builder.Clear

    For LoopC = 1 To NUMSKILLS
        If U.Stats.SkillDirty(LoopC) Then
            U.Stats.SkillDirty(LoopC) = False
            U.Persist.LastSkills(LoopC) = U.Stats.UserSkills(LoopC)
        End If
    Next LoopC
End Sub

Private Sub SaveCharacterPetsDB(ByRef U As t_User, ByRef QueryBreakdown As String)
    Dim QueryTimer As Long
    Dim Params() As Variant
    Dim LoopC As Long
    Dim ParamC As Long
    Dim petType As Integer
    ReDim Params(MAXMASCOTAS * 3 - 1)
    ParamC = 0
    For LoopC = 1 To MAXMASCOTAS
        Params(ParamC) = U.Id
        Params(ParamC + 1) = LoopC
        If IsValidNpcRef(U.MascotasIndex(LoopC)) Then
            If NpcList(U.MascotasIndex(LoopC).ArrayIndex).Contadores.TiempoExistencia = 0 Then
                petType = U.MascotasType(LoopC)
            Else
                petType = 0
            End If
        Else
            petType = U.MascotasType(LoopC)
        End If
        Params(ParamC + 2) = petType
        ParamC = ParamC + 3
    Next LoopC
    QueryTimer = GetTickCountRaw()
    Call Execute(QUERY_UPSERT_PETS, Params)
    Call AppendQueryDuration(QueryBreakdown, "upsert pets", QueryTimer)
    For LoopC = 1 To MAXMASCOTAS
        If IsValidNpcRef(U.MascotasIndex(LoopC)) Then
            If NpcList(U.MascotasIndex(LoopC).ArrayIndex).Contadores.TiempoExistencia = 0 Then
                petType = U.MascotasType(LoopC)
            Else
                petType = 0
            End If
        Else
            petType = U.MascotasType(LoopC)
        End If
        U.Persist.LastPetType(LoopC) = petType
    Next LoopC
End Sub

Private Sub SaveCharacterQuestsDoneDB(ByRef U As t_User, ByRef QueryBreakdown As String, ByRef SqlBuilder As cStringBuilder)
    Dim QueryTimer As Long
    Dim Params() As Variant
    Dim LoopC As Long
    Dim ParamC As Long

    If Not HaveQuestsDoneChanged(U) Then
        If Len(QueryBreakdown) > 0 Then QueryBreakdown = QueryBreakdown & "; "
        QueryBreakdown = QueryBreakdown & "quests done skipped"
        Exit Sub
    End If
    If U.QuestStats.NumQuestsDone > 0 Then
        SqlBuilder.Append "REPLACE INTO quest_done (user_id, quest_id) VALUES "
        For LoopC = 1 To U.QuestStats.NumQuestsDone
            SqlBuilder.Append "(?, ?)"
            If LoopC < U.QuestStats.NumQuestsDone Then
                SqlBuilder.Append ", "
            End If
        Next LoopC
        ReDim Params(U.QuestStats.NumQuestsDone * 2 - 1)
        ParamC = 0
        For LoopC = 1 To U.QuestStats.NumQuestsDone
            Params(ParamC) = U.Id
            Params(ParamC + 1) = U.QuestStats.QuestsDone(LoopC)
            ParamC = ParamC + 2
        Next LoopC
        QueryTimer = GetTickCountRaw()
        Call Execute(SqlBuilder.ToString(), Params)
        Call AppendQueryDuration(QueryBreakdown, "replace quests done", QueryTimer)
        Call SqlBuilder.Clear
    Else
        QueryTimer = GetTickCountRaw()
        Call Execute("DELETE FROM quest_done WHERE user_id = ?;", U.Id)
        Call AppendQueryDuration(QueryBreakdown, "delete quests done", QueryTimer)
    End If

    If U.QuestStats.NumQuestsDone > 0 Then
        ReDim U.Persist.LastQuestsDone(1 To U.QuestStats.NumQuestsDone)
        For LoopC = 1 To U.QuestStats.NumQuestsDone
            U.Persist.LastQuestsDone(LoopC) = GetIntegerArrayValue(U.QuestStats.QuestsDone, LoopC)
        Next LoopC
    Else
        ReDim U.Persist.LastQuestsDone(0)
    End If
End Sub

Private Sub SaveCharacterInventorySkinsDB(ByVal UserIndex As Integer, ByRef QueryBreakdown As String)
    Dim QueryTimer As Long
    QueryTimer = GetTickCountRaw()
    Call SaveInventorySkins(UserIndex)
    Call AppendQueryDuration(QueryBreakdown, "save inventory skins", QueryTimer)
End Sub

Private Sub SaveCharacterQuestsDB(ByRef User As t_User, ByRef QueryBreakdown As String, ByRef SqlBuilder As cStringBuilder)
    Dim QueryTimer As Long
    Dim LoopC As Long
    Dim LoopK As Long
    Dim Tmp As Integer
    Dim DirtyQuestSlotsSaved As Long
    Dim DirtyQuestSlotsDeleted As Long
    Dim DirtyQuestSlotsTotal As Long
    Dim DirtyQuestSaveSlots() As Integer
    Dim DirtyQuestDeleteSlots() As Integer
    Dim QuestSlotToSave As Integer

    ' Split dirty quest slots into two groups:
    '   1) Active slots (QuestIndex > 0) that must be upserted.
    '   2) Cleared slots (QuestIndex = 0) that must be deleted.
    For LoopC = 1 To MAXUSERQUESTS
        If User.QuestStats.Quests(LoopC).Dirty Then
            If User.QuestStats.Quests(LoopC).QuestIndex > 0 Then
                DirtyQuestSlotsSaved = DirtyQuestSlotsSaved + 1
                If DirtyQuestSlotsSaved = 1 Then
                    ReDim DirtyQuestSaveSlots(1 To 1)
                Else
                    ReDim Preserve DirtyQuestSaveSlots(1 To DirtyQuestSlotsSaved)
                End If
                DirtyQuestSaveSlots(DirtyQuestSlotsSaved) = LoopC
            Else
                DirtyQuestSlotsDeleted = DirtyQuestSlotsDeleted + 1
                If DirtyQuestSlotsDeleted = 1 Then
                    ReDim DirtyQuestDeleteSlots(1 To 1)
                Else
                    ReDim Preserve DirtyQuestDeleteSlots(1 To DirtyQuestSlotsDeleted)
                End If
                DirtyQuestDeleteSlots(DirtyQuestSlotsDeleted) = LoopC
            End If
        End If
    Next LoopC

    DirtyQuestSlotsTotal = DirtyQuestSlotsSaved + DirtyQuestSlotsDeleted

    ' Persist only dirty active quest slots, keeping existing dash-separated
    ' serialization for NPC kill and target progress columns.
    If DirtyQuestSlotsSaved > 0 Then
        SqlBuilder.Append "REPLACE INTO quest (user_id, number, quest_id, npcs, npcstarget) VALUES "
        For LoopC = 1 To DirtyQuestSlotsSaved
            QuestSlotToSave = DirtyQuestSaveSlots(LoopC)

            SqlBuilder.Append "("
            SqlBuilder.Append User.Id & ", "
            SqlBuilder.Append QuestSlotToSave & ", "
            SqlBuilder.Append User.QuestStats.Quests(QuestSlotToSave).QuestIndex & ", '"
            Tmp = QuestList(User.QuestStats.Quests(QuestSlotToSave).QuestIndex).RequiredNPCs
            If Tmp Then
                For LoopK = 1 To Tmp
                    SqlBuilder.Append CStr(User.QuestStats.Quests(QuestSlotToSave).NPCsKilled(LoopK))
                    If LoopK < Tmp Then
                        SqlBuilder.Append "-"
                    End If
                Next LoopK
            End If
            SqlBuilder.Append "', '"
            Tmp = QuestList(User.QuestStats.Quests(QuestSlotToSave).QuestIndex).RequiredTargetNPCs
            For LoopK = 1 To Tmp
                SqlBuilder.Append CStr(User.QuestStats.Quests(QuestSlotToSave).NPCsTarget(LoopK))
                If LoopK < Tmp Then
                    SqlBuilder.Append "-"
                End If
            Next LoopK
            SqlBuilder.Append "')"
            If LoopC < DirtyQuestSlotsSaved Then
                SqlBuilder.Append ", "
            End If
        Next LoopC

        QueryTimer = GetTickCountRaw()
        Call Execute(SqlBuilder.ToString())
        Call AppendQueryDuration(QueryBreakdown, "replace dirty quests", QueryTimer)
        Call SqlBuilder.Clear
    End If

    ' Remove dirty slots that were reset/abandoned and no longer have a quest.
    If DirtyQuestSlotsDeleted > 0 Then
        SqlBuilder.Append "DELETE FROM quest WHERE user_id = " & User.Id & " AND number IN ("
        For LoopC = 1 To DirtyQuestSlotsDeleted
            SqlBuilder.Append CStr(DirtyQuestDeleteSlots(LoopC))
            If LoopC < DirtyQuestSlotsDeleted Then
                SqlBuilder.Append ", "
            End If
        Next LoopC
        SqlBuilder.Append ")"

        QueryTimer = GetTickCountRaw()
        Call Execute(SqlBuilder.ToString())
        Call AppendQueryDuration(QueryBreakdown, "delete dirty quests", QueryTimer)
        Call SqlBuilder.Clear
    End If

    ' If no dirty slots were found, explicitly log that quest persistence was skipped.
    If DirtyQuestSlotsTotal = 0 Then
        If Len(QueryBreakdown) > 0 Then
            QueryBreakdown = QueryBreakdown & "; "
        End If
        QueryBreakdown = QueryBreakdown & "quests skipped"
    End If

    ' Emit counters so performance logs can show real quest write pressure.
    If Len(QueryBreakdown) > 0 Then
        QueryBreakdown = QueryBreakdown & "; "
    End If
    QueryBreakdown = QueryBreakdown & "dirty_quest_slots_saved = " & DirtyQuestSlotsSaved & "; dirty_quest_slots_deleted = " & DirtyQuestSlotsDeleted

    ' Clear dirty flags only after DB operations completed successfully.
    For LoopC = 1 To DirtyQuestSlotsSaved
        User.QuestStats.Quests(DirtyQuestSaveSlots(LoopC)).Dirty = False ' Cleared only after successful DB write.
    Next LoopC
    For LoopC = 1 To DirtyQuestSlotsDeleted
        User.QuestStats.Quests(DirtyQuestDeleteSlots(LoopC)).Dirty = False ' Cleared only after successful DB delete.
    Next LoopC

    For LoopC = 1 To MAXUSERQUESTS
        User.Persist.LastQuests(LoopC).QuestIndex = User.QuestStats.Quests(LoopC).QuestIndex

        If User.QuestStats.Quests(LoopC).QuestIndex > 0 Then
            Tmp = QuestList(User.QuestStats.Quests(LoopC).QuestIndex).RequiredNPCs
        Else
            Tmp = 0
        End If

        If Tmp > 0 Then
            ReDim User.Persist.LastQuests(LoopC).NPCsKilled(1 To Tmp)
            For LoopK = 1 To Tmp
                User.Persist.LastQuests(LoopC).NPCsKilled(LoopK) = GetIntegerArrayValue(User.QuestStats.Quests(LoopC).NPCsKilled, LoopK)
            Next LoopK
        Else
            ReDim User.Persist.LastQuests(LoopC).NPCsKilled(0)
        End If

        If User.QuestStats.Quests(LoopC).QuestIndex > 0 Then
            Tmp = QuestList(User.QuestStats.Quests(LoopC).QuestIndex).RequiredTargetNPCs
        Else
            Tmp = 0
        End If

        If Tmp > 0 Then
            ReDim User.Persist.LastQuests(LoopC).NPCsTarget(1 To Tmp)
            For LoopK = 1 To Tmp
                User.Persist.LastQuests(LoopC).NPCsTarget(LoopK) = GetIntegerArrayValue(User.QuestStats.Quests(LoopC).NPCsTarget, LoopK)
            Next LoopK
        Else
            ReDim User.Persist.LastQuests(LoopC).NPCsTarget(0)
        End If
    Next LoopC
End Sub


Private Sub AppendQueryDuration(ByRef QueryBreakdown As String, ByVal QueryName As String, ByVal QueryStartTime As Long)
    Dim elapsed As Double
    elapsed = TicksElapsed(QueryStartTime, GetTickCountRaw())
    If Len(QueryBreakdown) > 0 Then
        QueryBreakdown = QueryBreakdown & "; "
    End If
    QueryBreakdown = QueryBreakdown & QueryName & ": " & CLng(elapsed) & "ms"
End Sub

Private Sub LogSaveCharacterDuration(ByVal StartTime As Long, ByVal QueryBreakdown As String, ByVal CharacterName As String, ByVal CharacterId As Long)
    Dim nowRaw As Long
    Dim totalElapsed As Double
    nowRaw = GetTickCountRaw()
    totalElapsed = TicksElapsed(StartTime, nowRaw)
    If totalElapsed > SAVE_CHARACTER_TIME_LIMIT_MS Then
        Call LogPerformance("Performance warning at: save character [" & CharacterName & "] id:" & CharacterId & _
                           " elapsed time: " & CLng(totalElapsed) & " breakdown: " & QueryBreakdown)
    End If
End Sub

Private Function GetIntegerArrayLength(ByRef arr() As Integer) As Long
    On Error Resume Next
    GetIntegerArrayLength = UBound(arr)
    If Err.Number <> 0 Then
        Err.Clear
        GetIntegerArrayLength = 0
    End If
End Function

Private Function GetIntegerArrayValue(ByRef arr() As Integer, ByVal index As Long) As Integer
    On Error Resume Next
    GetIntegerArrayValue = arr(index)
    If Err.Number <> 0 Then
        Err.Clear
        GetIntegerArrayValue = 0
    End If
End Function

Public Sub InitUserPersistSnapshot(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Call UpdateSavedSpells(UserIndex)
        Call UpdateSavedInventory(UserIndex)
        Call UpdateSavedBank(UserIndex)
        Call UpdateSavedSkills(UserIndex)
        Call UpdateSavedPets(UserIndex)
        Call UpdateSavedQuests(UserIndex)
        Call UpdateSavedQuestsDone(UserIndex)
    End With
End Sub

Public Function HaveSpellsChanged(ByVal UserIndex As Integer) As Boolean
    Dim i As Long
    With UserList(UserIndex)
        For i = 1 To MAXUSERHECHIZOS
            If .Stats.UserHechizos(i) <> .Persist.LastSpells(i) Then
                HaveSpellsChanged = True
                Exit Function
            End If
        Next i
    End With
End Function

Public Sub UpdateSavedSpells(ByVal UserIndex As Integer)
    Dim i As Long
    With UserList(UserIndex)
        For i = 1 To MAXUSERHECHIZOS
            .Persist.LastSpells(i) = .Stats.UserHechizos(i)
        Next i
    End With
End Sub

Public Function HasInventoryChanged(ByVal UserIndex As Integer) As Boolean
    Dim i As Long
    With UserList(UserIndex)
        For i = 1 To MAX_INVENTORY_SLOTS
            If .invent.Object(i).ObjIndex <> .Persist.LastInventory(i).ObjIndex _
               Or .invent.Object(i).amount <> .Persist.LastInventory(i).amount _
               Or .invent.Object(i).Equipped <> .Persist.LastInventory(i).Equipped _
               Or .invent.Object(i).ElementalTags <> .Persist.LastInventory(i).ElementalTags Then
                HasInventoryChanged = True
                Exit Function
            End If
        Next i
    End With
End Function

Public Sub UpdateSavedInventory(ByVal UserIndex As Integer)
    Dim i As Long
    With UserList(UserIndex)
        For i = 1 To MAX_INVENTORY_SLOTS
            .Persist.LastInventory(i) = .invent.Object(i)
        Next i
    End With
End Sub

Public Function HasBankChanged(ByVal UserIndex As Integer) As Boolean
    Dim i As Long
    With UserList(UserIndex)
        For i = 1 To MAX_BANCOINVENTORY_SLOTS
            If .BancoInvent.Object(i).ObjIndex <> .Persist.LastBank(i).ObjIndex _
               Or .BancoInvent.Object(i).amount <> .Persist.LastBank(i).amount _
               Or .BancoInvent.Object(i).ElementalTags <> .Persist.LastBank(i).ElementalTags Then
                HasBankChanged = True
                Exit Function
            End If
        Next i
    End With
End Function

Public Sub UpdateSavedBank(ByVal UserIndex As Integer)
    Dim i As Long
    With UserList(UserIndex)
        For i = 1 To MAX_BANCOINVENTORY_SLOTS
            .Persist.LastBank(i) = .BancoInvent.Object(i)
        Next i
    End With
End Sub

Public Function HaveSkillsChanged(ByVal UserIndex As Integer) As Boolean
    Dim i As Long
    With UserList(UserIndex)
        For i = 1 To NUMSKILLS
            If .Stats.SkillDirty(i) Then
                HaveSkillsChanged = True
                Exit Function
            End If
        Next i
    End With
End Function

Public Sub UpdateSavedSkills(ByVal UserIndex As Integer)
    Dim i As Long
    With UserList(UserIndex)
        For i = 1 To NUMSKILLS
            .Persist.LastSkills(i) = .Stats.UserSkills(i)
            .Stats.SkillDirty(i) = False
        Next i
    End With
End Sub

Public Function HavePetsChanged(ByVal UserIndex As Integer) As Boolean
    Dim i As Long
    Dim petType As Integer
    With UserList(UserIndex)
        For i = 1 To MAXMASCOTAS
            If IsValidNpcRef(.MascotasIndex(i)) Then
                If NpcList(.MascotasIndex(i).ArrayIndex).Contadores.TiempoExistencia = 0 Then
                    petType = .MascotasType(i)
                Else
                    petType = 0
                End If
            Else
                petType = .MascotasType(i)
            End If

            If petType <> .Persist.LastPetType(i) Then
                HavePetsChanged = True
                Exit Function
            End If
        Next i
    End With
End Function

Public Sub UpdateSavedPets(ByVal UserIndex As Integer)
    Dim i As Long
    Dim petType As Integer
    With UserList(UserIndex)
        For i = 1 To MAXMASCOTAS
            If IsValidNpcRef(.MascotasIndex(i)) Then
                If NpcList(.MascotasIndex(i).ArrayIndex).Contadores.TiempoExistencia = 0 Then
                    petType = .MascotasType(i)
                Else
                    petType = 0
                End If
            Else
                petType = .MascotasType(i)
            End If
            .Persist.LastPetType(i) = petType
        Next i
    End With
End Sub

Public Function HaveQuestsChanged(ByVal UserIndex As Integer) As Boolean
    Dim i As Long
    Dim k As Long
    Dim required As Integer
    With UserList(UserIndex)
        For i = 1 To MAXUSERQUESTS
            If .QuestStats.Quests(i).QuestIndex <> .Persist.LastQuests(i).QuestIndex Then
                HaveQuestsChanged = True
                Exit Function
            End If

            If .QuestStats.Quests(i).QuestIndex > 0 Then
                required = QuestList(.QuestStats.Quests(i).QuestIndex).RequiredNPCs
            Else
                required = 0
            End If

            If required <> GetIntegerArrayLength(.Persist.LastQuests(i).NPCsKilled) Then
                HaveQuestsChanged = True
                Exit Function
            End If

            For k = 1 To required
                If GetIntegerArrayValue(.QuestStats.Quests(i).NPCsKilled, k) <> GetIntegerArrayValue(.Persist.LastQuests(i).NPCsKilled, k) Then
                    HaveQuestsChanged = True
                    Exit Function
                End If
            Next k

            If .QuestStats.Quests(i).QuestIndex > 0 Then
                required = QuestList(.QuestStats.Quests(i).QuestIndex).RequiredTargetNPCs
            Else
                required = 0
            End If

            If required <> GetIntegerArrayLength(.Persist.LastQuests(i).NPCsTarget) Then
                HaveQuestsChanged = True
                Exit Function
            End If

            For k = 1 To required
                If GetIntegerArrayValue(.QuestStats.Quests(i).NPCsTarget, k) <> GetIntegerArrayValue(.Persist.LastQuests(i).NPCsTarget, k) Then
                    HaveQuestsChanged = True
                    Exit Function
                End If
            Next k
        Next i
    End With
End Function

Public Sub UpdateSavedQuests(ByVal UserIndex As Integer)
    Dim i As Long
    Dim k As Long
    Dim required As Integer
    With UserList(UserIndex)
        For i = 1 To MAXUSERQUESTS
            .Persist.LastQuests(i).QuestIndex = .QuestStats.Quests(i).QuestIndex

            If .QuestStats.Quests(i).QuestIndex > 0 Then
                required = QuestList(.QuestStats.Quests(i).QuestIndex).RequiredNPCs
            Else
                required = 0
            End If

            If required > 0 Then
                ReDim .Persist.LastQuests(i).NPCsKilled(1 To required)
                For k = 1 To required
                    .Persist.LastQuests(i).NPCsKilled(k) = GetIntegerArrayValue(.QuestStats.Quests(i).NPCsKilled, k)
                Next k
            Else
                ReDim .Persist.LastQuests(i).NPCsKilled(0)
            End If

            If .QuestStats.Quests(i).QuestIndex > 0 Then
                required = QuestList(.QuestStats.Quests(i).QuestIndex).RequiredTargetNPCs
            Else
                required = 0
            End If

            If required > 0 Then
                ReDim .Persist.LastQuests(i).NPCsTarget(1 To required)
                For k = 1 To required
                    .Persist.LastQuests(i).NPCsTarget(k) = GetIntegerArrayValue(.QuestStats.Quests(i).NPCsTarget, k)
                Next k
            Else
                ReDim .Persist.LastQuests(i).NPCsTarget(0)
            End If
        Next i
    End With
End Sub
Public Function HaveQuestsDoneChanged(ByRef U As t_User) As Boolean
    Dim i As Long
    Dim savedCount As Long

    savedCount = GetIntegerArrayLength(U.Persist.LastQuestsDone)

    If U.QuestStats.NumQuestsDone <> savedCount Then
        HaveQuestsDoneChanged = True
        Exit Function
    End If

    For i = 1 To U.QuestStats.NumQuestsDone
        If GetIntegerArrayValue(U.QuestStats.QuestsDone, i) <> _
           GetIntegerArrayValue(U.Persist.LastQuestsDone, i) Then
            HaveQuestsDoneChanged = True
            Exit Function
        End If
    Next i
End Function

Public Sub UpdateSavedQuestsDone(ByVal UserIndex As Integer)
    Dim i As Long
    With UserList(UserIndex)
        If .QuestStats.NumQuestsDone > 0 Then
            ReDim .Persist.LastQuestsDone(1 To .QuestStats.NumQuestsDone)
            For i = 1 To .QuestStats.NumQuestsDone
                .Persist.LastQuestsDone(i) = GetIntegerArrayValue(.QuestStats.QuestsDone, i)
            Next i
        Else
            ReDim .Persist.LastQuestsDone(0)
        End If
    End With
End Sub

Public Sub SaveChangesInUser(ByVal UserIndex As Integer)
    On Error GoTo ErrorHandler

    Dim PerformanceTimer As Long
    Dim TotalPerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    Call PerformanceTestStart(TotalPerformanceTimer)

    Dim QueryBreakdown As String

    Call Builder.Clear

    With UserList(UserIndex)

        If Not .flags.UserLogged Then
            Call LogDatabaseError("Error trying to save a user not logged in SaveChangesInUser")
            Exit Sub
        End If

        Call SaveCharacterMainDB(UserList(UserIndex), QueryBreakdown)
        Call PerformTimeLimitCheck(PerformanceTimer, "SaveChangesInUser [" & .name & "] main data id:" & .Id, 50)

        If HaveSpellsChanged(UserIndex) Then
            Call SaveCharacterSpellsDB(UserList(UserIndex), QueryBreakdown)
            Call PerformTimeLimitCheck(PerformanceTimer, "SaveChangesInUser [" & .name & "] spells update id:" & .Id, 50)
        End If

        If HasInventoryChanged(UserIndex) Then
            Call SaveCharacterInventoryDB(UserList(UserIndex), QueryBreakdown)
            Call PerformTimeLimitCheck(PerformanceTimer, "SaveChangesInUser [" & .name & "] inventory update id:" & .Id, 50)
        End If

        If HasBankChanged(UserIndex) Then
            Call SaveCharacterBankInventoryDB(UserList(UserIndex), QueryBreakdown)
            Call PerformTimeLimitCheck(PerformanceTimer, "SaveChangesInUser [" & .name & "] bank update id:" & .Id, 50)
        Else
            If LenB(QueryBreakdown) <> 0 Then QueryBreakdown = QueryBreakdown & "; "
            QueryBreakdown = QueryBreakdown & "save bank inventory: skipped"
        End If

        If HaveSkillsChanged(UserIndex) Then
            Call SaveCharacterSkillsDB(UserList(UserIndex), QueryBreakdown)
            Call PerformTimeLimitCheck(PerformanceTimer, "SaveChangesInUser [" & .name & "] skills update id:" & .Id, 50)
        End If

        If HavePetsChanged(UserIndex) Then
            Call SaveCharacterPetsDB(UserList(UserIndex), QueryBreakdown)
            Call PerformTimeLimitCheck(PerformanceTimer, "SaveChangesInUser [" & .name & "] pets update id:" & .Id, 50)
        End If

        If HaveQuestsChanged(UserIndex) Then
            Call SaveCharacterQuestsDB(UserList(UserIndex), QueryBreakdown, Builder)
            Call PerformTimeLimitCheck(PerformanceTimer, "SaveChangesInUser [" & .name & "] quests update id:" & .Id, 50)
        End If

        Call SaveCharacterQuestsDoneDB(UserList(UserIndex), QueryBreakdown, Builder)
        Call PerformTimeLimitCheck(PerformanceTimer, "SaveChangesInUser [" & .name & "] quests done update id:" & .Id, 50)

        Call SaveCharacterInventorySkinsDB(UserIndex, QueryBreakdown)
        Call PerformTimeLimitCheck(PerformanceTimer, "SaveChangesInUser [" & .name & "] inventory skins update id:" & .Id, 50)

        .Counters.LastSave = GetTickCountRaw()

        Call PerformTimeLimitCheck(TotalPerformanceTimer, "SaveChangesInUser [" & .name & "] total id:" & .Id, 50)
        Call LogSaveCharacterDuration(PerformanceTimer, QueryBreakdown, .name, .Id)

    End With

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in SaveChangesInUser. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description)
End Sub

Public Sub SaveNewCharacterDB(ByVal UserIndex As Integer)
    On Error GoTo ErrorHandler
    
#If LOGIN_STRESS_TEST = 1 Then
    If UserList(UserIndex).AccountID = -9999 Then Exit Sub
#End If
    Dim LoopC    As Long
    Dim ParamC   As Integer
    Dim Params() As Variant
    With UserList(UserIndex)
        Dim i As Integer
        i = 0
        ReDim Params(0 To 26)
        '  ************ Basic user data *******************
        Params(post_increment(i)) = .name
        Params(post_increment(i)) = .AccountID
        Params(post_increment(i)) = .Stats.ELV
        Params(post_increment(i)) = .Stats.Exp
        Params(post_increment(i)) = .genero
        Params(post_increment(i)) = .raza
        Params(post_increment(i)) = .clase
        Params(post_increment(i)) = .Hogar
        Params(post_increment(i)) = .Desc
        Params(post_increment(i)) = .Stats.GLD
        Params(post_increment(i)) = .Stats.SkillPts
        Params(post_increment(i)) = .pos.Map
        Params(post_increment(i)) = .pos.x
        Params(post_increment(i)) = .pos.y
        Params(post_increment(i)) = .Char.body
        Params(post_increment(i)) = .Char.head
        Params(post_increment(i)) = .Char.WeaponAnim
        Params(post_increment(i)) = .Char.CascoAnim
        Params(post_increment(i)) = .Char.ShieldAnim
        Params(post_increment(i)) = .Stats.MaxHp
        Params(post_increment(i)) = .Stats.MinHp
        Params(post_increment(i)) = .Stats.MinMAN
        Params(post_increment(i)) = .Stats.MinSta
        Params(post_increment(i)) = .Stats.MinHam
        Params(post_increment(i)) = .Stats.MinAGU
        Params(post_increment(i)) = .flags.Desnudo
        Params(post_increment(i)) = .Faccion.Status
        Call Query(QUERY_SAVE_MAINPJ, Params)
        ' Para recibir el ID del user
        Dim RS As ADODB.Recordset
        Set RS = Query("SELECT last_insert_rowid()")
        If RS Is Nothing Then
            .Id = 1
        Else
            .Id = val(RS.Fields(0).value)
        End If
        ' ******************* SPELLS **********************
        ReDim Params(MAXUSERHECHIZOS * 3 - 1)
        ParamC = 0
        For LoopC = 1 To MAXUSERHECHIZOS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = .Stats.UserHechizos(LoopC)
            ParamC = ParamC + 3
        Next LoopC
        Call Execute(QUERY_SAVE_SPELLS, Params)
        ' ******************* INVENTORY *******************
        ReDim Params(MAX_INVENTORY_SLOTS * 6 - 1)
        ParamC = 0
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = .invent.Object(LoopC).ObjIndex
            Params(ParamC + 3) = .invent.Object(LoopC).amount
            Params(ParamC + 4) = .invent.Object(LoopC).Equipped
            Params(ParamC + 5) = .invent.Object(LoopC).ElementalTags
            ParamC = ParamC + 6
        Next LoopC
        Call Execute(QUERY_SAVE_INVENTORY, Params)
        ' ******************* SKILLS *******************
        ReDim Params(NUMSKILLS * 3 - 1)
        ParamC = 0
        For LoopC = 1 To NUMSKILLS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = .Stats.UserSkills(LoopC)
            ParamC = ParamC + 3
        Next LoopC
        Call Execute(QUERY_SAVE_SKILLS, Params)
        ' ******************* QUESTS *******************
        ReDim Params(MAXUSERQUESTS * 2 - 1)
        ParamC = 0
        For LoopC = 1 To MAXUSERQUESTS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            ParamC = ParamC + 2
        Next LoopC
        Call Execute(QUERY_SAVE_QUESTS, Params)
        ' ******************* PETS ********************
        ReDim Params(MAXMASCOTAS * 3 - 1)
        ParamC = 0
        For LoopC = 1 To MAXMASCOTAS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = 0
            ParamC = ParamC + 3
        Next LoopC
        Call Execute(QUERY_SAVE_PETS, Params)
    End With
    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error en SaveNewUserDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description)
End Sub

Function LoadSkinsInventory(ByVal UserIndex As Integer) As Boolean

Dim bFixed                      As Boolean
Dim i                           As Integer
Dim tID                         As Long
Dim sQuery                      As String
Dim RS                          As ADODB.Recordset

    On Error GoTo ErrHandler
    If Not IsPatreon(UserIndex) Then Exit Function

    With UserList(UserIndex)

        sQuery = "SELECT type_skin, skin_id, skin_equipped FROM inventory_item_skins WHERE user_id=" & .Id
        Set RS = Query(sQuery, .Id)

        '@ Existe el Skins?
        If RS.EOF Or RS.BOF Then
            LoadSkinsInventory = False
            Set RS = Nothing
            Exit Function
        End If

        If RS.RecordCount > 0 Then
            i = 1
            Do While Not RS.EOF
                If CInt(RS.Fields("skin_id")) > 0 Then
                    .Invent_Skins.Object(i).ObjIndex = CInt(RS.Fields("skin_id"))

                    Select Case ObjData(.Invent_Skins.Object(i).ObjIndex).OBJType

                        Case e_OBJType.otSkinsSpells
                            If ObjData(.Invent_Skins.Object(i).ObjIndex).HechizoIndex > 0 And CBool(RS.Fields("skin_equipped")) Then
                                .Invent_Skins.Object(i).Equipped = True
                                .Invent_Skins.Object(i).Type = e_OBJType.otSkinsSpells
                                .Stats.UserSkinsHechizos(ObjData(.Invent_Skins.Object(i).ObjIndex).HechizoIndex) = ObjData(.Invent_Skins.Object(i).ObjIndex).CreaFX
                            End If

                        Case Else

                            If CBool(RS.Fields("skin_equipped")) Then
                                .Invent_Skins.Object(i).Type = ObjData(.Invent_Skins.Object(i).ObjIndex).OBJType
                                If CanEquipSkin(UserIndex, i, False) Then
                                    Call SkinEquip(UserIndex, i, .Invent_Skins.Object(i).ObjIndex, True)
                                End If
                            End If

                    End Select

                    Call WriteChangeSkinSlot(UserIndex, ObjData(.Invent_Skins.Object(i).ObjIndex).OBJType, i)

                Else
                    .Invent_Skins.Object(i).ObjIndex = 0
                    .Invent_Skins.Object(i).Equipped = False
                End If
                i = i + 1
                RS.MoveNext
            Loop
            .Invent_Skins.count = RS.RecordCount
        Else
            .Invent_Skins.count = 0
        End If
    End With

    Set RS = Nothing
    LoadSkinsInventory = True

    Exit Function

ErrHandler:

    Set RS = Nothing
    LoadSkinsInventory = False
    Call Logging.TraceError(Err.Number, Err.Description, "CharacterPersistence.LoadSkinsInventory of Módulo Nick: " & UserList(UserIndex).name, Erl())

End Function

Function SaveInventorySkins(ByVal UserIndex As Integer) As Boolean

Dim i                           As Integer
Dim UpsertCount                 As Long
Dim DeleteCount                 As Long
Dim SkinId                      As Long
Dim UpsertParams()              As Variant
Dim DeleteParams()              As Variant
Dim UpsertSql                   As cStringBuilder
Dim DeleteSql                   As cStringBuilder
Dim UpsertParamIndex            As Long
Dim DeleteParamIndex            As Long

    On Error GoTo SaveInventorySkins_Error

    With UserList(UserIndex)
        If .Id > 0 And .Invent_Skins.count > 0 Then
            For i = 1 To .Invent_Skins.count
                SkinId = .Invent_Skins.Object(i).ObjIndex

                If SkinId > 0 Then
                    If .Invent_Skins.Object(i).Deleted Then
                        DeleteCount = DeleteCount + 1
                    Else
                        UpsertCount = UpsertCount + 1
                    End If
                End If
            Next i

            If UpsertCount > 0 Then
                ReDim UpsertParams(0 To (UpsertCount * 4) - 1)
                Set UpsertSql = New cStringBuilder
                UpsertSql.Append "INSERT INTO inventory_item_skins (user_id, skin_id, type_skin, skin_equipped) VALUES "

                UpsertParamIndex = 0

                For i = 1 To .Invent_Skins.count
                    SkinId = .Invent_Skins.Object(i).ObjIndex

                    If SkinId > 0 And Not .Invent_Skins.Object(i).Deleted Then
                        If UpsertParamIndex > 0 Then
                            UpsertSql.Append ","
                        End If

                        UpsertSql.Append "(?, ?, ?, ?)"
                        UpsertParams(UpsertParamIndex) = .Id
                        UpsertParams(UpsertParamIndex + 1) = SkinId
                        UpsertParams(UpsertParamIndex + 2) = ObjData(SkinId).OBJType
                        UpsertParams(UpsertParamIndex + 3) = Abs(CInt(.Invent_Skins.Object(i).Equipped))
                        UpsertParamIndex = UpsertParamIndex + 4
                    End If
                Next i

                UpsertSql.Append " ON CONFLICT(user_id, skin_id) DO UPDATE SET type_skin = excluded.type_skin, skin_equipped = excluded.skin_equipped"
                Call Execute(UpsertSql.ToString, UpsertParams)
                Set UpsertSql = Nothing
            End If

            If DeleteCount > 0 Then
                ReDim DeleteParams(0 To DeleteCount)
                Set DeleteSql = New cStringBuilder
                DeleteSql.Append "DELETE FROM inventory_item_skins WHERE user_id = ? AND skin_id IN ("
                DeleteParams(0) = .Id

                DeleteParamIndex = 1

                For i = 1 To .Invent_Skins.count
                    SkinId = .Invent_Skins.Object(i).ObjIndex

                    If SkinId > 0 And .Invent_Skins.Object(i).Deleted Then
                        If DeleteParamIndex > 1 Then
                            DeleteSql.Append ","
                        End If

                        DeleteSql.Append "?"
                        DeleteParams(DeleteParamIndex) = SkinId
                        DeleteParamIndex = DeleteParamIndex + 1
                    End If
                Next i

                DeleteSql.Append ")"
                Call Execute(DeleteSql.ToString, DeleteParams)
                Set DeleteSql = Nothing
            End If

            SaveInventorySkins = True
        Else
            SaveInventorySkins = False
        End If
    End With

    On Error GoTo 0
    Exit Function

SaveInventorySkins_Error:
    SaveInventorySkins = False
    Call Logging.TraceError(Err.Number, Err.Description, "CharacterPersistence.SaveInventorySkins Nick: " & UserList(UserIndex).name, Erl())

End Function
