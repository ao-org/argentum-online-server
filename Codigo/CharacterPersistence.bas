Attribute VB_Name = "CharacterPersistence"
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

Option Base 0

Private Function db_load_house_key(ByRef user As t_User) As Boolean
    db_load_house_key = False
    With user
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

Public Function GetCharacterName(ByVal UserId As Long) As String
    On Error GoTo GetCharacterName_Err
        Dim RS As ADODB.Recordset
100     Set RS = Query("select name from user where id=?", UserId)
102     If RS Is Nothing Then Exit Function
104     GetCharacterName = RS!name
        Exit Function
GetCharacterName_Err:
    Call LogDatabaseError("Error en GetCharacterName: " & UserId & ". " & Err.Number & " - " & Err.Description & ". Línea: " & Erl)
End Function

Public Function LoadCharacterBank(ByVal UserIndex As Integer) As Boolean
    On Error GoTo LoadCharacterBank_Err
100     With UserList(UserIndex)
            Dim RS As ADODB.Recordset
            Dim counter As Long
            Set RS = Query("SELECT number, item_id, amount,elemental_tags FROM bank_item WHERE user_id = ?;", .ID)
            counter = 0
368         If Not RS Is Nothing Then
372             While Not RS.EOF
374                 With .BancoInvent.Object(RS!Number)
376                     .ObjIndex = IIf(RS!item_id <= UBound(ObjData), RS!item_id, 0)
378                     If .ObjIndex <> 0 Then
380                         If LenB(ObjData(.ObjIndex).name) Then
                                counter = counter + 1
382                             .amount = RS!amount
                                .ElementalTags = RS!elemental_tags
                            Else
384                             .ObjIndex = 0
                            End If
                        End If
                    End With
386                 RS.MoveNext
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
            Dim RS As ADODB.Recordset
            Dim counter As Long
            Dim SQLQuery As String
            Dim max_slots_to_load As Integer
            max_slots_to_load = get_num_inv_slots_from_tier(.Stats.tipoUsuario)
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
    Call LogDatabaseError("Error en LoadCharacterFromDB LoadCharacterInventory: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.Description & ". Línea: " & Erl)
End Function

Public Function LoadCharacterFromDB(ByVal userIndex As Integer) As Boolean
    On Error GoTo ErrorHandler
    
    Dim RS As ADODB.Recordset
    Dim counter As Long
    LoadCharacterFromDB = False
    
    With UserList(UserIndex)
        ' Load main character data using the user name.
        Set RS = Query(QUERY_LOAD_MAINPJ, .name)
        If RS Is Nothing Then Exit Function

        Debug.Assert .AccountID > -1
        If CLng(RS!account_id) <> .AccountID Then
            Call CloseSocket(UserIndex)
            Exit Function
        End If

        ' Set up the Patreon tier early (needed by subsequent initialization).
        .Stats.tipoUsuario = GetPatronTierFromAccountID(.AccountID)
        
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
        
        Call RegisterUserName(.Id, .name)
        Call Execute("update account set last_ip = ? where id = ?", .ConnectionDetails.IP, .AccountID)
        .Stats.Creditos = 0
        
        ' If the user is a patron-type, load the house key.
        If .Stats.tipoUsuario = tAventurero Or .Stats.tipoUsuario = tHeroe Or .Stats.tipoUsuario = tLeyenda Then
            Call db_load_house_key(UserList(UserIndex))
        End If
    End With

    LoadCharacterFromDB = True
    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error en LoadCharacterFromDB: " & UserList(UserIndex).name & ". " & _
                           Err.Number & " - " & Err.Description & ". Línea: " & Erl)
End Function


Private Sub SetupUserBasicInfo(ByRef User As t_User, ByRef RS As ADODB.Recordset)
    With User
        .Id = RS!Id
        .name = RS!name
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
    Dim RS As ADODB.Recordset
    Dim counter As Long
    counter = 0
    Set RS = Query("SELECT number, item_id, amount, elemental_tags FROM bank_item WHERE user_id = ?;", User.ID)
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
    Set RS = Query("SELECT number, value FROM skillpoint WHERE user_id = ?;", User.Id)
    If Not RS Is Nothing Then
        While Not RS.EOF
            User.Stats.UserSkills(RS!Number) = RS!value
            RS.MoveNext
        Wend
    End If
End Sub

Private Sub SetupUserQuests(ByRef User As t_User)
    Dim RS As ADODB.Recordset
    Dim LoopC As Byte
    Set RS = Query("SELECT number, quest_id, npcs, npcstarget FROM quest WHERE user_id = ?;", User.Id)
    If Not RS Is Nothing Then
        While Not RS.EOF
            If Not IsNull(RS!Number) Then
                User.QuestStats.Quests(RS!Number).QuestIndex = RS!quest_id
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
    Dim RS As ADODB.Recordset
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

Public Sub SaveCharacterDB(ByVal userIndex As Integer)

        On Error GoTo ErrorHandler
        Dim PerformanceTimer As Long
        Call PerformanceTestStart(PerformanceTimer)
        Dim Params() As Variant
        Dim LoopC As Long
        Dim ParamC As Long
100     Call Builder.Clear

102     With UserList(userIndex)
            Debug.Assert .flags.UserLogged = True
            If Not .flags.UserLogged Then
                Call LogDatabaseError("Error trying to save an user not logged in SaveCharacterDB")
                Exit Sub
            End If
              
104         ReDim Params(64)

            Dim i As Integer
        
106         Params(post_increment(i)) = .Name
108         Params(post_increment(i)) = .Stats.ELV
110         Params(post_increment(i)) = .Stats.Exp
112         Params(post_increment(i)) = .genero
114         Params(post_increment(i)) = .raza
116         Params(post_increment(i)) = .clase
118         Params(post_increment(i)) = .Hogar
120         Params(post_increment(i)) = .Desc
122         Params(post_increment(i)) = .Stats.GLD
124         Params(post_increment(i)) = .Stats.Banco
126         Params(post_increment(i)) = .Stats.SkillPts
128         Params(post_increment(i)) = .flags.MascotasGuardadas
130         Params(post_increment(i)) = .pos.map
132         Params(post_increment(i)) = .pos.x
134         Params(post_increment(i)) = .pos.y
136         Params(post_increment(i)) = .MENSAJEINFORMACION
138         Params(post_increment(i)) = .Char.body
140         Params(post_increment(i)) = .OrigChar.originalhead
142         Params(post_increment(i)) = .Char.WeaponAnim
144         Params(post_increment(i)) = .Char.CascoAnim
146         Params(post_increment(i)) = .Char.ShieldAnim
148         Params(post_increment(i)) = .Char.Heading
175         Params(post_increment(i)) = .Stats.MaxHp
176         Params(post_increment(i)) = .Stats.MinHp
180         Params(post_increment(i)) = .Stats.MinMAN
184         Params(post_increment(i)) = .Stats.MinSta
188         Params(post_increment(i)) = .Stats.MinHam
192         Params(post_increment(i)) = .Stats.MinAGU
200         Params(post_increment(i)) = .Stats.NPCsMuertos
202         Params(post_increment(i)) = .Stats.UsuariosMatados
203         Params(post_increment(i)) = .Stats.PuntosPesca
206         Params(post_increment(i)) = .Stats.ELO
208         Params(post_increment(i)) = .flags.Desnudo
210         Params(post_increment(i)) = .flags.Envenenado
212         Params(post_increment(i)) = .flags.Incinerado
218         Params(post_increment(i)) = .flags.Muerto
220         Params(post_increment(i)) = .flags.Navegando
222         Params(post_increment(i)) = .flags.Paralizado
224         Params(post_increment(i)) = .flags.Montado
226         Params(post_increment(i)) = .flags.Silenciado
228         Params(post_increment(i)) = .flags.MinutosRestantes
230         Params(post_increment(i)) = .flags.SegundosPasados
232         Params(post_increment(i)) = .flags.SpouseId
234         Params(post_increment(i)) = .Counters.Pena
236         Params(post_increment(i)) = .flags.VecesQueMoriste
246         Params(post_increment(i)) = .Faccion.ciudadanosMatados
248         Params(post_increment(i)) = .Faccion.CriminalesMatados
250         Params(post_increment(i)) = .Faccion.RecibioArmaduraReal
252         Params(post_increment(i)) = .Faccion.RecibioArmaduraCaos
258         Params(post_increment(i)) = .Faccion.RecompensasReal
259         Params(post_increment(i)) = .Faccion.FactionScore
260         Params(post_increment(i)) = .Faccion.RecompensasCaos
262         Params(post_increment(i)) = .Faccion.Reenlistadas
264         Params(post_increment(i)) = .Faccion.NivelIngreso
266         Params(post_increment(i)) = .Faccion.MatadosIngreso
270         Params(post_increment(i)) = .Faccion.Status
272         Params(post_increment(i)) = .GuildIndex
274         Params(post_increment(i)) = .ChatCombate
276         Params(post_increment(i)) = .ChatGlobal
280         Params(post_increment(i)) = .Stats.Advertencias
282         Params(post_increment(i)) = .flags.ReturnPos.map
284         Params(post_increment(i)) = .flags.ReturnPos.x
286         Params(post_increment(i)) = .flags.ReturnPos.y
287         Params(post_increment(i)) = .Stats.JineteLevel


            ' WHERE block
288         Params(post_increment(i)) = .ID
            
            Call Execute(QUERY_UPDATE_MAINPJ, Params)

            ' ************************** User spells *********************************
334             ReDim Params(MAXUSERHECHIZOS * 3 - 1)
336             ParamC = 0
            
338             For LoopC = 1 To MAXUSERHECHIZOS
340                 Params(ParamC) = .ID
342                 Params(ParamC + 1) = LoopC
344                 Params(ParamC + 2) = .Stats.UserHechizos(LoopC)
                
346                 ParamC = ParamC + 3
348             Next LoopC
                
                Call Execute(QUERY_UPSERT_SPELLS, Params)
            
            ' ************************** User inventory *********************************
            ReDim Params(MAX_INVENTORY_SLOTS * 6 - 1)
            ParamC = 0

370         For LoopC = 1 To MAX_INVENTORY_SLOTS
372             Params(ParamC) = .ID
374             Params(ParamC + 1) = LoopC
376             Params(ParamC + 2) = .Invent.Object(LoopC).objIndex
378             Params(ParamC + 3) = .Invent.Object(LoopC).amount
379             Params(ParamC + 4) = .Invent.Object(LoopC).Equipped
                Params(ParamC + 5) = .invent.Object(LoopC).ElementalTags
                
382             ParamC = ParamC + 6
384         Next LoopC

            Call Execute(QUERY_UPSERT_INVENTORY, Params)

            ' ************************** User bank inventory *********************************
402             ReDim Params(MAX_BANCOINVENTORY_SLOTS * 5 - 1)
404             ParamC = 0
            
406             For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
408                 Params(ParamC) = .ID
410                 Params(ParamC + 1) = LoopC
412                 Params(ParamC + 2) = .BancoInvent.Object(LoopC).objIndex
414                 Params(ParamC + 3) = .BancoInvent.Object(LoopC).amount
                    Params(ParamC + 4) = .BancoInvent.Object(LoopC).ElementalTags
416                 ParamC = ParamC + 5
418             Next LoopC
    
                Call Execute(QUERY_SAVE_BANCOINV, Params)


            ' ************************** User skills *********************************
436             ReDim Params(NUMSKILLS * 3 - 1)
438             ParamC = 0
            
440             For LoopC = 1 To NUMSKILLS
442                 Params(ParamC) = .ID
444                 Params(ParamC + 1) = LoopC
446                 Params(ParamC + 2) = .Stats.UserSkills(LoopC)
                
448                 ParamC = ParamC + 3
450             Next LoopC
        
                Call Execute(QUERY_UPSERT_SKILLS, Params)


            ' ************************** User pets *********************************
468             ReDim Params(MAXMASCOTAS * 3 - 1)
470             ParamC = 0
                Dim petType As Integer
    
472             For LoopC = 1 To MAXMASCOTAS
474                 Params(ParamC) = .ID
476                 Params(ParamC + 1) = LoopC
    

478                 If IsValidNpcRef(.MascotasIndex(LoopC)) Then
                
480                     If NpcList(.MascotasIndex(LoopC).ArrayIndex).Contadores.TiempoExistencia = 0 Then
482                         petType = .MascotasType(LoopC)
                        Else
484                         petType = 0
                        End If
                    Else
486                     petType = .MascotasType(LoopC)
                    End If
    
488                 Params(ParamC + 2) = petType
                
490                 ParamC = ParamC + 3
492             Next LoopC
                
                Call Execute(QUERY_UPSERT_PETS, Params)

            ' ************************** User quests *********************************
526             Builder.Append "REPLACE INTO quest (user_id, number, quest_id, npcs, npcstarget) VALUES "
            
                Dim Tmp As Integer, LoopK As Long
    
528             For LoopC = 1 To MAXUSERQUESTS
530                 Builder.Append "("
532                 Builder.Append .ID & ", "
534                 Builder.Append LoopC & ", "
536                 Builder.Append .QuestStats.Quests(LoopC).QuestIndex & ", '"
                
538                 If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
540                     Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredNPCs
    
542                     If Tmp Then
    
544                         For LoopK = 1 To Tmp
546                             Builder.Append CStr(.QuestStats.Quests(LoopC).NPCsKilled(LoopK))
                            
548                             If LoopK < Tmp Then
550                                 Builder.Append "-"
                                End If
    
552                         Next LoopK
                        
    
                        End If
    
                    End If
                
554                 Builder.Append "', '"
                
556                 If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
                    
558                     Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredTargetNPCs
                        
560                     For LoopK = 1 To Tmp
    
562                         Builder.Append CStr(.QuestStats.Quests(LoopC).NPCsTarget(LoopK))
                        
564                         If LoopK < Tmp Then
566                             Builder.Append "-"
                            End If
                    
568                     Next LoopK
                
                    End If
                
570                 Builder.Append "')"
    
572                 If LoopC < MAXUSERQUESTS Then
574                     Builder.Append ", "
                    End If
    
576             Next LoopC

                Call Execute(Builder.ToString())

584             Call Builder.Clear
                        
            ' ************************** User completed quests *********************************
        
               If .QuestStats.NumQuestsDone > 0 Then
                
                    ' Armamos la query con los placeholders
590                 Builder.Append "REPLACE INTO quest_done (user_id, quest_id) VALUES "
                
592                 For LoopC = 1 To .QuestStats.NumQuestsDone
594                     Builder.Append "(?, ?)"
                
596                     If LoopC < .QuestStats.NumQuestsDone Then
598                         Builder.Append ", "
                        End If
                
600                 Next LoopC

                    ' Metemos los parametros
604                 ReDim Params(.QuestStats.NumQuestsDone * 2 - 1)
606                 ParamC = 0
                
608                 For LoopC = 1 To .QuestStats.NumQuestsDone
610                     Params(ParamC) = .ID
612                     Params(ParamC + 1) = .QuestStats.QuestsDone(LoopC)
                    
614                     ParamC = ParamC + 2
616                 Next LoopC
        
                    Call Execute(Builder.ToString(), Params)

626                 Call Builder.Clear
                End If
                
        Call PerformTimeLimitCheck(PerformanceTimer, "save character id:" & .id, 50)
        End With
        
        Exit Sub

ErrorHandler:
636     Call LogDatabaseError("Error en SaveUserDatabase. UserName: " & UserList(userIndex).Name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveNewCharacterDB(ByVal userIndex As Integer)
        On Error GoTo ErrorHandler
        Dim LoopC As Long
        Dim ParamC As Integer
        Dim Params() As Variant
    
102     With UserList(userIndex)
        
            Dim i As Integer
            i = 0
104         ReDim Params(0 To 26)

            '  ************ Basic user data *******************
106         Params(post_increment(i)) = .Name
108         Params(post_increment(i)) = .AccountID
110         Params(post_increment(i)) = .Stats.ELV
112         Params(post_increment(i)) = .Stats.Exp
114         Params(post_increment(i)) = .genero
116         Params(post_increment(i)) = .raza
118         Params(post_increment(i)) = .clase
120         Params(post_increment(i)) = .Hogar
122         Params(post_increment(i)) = .Desc
124         Params(post_increment(i)) = .Stats.GLD
126         Params(post_increment(i)) = .Stats.SkillPts
128         Params(post_increment(i)) = .pos.map
130         Params(post_increment(i)) = .pos.x
132         Params(post_increment(i)) = .pos.y
134         Params(post_increment(i)) = .Char.body
136         Params(post_increment(i)) = .Char.head
138         Params(post_increment(i)) = .Char.WeaponAnim
140         Params(post_increment(i)) = .Char.CascoAnim
142         Params(post_increment(i)) = .Char.ShieldAnim
169         Params(post_increment(i)) = .Stats.MaxHp
170         Params(post_increment(i)) = .Stats.MinHp
174         Params(post_increment(i)) = .Stats.MinMAN
178         Params(post_increment(i)) = .Stats.MinSta
182         Params(post_increment(i)) = .Stats.MinHam
186         Params(post_increment(i)) = .Stats.MinAGU
194         Params(post_increment(i)) = .flags.Desnudo
196         Params(post_increment(i)) = .Faccion.Status
198         Call Query(QUERY_SAVE_MAINPJ, Params)

            ' Para recibir el ID del user
            Dim RS As ADODB.Recordset
            Set RS = Query("SELECT last_insert_rowid()")

202         If RS Is Nothing Then
204             .ID = 1
            Else
206             .ID = val(RS.Fields(0).value)
            End If
                
            ' ******************* SPELLS **********************
226         ReDim Params(MAXUSERHECHIZOS * 3 - 1)
228         ParamC = 0
        
230         For LoopC = 1 To MAXUSERHECHIZOS
232             Params(ParamC) = .ID
234             Params(ParamC + 1) = LoopC
236             Params(ParamC + 2) = .Stats.UserHechizos(LoopC)
            
238             ParamC = ParamC + 3
240         Next LoopC

            Call Execute(QUERY_SAVE_SPELLS, Params)
        
            ' ******************* INVENTORY *******************
244         ReDim Params(MAX_INVENTORY_SLOTS * 6 - 1)
246         ParamC = 0
        
248         For LoopC = 1 To MAX_INVENTORY_SLOTS
250             Params(ParamC) = .ID
252             Params(ParamC + 1) = LoopC
254             Params(ParamC + 2) = .Invent.Object(LoopC).objIndex
256             Params(ParamC + 3) = .Invent.Object(LoopC).amount
258             Params(ParamC + 4) = .Invent.Object(LoopC).Equipped
                Params(ParamC + 5) = .invent.Object(LoopC).ElementalTags
260             ParamC = ParamC + 6
262         Next LoopC
        
            Call Execute(QUERY_SAVE_INVENTORY, Params)
        
            ' ******************* SKILLS *******************
266         ReDim Params(NUMSKILLS * 3 - 1)
268         ParamC = 0
        
270         For LoopC = 1 To NUMSKILLS
272             Params(ParamC) = .ID
274             Params(ParamC + 1) = LoopC
276             Params(ParamC + 2) = .Stats.UserSkills(LoopC)
            
278             ParamC = ParamC + 3
280         Next LoopC
        
            Call Execute(QUERY_SAVE_SKILLS, Params)
        
            ' ******************* QUESTS *******************
284         ReDim Params(MAXUSERQUESTS * 2 - 1)
286         ParamC = 0
        
288         For LoopC = 1 To MAXUSERQUESTS
290             Params(ParamC) = .ID
292             Params(ParamC + 1) = LoopC
            
294             ParamC = ParamC + 2
296         Next LoopC
        
            Call Execute(QUERY_SAVE_QUESTS, Params)
        
            ' ******************* PETS ********************
300         ReDim Params(MAXMASCOTAS * 3 - 1)
302         ParamC = 0
        
304         For LoopC = 1 To MAXMASCOTAS
306             Params(ParamC) = .ID
308             Params(ParamC + 1) = LoopC
310             Params(ParamC + 2) = 0
            
312             ParamC = ParamC + 3
314         Next LoopC
    
            Call Execute(QUERY_SAVE_PETS, Params)
    
        End With

        Exit Sub

ErrorHandler:
    
322     Call LogDatabaseError("Error en SaveNewUserDatabase. UserName: " & UserList(userIndex).Name & ". " & Err.Number & " - " & Err.Description)

End Sub






