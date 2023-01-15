Attribute VB_Name = "CharacterPersistence"
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

Public Function LoadCharacterFromDB(ByVal userIndex As Integer) As Boolean
        Dim counter As Long
        On Error GoTo ErrorHandler
        LoadCharacterFromDB = False
         With UserList(userIndex)
            
            Dim RS As ADODB.Recordset
            Set RS = Query(QUERY_LOAD_MAINPJ, .Name)

            If RS Is Nothing Then Exit Function
            Debug.Assert .AccountID > -1 ' You need PYMMO =1 if this fails
            If (CLng(RS!account_id) <> .AccountID) Then
                Call CloseSocket(userIndex)
                LoadCharacterFromDB = False
                Exit Function
            End If
            
            If (RS!is_banned) Then
                Dim BanNick     As String
                Dim BaneoMotivo As String
                BanNick = RS!banned_by
                BaneoMotivo = RS!ban_reason
                
                If LenB(BanNick) = 0 Then BanNick = "*Error en la base de datos*"
                If LenB(BaneoMotivo) = 0 Then BaneoMotivo = "*No se registra el motivo del baneo.*"
            
                Call WriteShowMessageBox(UserIndex, "Se te ha prohibido la entrada al juego debido a " & BaneoMotivo & ". Esta decisión fue tomada por " & BanNick & ".")
            
                Call CloseSocket(userIndex)
                LoadCharacterFromDB = False
                Exit Function
            End If

            If (RS!is_locked_in_mao) Then
                Call WriteShowMessageBox(UserIndex, "El personaje que estás intentando loguear se encuentra en venta, para desbloquearlo deberás hacerlo desde la página web.")
                LoadCharacterFromDB = False
                Call CloseSocket(userIndex)
                Exit Function
            End If
            
        
            .InUse = True
            'Start setting data
106         .ID = RS!ID
108         .Name = RS!Name
110         .Stats.ELV = RS!level
112         .Stats.Exp = RS!Exp
114         .genero = RS!genre_id
116         .raza = RS!race_id
118         .clase = RS!class_id
120         .Hogar = RS!home_id
122         .Desc = RS!Description
124         .Stats.GLD = RS!gold
126         .Stats.Banco = RS!bank_gold
128         .Stats.SkillPts = RS!free_skillpoints
130         .pos.map = RS!pos_map
132         .pos.x = RS!pos_X
134         .pos.y = RS!pos_Y
136         .MENSAJEINFORMACION = RS!message_info
138         .OrigChar.body = RS!body_id
140         .OrigChar.head = RS!head_id
142         .OrigChar.WeaponAnim = RS!weapon_id
144         .OrigChar.CascoAnim = RS!helmet_id
146         .OrigChar.ShieldAnim = RS!shield_id
148         .OrigChar.Heading = RS!Heading
152         .Invent.ArmourEqpSlot = SanitizeNullValue(RS!slot_armour, 0)
154         .Invent.WeaponEqpSlot = SanitizeNullValue(RS!slot_weapon, 0)
156         .Invent.CascoEqpSlot = SanitizeNullValue(RS!slot_helmet, 0)
158         .Invent.EscudoEqpSlot = SanitizeNullValue(RS!slot_shield, 0)
160         .Invent.MunicionEqpSlot = SanitizeNullValue(RS!slot_ammo, 0)
162         .Invent.BarcoSlot = SanitizeNullValue(RS!slot_ship, 0)
164         .Invent.MonturaSlot = SanitizeNullValue(RS!slot_mount, 0)
166         .invent.DañoMagicoEqpSlot = SanitizeNullValue(RS!slot_dm, 0)
168         .Invent.ResistenciaEqpSlot = SanitizeNullValue(RS!slot_rm, 0)
170         .Invent.NudilloSlot = SanitizeNullValue(RS!slot_knuckles, 0)
172         .Invent.HerramientaEqpSlot = SanitizeNullValue(RS!slot_tool, 0)
174         .Invent.MagicoSlot = SanitizeNullValue(RS!slot_magic, 0)
176         .Stats.MinHp = RS!min_hp
178         .Stats.MaxHp = RS!max_hp
180         .Stats.MinMAN = RS!min_man
182         .Stats.MaxMAN = RS!max_man
184         .Stats.MinSta = RS!min_sta
186         .Stats.MaxSta = RS!max_sta
188         .Stats.MinHam = RS!min_ham
190         .Stats.MaxHam = RS!max_ham
192         .Stats.MinAGU = RS!min_sed
194         .Stats.MaxAGU = RS!max_sed
196         .Stats.MinHIT = RS!min_hit
198         .Stats.MaxHit = RS!max_hit
200         .Stats.NPCsMuertos = RS!killed_npcs
202         .Stats.UsuariosMatados = RS!killed_users
203         .Stats.PuntosPesca = RS!puntos_pesca
204         .Stats.InventLevel = RS!invent_level
206         .Stats.ELO = RS!ELO
208         .flags.Desnudo = RS!is_naked
210         .flags.Envenenado = RS!is_poisoned
211         .flags.Incinerado = RS!is_incinerated
212         .flags.Escondido = False
218         .flags.Ban = RS!is_banned
220         .flags.Muerto = RS!is_dead
222         .flags.Navegando = RS!is_sailing
224         .flags.Paralizado = RS!is_paralyzed
226         .flags.VecesQueMoriste = RS!deaths
228         .flags.Montado = RS!is_mounted
230         .flags.Pareja = RS!spouse
232         .flags.Casado = IIf(Len(.flags.Pareja) > 0, 1, 0)
234         .flags.Silenciado = RS!is_silenced
236         .flags.MinutosRestantes = RS!silence_minutes_left
238         .flags.SegundosPasados = RS!silence_elapsed_seconds
240         .flags.MascotasGuardadas = RS!pets_saved

246         .flags.ReturnPos.map = RS!return_map
248         .flags.ReturnPos.x = RS!return_x
250         .flags.ReturnPos.y = RS!return_y
        
252         .Counters.Pena = RS!counter_pena
        
254         .ChatGlobal = RS!chat_global
256         .ChatCombate = RS!chat_combate

260         .flags.Privilegios = .flags.Privilegios ' Or e_PlayerType.RoyalCouncil

270         .Faccion.ciudadanosMatados = RS!ciudadanos_matados
272         .Faccion.CriminalesMatados = RS!criminales_matados
274         .Faccion.RecibioArmaduraReal = RS!recibio_armadura_real
276         .Faccion.RecibioArmaduraCaos = RS!recibio_armadura_caos
278         .Faccion.RecibioExpInicialReal = RS!recibio_exp_real
280         .Faccion.RecibioExpInicialCaos = RS!recibio_exp_caos
282         .Faccion.RecompensasReal = RS!recompensas_real
284         .Faccion.RecompensasCaos = RS!recompensas_caos
286         .Faccion.Reenlistadas = RS!Reenlistadas
288         .Faccion.NivelIngreso = SanitizeNullValue(RS!nivel_ingreso, 0)
290         .Faccion.MatadosIngreso = SanitizeNullValue(RS!matados_ingreso, 0)
292         .Faccion.NextRecompensa = SanitizeNullValue(RS!siguiente_recompensa, 0)
294         .Faccion.Status = RS!Status

296         .GuildIndex = SanitizeNullValue(RS!Guild_Index, 0)
            .LastGuildRejection = SanitizeNullValue(RS!guild_rejected_because, vbNullString)
 
298         .Stats.Advertencias = RS!warnings

            'User spells
            Set RS = Query("SELECT number, spell_id FROM spell WHERE user_id = ?;", .ID)

316         If Not RS Is Nothing Then

320             While Not RS.EOF

322                 .Stats.UserHechizos(RS!Number) = RS!spell_id

324                 RS.MoveNext
                Wend

            End If

            'User pets
            Set RS = Query("SELECT number, pet_id FROM pet WHERE user_id = ?;", .ID)

328         If Not RS Is Nothing Then

332             While Not RS.EOF

334                 .MascotasType(RS!Number) = RS!pet_id
                
336                 If val(RS!pet_id) <> 0 Then
338                     .NroMascotas = .NroMascotas + 1

                    End If

340                 RS.MoveNext
                Wend

            End If

            'User inventory
            Set RS = Query("SELECT number, item_id, is_equipped, amount FROM inventory_item WHERE user_id = ?;", .ID)

            counter = 0
            
344         If Not RS Is Nothing Then

348             While Not RS.EOF

350                 With .Invent.Object(RS!Number)
352                     .objIndex = RS!item_id
                
354                     If .objIndex <> 0 Then
356                         If LenB(ObjData(.objIndex).Name) Then
                                counter = counter + 1
                                
358                             .amount = RS!amount
360                             .Equipped = RS!is_equipped
                            Else
362                             .objIndex = 0

                            End If

                        End If

                    End With

364                 RS.MoveNext
                Wend
                
                .Invent.NroItems = counter
            End If

            'User bank inventory
            Set RS = Query("SELECT number, item_id, amount FROM bank_item WHERE user_id = ?;", .ID)
            
            counter = 0
            
368         If Not RS Is Nothing Then

372             While Not RS.EOF

374                 With .BancoInvent.Object(RS!Number)
376                     .objIndex = RS!item_id
                
378                     If .objIndex <> 0 Then
380                         If LenB(ObjData(.objIndex).Name) Then
                                counter = counter + 1
                                
382                             .amount = RS!amount
                            Else
384                             .objIndex = 0

                            End If

                        End If

                    End With

386                 RS.MoveNext
                Wend
                
                .BancoInvent.NroItems = counter
            End If
            
            'User skills
            Set RS = Query("SELECT number, value FROM skillpoint WHERE user_id = ?;", .ID)

390         If Not RS Is Nothing Then

394             While Not RS.EOF

396                 .Stats.UserSkills(RS!Number) = RS!value

398                 RS.MoveNext
                Wend

            End If

            Dim LoopC As Byte
        
            'User quests
            Set RS = Query("SELECT number, quest_id, npcs, npcstarget FROM quest WHERE user_id = ?;", .ID)

402         If Not RS Is Nothing Then

406             While Not RS.EOF

408                 .QuestStats.Quests(RS!Number).QuestIndex = RS!quest_id
                
410                 If .QuestStats.Quests(RS!Number).QuestIndex > 0 Then
412                     If QuestList(.QuestStats.Quests(RS!Number).QuestIndex).RequiredNPCs Then

                            Dim NPCs() As String

414                         NPCs = Split(RS!NPCs, "-")
416                         ReDim .QuestStats.Quests(RS!Number).NPCsKilled(1 To QuestList(.QuestStats.Quests(RS!Number).QuestIndex).RequiredNPCs)

418                         For LoopC = 1 To QuestList(.QuestStats.Quests(RS!Number).QuestIndex).RequiredNPCs
420                             .QuestStats.Quests(RS!Number).NPCsKilled(LoopC) = val(NPCs(LoopC - 1))
422                         Next LoopC

                        End If
                    
424                     If QuestList(.QuestStats.Quests(RS!Number).QuestIndex).RequiredTargetNPCs Then

                            Dim NPCsTarget() As String

426                         NPCsTarget = Split(RS!NPCsTarget, "-")
428                         ReDim .QuestStats.Quests(RS!Number).NPCsTarget(1 To QuestList(.QuestStats.Quests(RS!Number).QuestIndex).RequiredTargetNPCs)

430                         For LoopC = 1 To QuestList(.QuestStats.Quests(RS!Number).QuestIndex).RequiredTargetNPCs
432                             .QuestStats.Quests(RS!Number).NPCsTarget(LoopC) = val(NPCsTarget(LoopC - 1))
434                         Next LoopC

                        End If

                    End If

436                 RS.MoveNext
                Wend

            End If
        
            'User quests done
            Set RS = Query("SELECT quest_id FROM quest_done WHERE user_id = ?;", .ID)

440         If Not RS Is Nothing Then
442             .QuestStats.NumQuestsDone = RS.RecordCount
                
                If (.QuestStats.NumQuestsDone > 0) Then
444                 ReDim .QuestStats.QuestsDone(1 To .QuestStats.NumQuestsDone)
    
448                 LoopC = 1
    
450                 While Not RS.EOF
                
452                     .QuestStats.QuestsDone(LoopC) = RS!quest_id
454                     LoopC = LoopC + 1
    
456                     RS.MoveNext
                    Wend
                End If
            End If


            
            Call Execute("update account set last_ip = ? where id = ?", .IP, .AccountID)
            
            .Stats.Creditos = 0
            Set RS = Query("Select is_active_patron from account where id = ?", .AccountID)
            If Not RS Is Nothing Then
                Dim tipo_usuario_db As Long
                tipo_usuario_db = RS!is_active_patron
                Select Case tipo_usuario_db
                    Case patron_tier_aventurero
                        .Stats.tipoUsuario = e_TipoUsuario.tAventurero
                    Case patron_tier_heroe
                        .Stats.tipoUsuario = e_TipoUsuario.tHeroe
                    Case patron_tier_leyenda
                        .Stats.tipoUsuario = e_TipoUsuario.tLeyenda
                    Case Else
                         .Stats.tipoUsuario = e_TipoUsuario.tNormal
                End Select
                
                If .Stats.tipoUsuario = tAventurero Or .Stats.tipoUsuario = tHeroe Or .Stats.tipoUsuario = tLeyenda Then
                    'Only load the house key if we are dealing with a patron
                    Call db_load_house_key(UserList(userIndex))
                End If
            Else
                'If we can't access patron info we set the user to normal
                .Stats.tipoUsuario = e_TipoUsuario.tNormal
            End If
        End With
        
        LoadCharacterFromDB = True
        
        Exit Function

ErrorHandler:
478     Call LogDatabaseError("Error en LoadCharacterFromDB: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description & ". Línea: " & Erl)

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

              
104         ReDim Params(84)

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
140         Params(post_increment(i)) = .OrigChar.head
142         Params(post_increment(i)) = .Char.WeaponAnim
144         Params(post_increment(i)) = .Char.CascoAnim
146         Params(post_increment(i)) = .Char.ShieldAnim
148         Params(post_increment(i)) = .Char.Heading
152         Params(post_increment(i)) = .Invent.ArmourEqpSlot
154         Params(post_increment(i)) = .Invent.WeaponEqpSlot
156         Params(post_increment(i)) = .Invent.EscudoEqpSlot
158         Params(post_increment(i)) = .Invent.CascoEqpSlot
160         Params(post_increment(i)) = .Invent.MunicionEqpSlot
162         Params(post_increment(i)) = .invent.DañoMagicoEqpSlot
164         Params(post_increment(i)) = .Invent.ResistenciaEqpSlot
166         Params(post_increment(i)) = .Invent.HerramientaEqpSlot
168         Params(post_increment(i)) = .Invent.MagicoSlot
170         Params(post_increment(i)) = .Invent.NudilloSlot
172         Params(post_increment(i)) = .Invent.BarcoSlot
174         Params(post_increment(i)) = .Invent.MonturaSlot
176         Params(post_increment(i)) = .Stats.MinHp
178         Params(post_increment(i)) = .Stats.MaxHp
180         Params(post_increment(i)) = .Stats.MinMAN
182         Params(post_increment(i)) = .Stats.MaxMAN
184         Params(post_increment(i)) = .Stats.MinSta
186         Params(post_increment(i)) = .Stats.MaxSta
188         Params(post_increment(i)) = .Stats.MinHam
190         Params(post_increment(i)) = .Stats.MaxHam
192         Params(post_increment(i)) = .Stats.MinAGU
194         Params(post_increment(i)) = .Stats.MaxAGU
196         Params(post_increment(i)) = .Stats.MinHIT
198         Params(post_increment(i)) = .Stats.MaxHit
200         Params(post_increment(i)) = .Stats.NPCsMuertos
202         Params(post_increment(i)) = .Stats.UsuariosMatados
203         Params(post_increment(i)) = .Stats.PuntosPesca
204         Params(post_increment(i)) = .Stats.InventLevel
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
232         Params(post_increment(i)) = .flags.Pareja
234         Params(post_increment(i)) = .Counters.Pena
236         Params(post_increment(i)) = .flags.VecesQueMoriste
246         Params(post_increment(i)) = .Faccion.ciudadanosMatados
248         Params(post_increment(i)) = .Faccion.CriminalesMatados
250         Params(post_increment(i)) = .Faccion.RecibioArmaduraReal
252         Params(post_increment(i)) = .Faccion.RecibioArmaduraCaos
254         Params(post_increment(i)) = .Faccion.RecibioExpInicialReal
256         Params(post_increment(i)) = .Faccion.RecibioExpInicialCaos
258         Params(post_increment(i)) = .Faccion.RecompensasReal
260         Params(post_increment(i)) = .Faccion.RecompensasCaos
262         Params(post_increment(i)) = .Faccion.Reenlistadas
264         Params(post_increment(i)) = .Faccion.NivelIngreso
266         Params(post_increment(i)) = .Faccion.MatadosIngreso
268         Params(post_increment(i)) = .Faccion.NextRecompensa
270         Params(post_increment(i)) = .Faccion.Status
272         Params(post_increment(i)) = .GuildIndex
274         Params(post_increment(i)) = .ChatCombate
276         Params(post_increment(i)) = .ChatGlobal
280         Params(post_increment(i)) = .Stats.Advertencias
282         Params(post_increment(i)) = .flags.ReturnPos.map
284         Params(post_increment(i)) = .flags.ReturnPos.x
286         Params(post_increment(i)) = .flags.ReturnPos.y

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
366             ReDim Params(MAX_INVENTORY_SLOTS * 5 - 1)
368             ParamC = 0
            
370             For LoopC = 1 To MAX_INVENTORY_SLOTS
372                 Params(ParamC) = .ID
374                 Params(ParamC + 1) = LoopC
376                 Params(ParamC + 2) = .Invent.Object(LoopC).objIndex
378                 Params(ParamC + 3) = .Invent.Object(LoopC).amount
380                 Params(ParamC + 4) = .Invent.Object(LoopC).Equipped
                
382                 ParamC = ParamC + 5
384             Next LoopC

                Call Execute(QUERY_UPSERT_INVENTORY, Params)

            ' ************************** User bank inventory *********************************
402             ReDim Params(MAX_BANCOINVENTORY_SLOTS * 4 - 1)
404             ParamC = 0
            
406             For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
408                 Params(ParamC) = .ID
410                 Params(ParamC + 1) = LoopC
412                 Params(ParamC + 2) = .BancoInvent.Object(LoopC).objIndex
414                 Params(ParamC + 3) = .BancoInvent.Object(LoopC).amount
                
416                 ParamC = ParamC + 4
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
104         ReDim Params(0 To 44)

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
146         Params(post_increment(i)) = .Invent.ArmourEqpSlot
148         Params(post_increment(i)) = .Invent.WeaponEqpSlot
150         Params(post_increment(i)) = .Invent.EscudoEqpSlot
152         Params(post_increment(i)) = .Invent.CascoEqpSlot
154         Params(post_increment(i)) = .Invent.MunicionEqpSlot
156         Params(post_increment(i)) = .invent.DañoMagicoEqpSlot
158         Params(post_increment(i)) = .Invent.ResistenciaEqpSlot
160         Params(post_increment(i)) = .Invent.HerramientaEqpSlot
162         Params(post_increment(i)) = .Invent.MagicoSlot
164         Params(post_increment(i)) = .Invent.NudilloSlot
166         Params(post_increment(i)) = .Invent.BarcoSlot
168         Params(post_increment(i)) = .Invent.MonturaSlot
170         Params(post_increment(i)) = .Stats.MinHp
172         Params(post_increment(i)) = .Stats.MaxHp
174         Params(post_increment(i)) = .Stats.MinMAN
176         Params(post_increment(i)) = .Stats.MaxMAN
178         Params(post_increment(i)) = .Stats.MinSta
180         Params(post_increment(i)) = .Stats.MaxSta
182         Params(post_increment(i)) = .Stats.MinHam
184         Params(post_increment(i)) = .Stats.MaxHam
186         Params(post_increment(i)) = .Stats.MinAGU
188         Params(post_increment(i)) = .Stats.MaxAGU
190         Params(post_increment(i)) = .Stats.MinHIT
192         Params(post_increment(i)) = .Stats.MaxHit
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
244         ReDim Params(MAX_INVENTORY_SLOTS * 5 - 1)
246         ParamC = 0
        
248         For LoopC = 1 To MAX_INVENTORY_SLOTS
250             Params(ParamC) = .ID
252             Params(ParamC + 1) = LoopC
254             Params(ParamC + 2) = .Invent.Object(LoopC).objIndex
256             Params(ParamC + 3) = .Invent.Object(LoopC).amount
258             Params(ParamC + 4) = .Invent.Object(LoopC).Equipped
            
260             ParamC = ParamC + 5
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






