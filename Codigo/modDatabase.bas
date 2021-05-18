Attribute VB_Name = "Database"
'Argentum Online Libre
'Database connection module
'Obtained from GS-Zone
'Adapted and modified by Juan Andres Dalmasso (CHOTS)
'September 2018
'Rewrited for Argentum20 by Alexis Caraballo (WyroX)
'October 2020

Option Explicit

Public Database_Enabled     As Boolean
Public Database_DataSource  As String
Public Database_Host        As String
Public Database_Name        As String
Public Database_Username    As String
Public Database_Password    As String

Private Database_Connection As ADODB.Connection
Private Command             As ADODB.Command
Public QueryData           As ADODB.Recordset
Private RecordsAffected     As Long

Private QueryBuilder        As cStringBuilder
Private ConnectedOnce       As Boolean

Public Sub Database_Connect()

        '************************************************************************************
        'Author: Juan Andres Dalmasso
        'Last Modification: 17/10/2020
        '21/09/2019 Jopi - Agregue soporte a conexion via DSN. Solo para usuarios avanzados.
        '17/10/2020 WyroX - Agrego soporte a multiples statements en la misma query
        '************************************************************************************
        On Error GoTo ErrorHandler
 
100     Set Database_Connection = New ADODB.Connection
    
102     If Len(Database_DataSource) <> 0 Then
    
104         Database_Connection.ConnectionString = "DATA SOURCE=" & Database_DataSource & ";"
        
        Else
    
106         Database_Connection.ConnectionString = "DRIVER={MySQL ODBC 8.0 ANSI Driver};" & _
                                                   "SERVER=" & Database_Host & ";" & _
                                                   "DATABASE=" & Database_Name & ";" & _
                                                   "USER=" & Database_Username & ";" & _
                                                   "PASSWORD=" & Database_Password & ";" & _
                                                   "OPTION=3;MULTI_STATEMENTS=1"
                                               
        End If
    
108     Debug.Print Database_Connection.ConnectionString
    
110     Database_Connection.CursorLocation = adUseClient
    
112     Call Database_Connection.Open

114     ConnectedOnce = True

        Exit Sub
    
ErrorHandler:
116     Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description)
    
118     If Not ConnectedOnce Then
120         Call MsgBox("No se pudo conectar a la base de datos. Mas información en logs/Database.log", vbCritical, "OBDC - Error")
122         Call CerrarServidor
        End If

End Sub

Public Sub Database_Close()

        '***************************************************
        'Author: Juan Andres Dalmasso
        'Last Modification: 18/09/2018
        'Nota by WyroX: Cerrar la conexion tambien libera
        'los recursos y cierra los RecordSet generados
        '***************************************************
        On Error GoTo ErrorHandler
    
100     Set Command = Nothing
     
102     Call Database_Connection.Close
    
104     Set Database_Connection = Nothing
     
        Exit Sub
     
ErrorHandler:
106     Call LogDatabaseError("Unable to close Mysql Database: " & Err.Number & " - " & Err.Description)

End Sub


Public Sub SaveNewUserDatabase(ByVal UserIndex As Integer)

        On Error GoTo ErrorHandler
    
        Dim LoopC As Long
        Dim ParamC As Long
        Dim Params() As Variant
    
        'Constructor de queries.
        'Me permite concatenar strings MUCHO MAS rapido
100     Set QueryBuilder = New cStringBuilder
    
102     With UserList(UserIndex)
        
104         ReDim Params(45)
        
            '  ************ Basic user data *******************
106         Params(0) = .name
108         Params(1) = .AccountId
110         Params(2) = .Stats.ELV
112         Params(3) = .Stats.Exp
114         Params(4) = .genero
116         Params(5) = .raza
118         Params(6) = .clase
120         Params(7) = .Hogar
122         Params(8) = .Desc
124         Params(9) = .Stats.GLD
126         Params(10) = .Stats.SkillPts
128         Params(11) = .Pos.Map
130         Params(12) = .Pos.X
132         Params(13) = .Pos.Y
134         Params(14) = .Char.Body
136         Params(15) = .Char.Head
138         Params(16) = .Char.WeaponAnim
140         Params(17) = .Char.CascoAnim
142         Params(18) = .Char.ShieldAnim
144         Params(19) = .Invent.NroItems
146         Params(20) = .Invent.ArmourEqpSlot
148         Params(21) = .Invent.WeaponEqpSlot
150         Params(22) = .Invent.EscudoEqpSlot
152         Params(23) = .Invent.CascoEqpSlot
154         Params(24) = .Invent.MunicionEqpSlot
156         Params(25) = .Invent.DañoMagicoEqpSlot
158         Params(26) = .Invent.ResistenciaEqpSlot
160         Params(27) = .Invent.HerramientaEqpSlot
162         Params(28) = .Invent.MagicoSlot
164         Params(29) = .Invent.NudilloSlot
166         Params(30) = .Invent.BarcoSlot
168         Params(31) = .Invent.MonturaSlot
170         Params(32) = .Stats.MinHp
172         Params(33) = .Stats.MaxHp
174         Params(34) = .Stats.MinMAN
176         Params(35) = .Stats.MaxMAN
178         Params(36) = .Stats.MinSta
180         Params(37) = .Stats.MaxSta
182         Params(38) = .Stats.MinHam
184         Params(39) = .Stats.MaxHam
186         Params(40) = .Stats.MinAGU
188         Params(41) = .Stats.MaxAGU
190         Params(42) = .Stats.MinHIT
192         Params(43) = .Stats.MaxHit
194         Params(44) = .flags.Desnudo
196         Params(45) = .Faccion.Status
        
198         Call MakeQuery(QUERY_SAVE_MAINPJ, True, Params)

            ' Para recibir el ID del user
200         Call MakeQuery("SELECT LAST_INSERT_ID();", False)

202         If QueryData Is Nothing Then
204             .Id = 1
            Else
206             .Id = val(QueryData.Fields(0).Value)
            End If
        
            ' ******************* ATRIBUTOS *******************
208         ReDim Params(NUMATRIBUTOS * 3 - 1)
210         ParamC = 0
        
212         For LoopC = 1 To NUMATRIBUTOS
214             Params(ParamC) = .Id
216             Params(ParamC + 1) = LoopC
218             Params(ParamC + 2) = .Stats.UserAtributos(LoopC)
            
220             ParamC = ParamC + 3
222         Next LoopC
        
224         Call MakeQuery(QUERY_SAVE_ATTRIBUTES, True, Params)
        
            ' ******************* SPELLS **********************
226         ReDim Params(MAXUSERHECHIZOS * 3 - 1)
228         ParamC = 0
        
230         For LoopC = 1 To MAXUSERHECHIZOS
232             Params(ParamC) = .Id
234             Params(ParamC + 1) = LoopC
236             Params(ParamC + 2) = .Stats.UserHechizos(LoopC)
            
238             ParamC = ParamC + 3
240         Next LoopC

242         Call MakeQuery(QUERY_SAVE_SPELLS, True, Params)
        
            ' ******************* INVENTORY *******************
244         ReDim Params(MAX_INVENTORY_SLOTS * 5 - 1)
246         ParamC = 0
        
248         For LoopC = 1 To MAX_INVENTORY_SLOTS
250             Params(ParamC) = .Id
252             Params(ParamC + 1) = LoopC
254             Params(ParamC + 2) = .Invent.Object(LoopC).ObjIndex
256             Params(ParamC + 3) = .Invent.Object(LoopC).amount
258             Params(ParamC + 4) = .Invent.Object(LoopC).Equipped
            
260             ParamC = ParamC + 5
262         Next LoopC
        
264         Call MakeQuery(QUERY_SAVE_INVENTORY, True, Params)
        
            ' ******************* SKILLS *******************
266         ReDim Params(NUMSKILLS * 3 - 1)
268         ParamC = 0
        
270         For LoopC = 1 To NUMSKILLS
272             Params(ParamC) = .Id
274             Params(ParamC + 1) = LoopC
276             Params(ParamC + 2) = .Stats.UserSkills(LoopC)
            
278             ParamC = ParamC + 3
280         Next LoopC
        
282         Call MakeQuery(QUERY_SAVE_SKILLS, True, Params)
        
            ' ******************* QUESTS *******************
284         ReDim Params(MAXUSERQUESTS * 2 - 1)
286         ParamC = 0
        
288         For LoopC = 1 To MAXUSERQUESTS
290             Params(ParamC) = .Id
292             Params(ParamC + 1) = LoopC
            
294             ParamC = ParamC + 2
296         Next LoopC
        
298         Call MakeQuery(QUERY_SAVE_QUESTS, True, Params)
        
            ' ******************* PETS ********************
300         ReDim Params(MAXMASCOTAS * 3 - 1)
302         ParamC = 0
        
304         For LoopC = 1 To MAXMASCOTAS
306             Params(ParamC) = .Id
308             Params(ParamC + 1) = LoopC
310             Params(ParamC + 2) = 0
            
312             ParamC = ParamC + 3
314         Next LoopC
    
316         Call MakeQuery(QUERY_SAVE_PETS, True, Params)
        
318         Set QueryBuilder = Nothing
    
        End With

        Exit Sub

ErrorHandler:
    
320     Set QueryBuilder = Nothing
    
322     Call LogDatabaseError("Error en SaveNewUserDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserDatabase(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean = False)

        On Error GoTo ErrorHandler
    
        Dim Params() As Variant
        Dim LoopC As Long
        Dim ParamC As Long
    
        'Constructor de queries.
        'Me permite concatenar strings MUCHO MAS rapido
100     Set QueryBuilder = New cStringBuilder

        'Basic user data
102     With UserList(UserIndex)
        
104         ReDim Params(91)
        
106         Params(0) = .name
108         Params(1) = .Stats.ELV
110         Params(2) = CLng(.Stats.Exp)
112         Params(3) = .genero
114         Params(4) = .raza
116         Params(5) = .clase
118         Params(6) = .Hogar
120         Params(7) = .Desc
122         Params(8) = .Stats.GLD
124         Params(9) = .Stats.Banco
126         Params(10) = .Stats.SkillPts
128         Params(11) = .flags.MascotasGuardadas
130         Params(12) = .Pos.Map
132         Params(13) = .Pos.X
134         Params(14) = .Pos.Y
136         Params(15) = .flags.lastMap
138         Params(16) = .MENSAJEINFORMACION
140         Params(17) = .Char.Body
142         Params(18) = .OrigChar.Head
144         Params(19) = .Char.WeaponAnim
146         Params(20) = .Char.CascoAnim
148         Params(21) = .Char.ShieldAnim
150         Params(22) = .Char.Heading
152         Params(23) = .Invent.NroItems
154         Params(24) = .Invent.ArmourEqpSlot
156         Params(25) = .Invent.WeaponEqpSlot
158         Params(26) = .Invent.EscudoEqpSlot
160         Params(27) = .Invent.CascoEqpSlot
162         Params(28) = .Invent.MunicionEqpSlot
164         Params(29) = .Invent.DañoMagicoEqpSlot
166         Params(30) = .Invent.ResistenciaEqpSlot
168         Params(31) = .Invent.HerramientaEqpSlot
170         Params(32) = .Invent.MagicoSlot
172         Params(33) = .Invent.NudilloSlot
174         Params(34) = .Invent.BarcoSlot
176         Params(35) = .Invent.MonturaSlot
178         Params(36) = .Stats.MinHp
180         Params(37) = .Stats.MaxHp
182         Params(38) = .Stats.MinMAN
184         Params(39) = .Stats.MaxMAN
186         Params(40) = .Stats.MinSta
188         Params(41) = .Stats.MaxSta
190         Params(42) = .Stats.MinHam
192         Params(43) = .Stats.MaxHam
194         Params(44) = .Stats.MinAGU
196         Params(45) = .Stats.MaxAGU
198         Params(46) = .Stats.MinHIT
200         Params(47) = .Stats.MaxHit
202         Params(48) = .Stats.NPCsMuertos
204         Params(49) = .Stats.UsuariosMatados
206         Params(50) = .Stats.InventLevel
208         Params(51) = .flags.Desnudo
210         Params(52) = .flags.Envenenado
212         Params(53) = .flags.Escondido
214         Params(54) = .flags.Hambre
216         Params(55) = .flags.Sed
218         Params(56) = .flags.Muerto
220         Params(57) = .flags.Navegando
222         Params(58) = .flags.Paralizado
224         Params(59) = .flags.Montado
226         Params(60) = .flags.Silenciado
228         Params(61) = .flags.MinutosRestantes
230         Params(62) = .flags.SegundosPasados
232         Params(63) = .flags.Pareja
234         Params(64) = .Counters.Pena
236         Params(65) = .flags.VecesQueMoriste
238         Params(66) = (.flags.Privilegios And PlayerType.RoyalCouncil)
240         Params(67) = (.flags.Privilegios And PlayerType.ChaosCouncil)
242         Params(68) = .Faccion.ArmadaReal
244         Params(69) = .Faccion.FuerzasCaos
246         Params(70) = .Faccion.ciudadanosMatados
248         Params(71) = .Faccion.CriminalesMatados
250         Params(72) = .Faccion.RecibioArmaduraReal
252         Params(73) = .Faccion.RecibioArmaduraCaos
254         Params(74) = .Faccion.RecibioExpInicialReal
256         Params(75) = .Faccion.RecibioExpInicialCaos
258         Params(76) = .Faccion.RecompensasReal
260         Params(77) = .Faccion.RecompensasCaos
262         Params(78) = .Faccion.Reenlistadas
264         Params(79) = .Faccion.NivelIngreso
266         Params(80) = .Faccion.MatadosIngreso
268         Params(81) = .Faccion.NextRecompensa
270         Params(82) = .Faccion.Status
272         Params(83) = .GuildIndex
274         Params(84) = .ChatCombate
276         Params(85) = .ChatGlobal
278         Params(86) = IIf(Logout, 0, 1)
280         Params(87) = .Stats.Advertencias
            Params(88) = .flags.ReturnPos.Map
            Params(89) = .flags.ReturnPos.X
            Params(90) = .flags.ReturnPos.Y
        
            ' WHERE block
282         Params(91) = .Id
        
284         Call MakeQuery(QUERY_UPDATE_MAINPJ, True, Params)
        
286         Call QueryBuilder.Clear

            ' ************************** User attributes ****************************
288         ReDim Params(NUMATRIBUTOS * 3 - 1)
290         ParamC = 0
        
292         For LoopC = 1 To NUMATRIBUTOS
294             Params(ParamC) = .Id
296             Params(ParamC + 1) = LoopC
298             Params(ParamC + 2) = .Stats.UserAtributosBackUP(LoopC)
            
300             ParamC = ParamC + 3
302         Next LoopC
        
304         Call MakeQuery(QUERY_UPSERT_ATTRIBUTES, True, Params)

            ' ************************** User spells *********************************
306         ReDim Params(MAXUSERHECHIZOS * 3 - 1)
308         ParamC = 0
        
310         For LoopC = 1 To MAXUSERHECHIZOS
312             Params(ParamC) = .Id
314             Params(ParamC + 1) = LoopC
316             Params(ParamC + 2) = .Stats.UserHechizos(LoopC)
            
318             ParamC = ParamC + 3
320         Next LoopC
        
322         Call MakeQuery(QUERY_UPSERT_SPELLS, True, Params)

            ' ************************** User inventory *********************************
324         ReDim Params(MAX_INVENTORY_SLOTS * 5 - 1)
326         ParamC = 0
        
328         For LoopC = 1 To MAX_INVENTORY_SLOTS
330             Params(ParamC) = .Id
332             Params(ParamC + 1) = LoopC
334             Params(ParamC + 2) = .Invent.Object(LoopC).ObjIndex
336             Params(ParamC + 3) = .Invent.Object(LoopC).amount
338             Params(ParamC + 4) = .Invent.Object(LoopC).Equipped
            
340             ParamC = ParamC + 5
342         Next LoopC
        
344         Call MakeQuery(QUERY_UPSERT_INVENTORY, True, Params)

            ' ************************** User bank inventory *********************************
346         ReDim Params(MAX_BANCOINVENTORY_SLOTS * 4 - 1)
348         ParamC = 0
        
350         For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
352             Params(ParamC) = .Id
354             Params(ParamC + 1) = LoopC
356             Params(ParamC + 2) = .BancoInvent.Object(LoopC).ObjIndex
358             Params(ParamC + 3) = .BancoInvent.Object(LoopC).amount
            
360             ParamC = ParamC + 4
362         Next LoopC
        
364         Call MakeQuery(QUERY_SAVE_BANCOINV, True, Params)

            ' ************************** User skills *********************************
366         ReDim Params(NUMSKILLS * 3 - 1)
368         ParamC = 0
        
370         For LoopC = 1 To NUMSKILLS
372             Params(ParamC) = .Id
374             Params(ParamC + 1) = LoopC
376             Params(ParamC + 2) = .Stats.UserSkills(LoopC)
            
378             ParamC = ParamC + 3
380         Next LoopC
        
382         Call MakeQuery(QUERY_UPSERT_SKILLS, True, Params)

            ' ************************** User pets *********************************
384         ReDim Params(MAXMASCOTAS * 3 - 1)
386         ParamC = 0
            Dim petType As Integer

388         For LoopC = 1 To MAXMASCOTAS
390             Params(ParamC) = .Id
392             Params(ParamC + 1) = LoopC

                'CHOTS | I got this logic from SaveUserToCharfile
394             If .MascotasIndex(LoopC) > 0 Then
            
396                 If NpcList(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
398                     petType = .MascotasType(LoopC)
                    Else
400                     petType = 0
                    End If

                Else
402                 petType = .MascotasType(LoopC)

                End If

404             Params(ParamC + 2) = petType
            
406             ParamC = ParamC + 3
408         Next LoopC
        
410         Call MakeQuery(QUERY_UPSERT_PETS, True, Params)
        
            ' ************************** User connection logs *********************************
        
            'Agrego ip del user
412         Call MakeQuery(QUERY_SAVE_CONNECTION, True, .Id, .ip)
        
            'Borro la mas vieja si hay mas de 5 (WyroX: si alguien sabe una forma mejor de hacerlo me avisa)
414         Call MakeQuery(QUERY_DELETE_LAST_CONNECTIONS, True, .Id, .Id)
        
            ' ************************** User quests *********************************
416         QueryBuilder.Append "INSERT INTO quest (user_id, number, quest_id, npcs, npcstarget) VALUES "
        
            Dim Tmp As Integer, LoopK As Long

418         For LoopC = 1 To MAXUSERQUESTS
420             QueryBuilder.Append "("
422             QueryBuilder.Append .Id & ", "
424             QueryBuilder.Append LoopC & ", "
426             QueryBuilder.Append .QuestStats.Quests(LoopC).QuestIndex & ", '"
            
428             If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
430                 Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredNPCs

432                 If Tmp Then

434                     For LoopK = 1 To Tmp
436                         QueryBuilder.Append CStr(.QuestStats.Quests(LoopC).NPCsKilled(LoopK))
                        
438                         If LoopK < Tmp Then
440                             QueryBuilder.Append "-"
                            End If

442                     Next LoopK
                    

                    End If

                End If
            
444             QueryBuilder.Append "', '"
            
446             If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
                
448                 Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredTargetNPCs
                    
450                 For LoopK = 1 To Tmp

452                     QueryBuilder.Append CStr(.QuestStats.Quests(LoopC).NPCsTarget(LoopK))
                    
454                     If LoopK < Tmp Then
456                         QueryBuilder.Append "-"
                        End If
                
458                 Next LoopK
            
                End If
            
460             QueryBuilder.Append "')"

462             If LoopC < MAXUSERQUESTS Then
464                 QueryBuilder.Append ", "
                End If

466         Next LoopC
        
468         QueryBuilder.Append " ON DUPLICATE KEY UPDATE quest_id=VALUES(quest_id), npcs=VALUES(npcs); "
        
470         Call MakeQuery(QueryBuilder.ToString, True)
        
472         Call QueryBuilder.Clear
        
            ' ************************** User completed quests *********************************
474         If .QuestStats.NumQuestsDone > 0 Then
            
                ' Armamos la query con los placeholders
476             QueryBuilder.Append "INSERT INTO quest_done (user_id, quest_id) VALUES "
            
478             For LoopC = 1 To .QuestStats.NumQuestsDone
480                 QueryBuilder.Append "(?, ?)"
            
482                 If LoopC < .QuestStats.NumQuestsDone Then
484                     QueryBuilder.Append ", "
                    End If
            
486             Next LoopC
                    
488             QueryBuilder.Append " ON DUPLICATE KEY UPDATE quest_id=VALUES(quest_id); "
            
                ' Metemos los parametros
490             ReDim Params(.QuestStats.NumQuestsDone * 2 - 1)
492             ParamC = 0
            
494             For LoopC = 1 To .QuestStats.NumQuestsDone
496                 Params(ParamC) = .Id
498                 Params(ParamC + 1) = .QuestStats.QuestsDone(LoopC)
                
500                 ParamC = ParamC + 2
502             Next LoopC
            
                ' Mandamos la query
504             Call MakeQuery(QueryBuilder.ToString, True, Params)
        
            End If

            ' ************************** User logout *********************************
            ' Si deslogueó, actualizo la cuenta
506         If Logout Then
508             Call MakeQuery("UPDATE account SET logged = logged - 1 WHERE id = ?", True, .AccountId)
            End If

        End With
    
510     Set QueryBuilder = Nothing
    
        Exit Sub

ErrorHandler:

512     Set QueryBuilder = Nothing
    
514     Call LogDatabaseError("Error en SaveUserDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description)

End Sub

Sub LoadUserDatabase(ByVal UserIndex As Integer)

        On Error GoTo ErrorHandler

        'Basic user data
100     With UserList(UserIndex)

102         Call MakeQuery("SELECT *, DATE_FORMAT(fecha_ingreso, '%Y-%m-%d') as 'fecha_ingreso_format' FROM user WHERE name = ?;", False, .name)

104         If QueryData Is Nothing Then Exit Sub

            'Start setting data
106         .Id = QueryData!Id
108         .name = QueryData!name
110         .Stats.ELV = QueryData!level
112         .Stats.Exp = QueryData!Exp
114         .genero = QueryData!genre_id
116         .raza = QueryData!race_id
118         .clase = QueryData!class_id
120         .Hogar = QueryData!home_id
122         .Desc = QueryData!Description
124         .Stats.GLD = QueryData!gold
126         .Stats.Banco = QueryData!bank_gold
128         .Stats.SkillPts = QueryData!free_skillpoints
            '.Counters.AsignedSkills = QueryData!assigned_skillpoints
130         .Pos.Map = QueryData!pos_map
132         .Pos.X = QueryData!pos_x
134         .Pos.Y = QueryData!pos_y
136         .flags.lastMap = QueryData!last_map
138         .MENSAJEINFORMACION = QueryData!message_info
140         .OrigChar.Body = QueryData!body_id
142         .OrigChar.Head = QueryData!head_id
144         .OrigChar.WeaponAnim = QueryData!weapon_id
146         .OrigChar.CascoAnim = QueryData!helmet_id
148         .OrigChar.ShieldAnim = QueryData!shield_id
150         .OrigChar.Heading = QueryData!Heading
152         .Invent.NroItems = QueryData!items_Amount
154         .Invent.ArmourEqpSlot = SanitizeNullValue(QueryData!slot_armour, 0)
156         .Invent.WeaponEqpSlot = SanitizeNullValue(QueryData!slot_weapon, 0)
158         .Invent.CascoEqpSlot = SanitizeNullValue(QueryData!slot_helmet, 0)
160         .Invent.EscudoEqpSlot = SanitizeNullValue(QueryData!slot_shield, 0)
162         .Invent.MunicionEqpSlot = SanitizeNullValue(QueryData!slot_ammo, 0)
164         .Invent.BarcoSlot = SanitizeNullValue(QueryData!slot_ship, 0)
166         .Invent.MonturaSlot = SanitizeNullValue(QueryData!slot_mount, 0)
168         .Invent.DañoMagicoEqpSlot = SanitizeNullValue(QueryData!slot_dm, 0)
170         .Invent.ResistenciaEqpSlot = SanitizeNullValue(QueryData!slot_rm, 0)
172         .Invent.NudilloSlot = SanitizeNullValue(QueryData!slot_knuckles, 0)
174         .Invent.HerramientaEqpSlot = SanitizeNullValue(QueryData!slot_tool, 0)
176         .Invent.MagicoSlot = SanitizeNullValue(QueryData!slot_magic, 0)
178         .Stats.MinHp = QueryData!min_hp
180         .Stats.MaxHp = QueryData!max_hp
182         .Stats.MinMAN = QueryData!min_man
184         .Stats.MaxMAN = QueryData!max_man
186         .Stats.MinSta = QueryData!min_sta
188         .Stats.MaxSta = QueryData!max_sta
190         .Stats.MinHam = QueryData!min_ham
192         .Stats.MaxHam = QueryData!max_ham
194         .Stats.MinAGU = QueryData!min_sed
196         .Stats.MaxAGU = QueryData!max_sed
198         .Stats.MinHIT = QueryData!min_hit
200         .Stats.MaxHit = QueryData!max_hit
202         .Stats.NPCsMuertos = QueryData!killed_npcs
204         .Stats.UsuariosMatados = QueryData!killed_users
206         .Stats.InventLevel = QueryData!invent_level
            '.Reputacion.AsesinoRep = QueryData!rep_asesino
            '.Reputacion.BandidoRep = QueryData!rep_bandido
            '.Reputacion.BurguesRep = QueryData!rep_burgues
            '.Reputacion.LadronesRep = QueryData!rep_ladron
            '.Reputacion.NobleRep = QueryData!rep_noble
            '.Reputacion.PlebeRep = QueryData!rep_plebe
            '.Reputacion.Promedio = QueryData!rep_average
208         .flags.Desnudo = QueryData!is_naked
210         .flags.Envenenado = QueryData!is_poisoned
212         .flags.Escondido = QueryData!is_hidden
214         .flags.Hambre = QueryData!is_hungry
216         .flags.Sed = QueryData!is_thirsty
218         .flags.Ban = QueryData!is_banned
220         .flags.Muerto = QueryData!is_dead
222         .flags.Navegando = QueryData!is_sailing
224         .flags.Paralizado = QueryData!is_paralyzed
226         .flags.VecesQueMoriste = QueryData!deaths
228         .flags.Montado = QueryData!is_mounted
230         .flags.Pareja = QueryData!spouse
232         .flags.Casado = IIf(Len(.flags.Pareja) > 0, 1, 0)
234         .flags.Silenciado = QueryData!is_silenced
236         .flags.MinutosRestantes = QueryData!silence_minutes_left
238         .flags.SegundosPasados = QueryData!silence_elapsed_seconds
240         .flags.MascotasGuardadas = QueryData!pets_saved
242         .flags.ScrollExp = 1 'TODO: sacar
244         .flags.ScrollOro = 1 'TODO: sacar

            .flags.ReturnPos.Map = QueryData!return_map
            .flags.ReturnPos.X = QueryData!return_x
            .flags.ReturnPos.Y = QueryData!return_y
        
246         .Counters.Pena = QueryData!counter_pena
        
248         .ChatGlobal = QueryData!chat_global
250         .ChatCombate = QueryData!chat_combate

252         If QueryData!pertenece_consejo_real Then
254             .flags.Privilegios = .flags.Privilegios Or PlayerType.RoyalCouncil

            End If

256         If QueryData!pertenece_consejo_caos Then
258             .flags.Privilegios = .flags.Privilegios Or PlayerType.ChaosCouncil

            End If

260         .Faccion.ArmadaReal = QueryData!pertenece_real
262         .Faccion.FuerzasCaos = QueryData!pertenece_caos
264         .Faccion.ciudadanosMatados = QueryData!ciudadanos_matados
266         .Faccion.CriminalesMatados = QueryData!criminales_matados
268         .Faccion.RecibioArmaduraReal = QueryData!recibio_armadura_real
270         .Faccion.RecibioArmaduraCaos = QueryData!recibio_armadura_caos
272         .Faccion.RecibioExpInicialReal = QueryData!recibio_exp_real
274         .Faccion.RecibioExpInicialCaos = QueryData!recibio_exp_caos
276         .Faccion.RecompensasReal = QueryData!recompensas_real
278         .Faccion.RecompensasCaos = QueryData!recompensas_caos
280         .Faccion.Reenlistadas = QueryData!Reenlistadas
282         .Faccion.NivelIngreso = SanitizeNullValue(QueryData!nivel_ingreso, 0)
284         .Faccion.MatadosIngreso = SanitizeNullValue(QueryData!matados_ingreso, 0)
286         .Faccion.NextRecompensa = SanitizeNullValue(QueryData!siguiente_recompensa, 0)
288         .Faccion.Status = QueryData!Status

290         .GuildIndex = SanitizeNullValue(QueryData!Guild_Index, 0)
        
292         .Stats.Advertencias = QueryData!warnings
        
            'User attributes
294         Call MakeQuery("SELECT * FROM attribute WHERE user_id = ?;", False, .Id)
    
296         If Not QueryData Is Nothing Then
298             QueryData.MoveFirst

300             While Not QueryData.EOF

302                 .Stats.UserAtributos(QueryData!Number) = QueryData!Value
304                 .Stats.UserAtributosBackUP(QueryData!Number) = .Stats.UserAtributos(QueryData!Number)

306                 QueryData.MoveNext
                Wend

            End If

            'User spells
308         Call MakeQuery("SELECT * FROM spell WHERE user_id = ?;", False, .Id)

310         If Not QueryData Is Nothing Then
312             QueryData.MoveFirst

314             While Not QueryData.EOF

316                 .Stats.UserHechizos(QueryData!Number) = QueryData!spell_id

318                 QueryData.MoveNext
                Wend

            End If

            'User pets
320         Call MakeQuery("SELECT * FROM pet WHERE user_id = ?;", False, .Id)

322         If Not QueryData Is Nothing Then
324             QueryData.MoveFirst

326             While Not QueryData.EOF

328                 .MascotasType(QueryData!Number) = QueryData!pet_id
                
330                 If val(QueryData!pet_id) <> 0 Then
332                     .NroMascotas = .NroMascotas + 1

                    End If

334                 QueryData.MoveNext
                Wend

            End If

            'User inventory
336         Call MakeQuery("SELECT * FROM inventory_item WHERE user_id = ?;", False, .Id)

338         If Not QueryData Is Nothing Then
340             QueryData.MoveFirst

342             While Not QueryData.EOF

344                 With .Invent.Object(QueryData!Number)
346                     .ObjIndex = QueryData!item_id
                
348                     If .ObjIndex <> 0 Then
350                         If LenB(ObjData(.ObjIndex).name) Then
352                             .amount = QueryData!amount
354                             .Equipped = QueryData!is_equipped
                            Else
356                             .ObjIndex = 0

                            End If

                        End If

                    End With

358                 QueryData.MoveNext
                Wend

            End If

            'User bank inventory
360         Call MakeQuery("SELECT * FROM bank_item WHERE user_id = ?;", False, .Id)

362         If Not QueryData Is Nothing Then
364             QueryData.MoveFirst

366             While Not QueryData.EOF

368                 With .BancoInvent.Object(QueryData!Number)
370                     .ObjIndex = QueryData!item_id
                
372                     If .ObjIndex <> 0 Then
374                         If LenB(ObjData(.ObjIndex).name) Then
376                             .amount = QueryData!amount
                            Else
378                             .ObjIndex = 0

                            End If

                        End If

                    End With

380                 QueryData.MoveNext
                Wend

            End If

            'User skills
382         Call MakeQuery("SELECT * FROM skillpoint WHERE user_id = ?;", False, .Id)

384         If Not QueryData Is Nothing Then
386             QueryData.MoveFirst

388             While Not QueryData.EOF

390                 .Stats.UserSkills(QueryData!Number) = QueryData!Value
                    '.Stats.ExpSkills(QueryData!Number) = QueryData!Exp
                    '.Stats.EluSkills(QueryData!Number) = QueryData!ELU

392                 QueryData.MoveNext
                Wend

            End If

            'User friends
            'Call MakeQuery("SELECT * FROM friend WHERE user_id = ?;", False, .Id)

            'If Not QueryData Is Nothing Then
            '    QueryData.MoveFirst

            '    While Not QueryData.EOF

            '200     .Amigos(QueryData!Number).Nombre = QueryData!friend
            '200     .Amigos(QueryData!Number).Ignorado = QueryData!Ignored

            '        QueryData.MoveNext
            '    Wend
            'End If
        
            Dim LoopC As Byte
        
            'User quests
394         Call MakeQuery("SELECT * FROM quest WHERE user_id = ?;", False, .Id)

396         If Not QueryData Is Nothing Then
398             QueryData.MoveFirst

400             While Not QueryData.EOF

402                 .QuestStats.Quests(QueryData!Number).QuestIndex = QueryData!quest_id
                
404                 If .QuestStats.Quests(QueryData!Number).QuestIndex > 0 Then
406                     If QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredNPCs Then

                            Dim NPCs() As String

408                         NPCs = Split(QueryData!NPCs, "-")
410                         ReDim .QuestStats.Quests(QueryData!Number).NPCsKilled(1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredNPCs)

412                         For LoopC = 1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredNPCs
414                             .QuestStats.Quests(QueryData!Number).NPCsKilled(LoopC) = val(NPCs(LoopC - 1))
416                         Next LoopC

                        End If
                    
418                     If QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredTargetNPCs Then

                            Dim NPCsTarget() As String

420                         NPCsTarget = Split(QueryData!NPCsTarget, "-")
422                         ReDim .QuestStats.Quests(QueryData!Number).NPCsTarget(1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredTargetNPCs)

424                         For LoopC = 1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredTargetNPCs
426                             .QuestStats.Quests(QueryData!Number).NPCsTarget(LoopC) = val(NPCsTarget(LoopC - 1))
428                         Next LoopC

                        End If

                    End If

430                 QueryData.MoveNext
                Wend

            End If
        
            'User quests done
432         Call MakeQuery("SELECT * FROM quest_done WHERE user_id = ?;", False, .Id)

434         If Not QueryData Is Nothing Then
436             .QuestStats.NumQuestsDone = QueryData.RecordCount
                
438             ReDim .QuestStats.QuestsDone(1 To .QuestStats.NumQuestsDone)
        
440             QueryData.MoveFirst
            
442             LoopC = 1

444             While Not QueryData.EOF
            
446                 .QuestStats.QuestsDone(LoopC) = QueryData!quest_id
448                 LoopC = LoopC + 1

450                 QueryData.MoveNext
                Wend

            End If
        
            'User mail
            'TODO:
        
            ' Llaves
452         Call MakeQuery("SELECT key_obj FROM house_key WHERE account_id = ?", False, .AccountId)

454         If Not QueryData Is Nothing Then
456             QueryData.MoveFirst

458             LoopC = 1

460             While Not QueryData.EOF

462                 .Keys(LoopC) = QueryData!key_obj
464                 LoopC = LoopC + 1

466                 QueryData.MoveNext
                Wend

            End If

        End With

        Exit Sub

ErrorHandler:
468     Call LogDatabaseError("Error en LoadUserDatabase: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description & ". Línea: " & Erl)

470     Resume Next

End Sub

Public Function MakeQuery(query As String, ByVal NoResult As Boolean, ParamArray Query_Parameters() As Variant) As Boolean
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Hace una unica query a la db. Asume una conexion.
        ' Si NoResult = False, el metodo lee el resultado de la query
        ' Guarda el resultado en QueryData
    
        On Error GoTo ErrorHandler
    
        Dim Params As Variant

100     Set Command = New ADODB.Command
    
102     With Command

104         .ActiveConnection = Database_Connection
106         .CommandType = adCmdText
108         .NamedParameters = False
110         .CommandText = query
        
112         If UBound(Query_Parameters) < 0 Then
114             Params = Null
            
            Else
116             Params = Query_Parameters
            
118             If IsArray(Query_Parameters(0)) Then
120                 .Prepared = True
122                 Params = Query_Parameters(0)
                End If

            End If

124         If NoResult Then
126             Call .Execute(RecordsAffected, Params, adExecuteNoRecords)
    
            Else
128             Set QueryData = .Execute(RecordsAffected, Params)
    
130             If QueryData.BOF Or QueryData.EOF Then
132                 Set QueryData = Nothing
                End If
    
            End If
        
        End With
    
        Exit Function
    
ErrorHandler:

        Dim ErrNumber As Long, ErrDesc As String
134     ErrNumber = Err.Number
136     ErrDesc = Err.Description

138     If Not adoIsConnected(Database_Connection) Then
140         Call LogDatabaseError("Alerta en MakeQuery: Se perdió la conexión con la DB. Reconectando.")
142         Call Database_Connect
144         Resume
        
        Else
146         Call LogDatabaseError("Error en MakeQuery: query = '" & query & "'. " & ErrNumber & " - " & ErrDesc)
        
            On Error GoTo 0

148         Err.raise ErrNumber, "MakeQuery", ErrDesc

        End If

End Function

Private Function GetDBValue(Tabla As String, ColumnaGet As String, ColumnaTest As String, ValueTest As Variant) As Variant
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para leer un unico valor de una unica fila

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = ?;", False, ValueTest)
    
        'Revisamos si recibio un resultado
102     If QueryData Is Nothing Then Exit Function

        'Obtenemos la variable
104     GetDBValue = QueryData.Fields(0).Value

        Exit Function
    
ErrorHandler:
106     Call LogDatabaseError("Error en GetDBValue: SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = '" & ValueTest & "';" & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetCuentaValue(CuentaEmail As String, Columna As String) As Variant
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que leer un unico valor de la cuenta
        
        On Error GoTo GetCuentaValue_Err
        
100     GetCuentaValue = GetDBValue("account", Columna, "email", LCase$(CuentaEmail))

        
        Exit Function

GetCuentaValue_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetCuentaValue", Erl)
104     Resume Next
        
End Function

Public Function GetUserValue(CharName As String, Columna As String) As Variant
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que leer un unico valor del char
        
        On Error GoTo GetUserValue_Err
        
100     GetUserValue = GetDBValue("user", Columna, "name", CharName)

        
        Exit Function

GetUserValue_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetUserValue", Erl)
104     Resume Next
        
End Function

Public Sub SetDBValue(Tabla As String, ColumnaSet As String, ByVal ValueSet As Variant, ColumnaTest As String, ByVal ValueTest As Variant)
    ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para escribir un unico valor de una unica fila

        On Error GoTo ErrorHandler
    
        'Hacemos la query
100     Call MakeQuery("UPDATE " & Tabla & " SET " & ColumnaSet & " = ? WHERE " & ColumnaTest & " = ?;", True, ValueSet, ValueTest)

        Exit Sub
    
ErrorHandler:
102     Call LogDatabaseError("Error en SetDBValue: UPDATE " & Tabla & " SET " & ColumnaSet & " = " & ValueSet & " WHERE " & ColumnaTest & " = " & ValueTest & ";" & ". " & Err.Number & " - " & Err.Description)

End Sub

Private Sub SetCuentaValue(CuentaEmail As String, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor de la cuenta
        
        On Error GoTo SetCuentaValue_Err
        
100     Call SetDBValue("account", Columna, Value, "email", LCase$(CuentaEmail))

        
        Exit Sub

SetCuentaValue_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetCuentaValue", Erl)
104     Resume Next
        
End Sub

Private Sub SetUserValue(CharName As String, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor del char
        
        On Error GoTo SetUserValue_Err
        
100     Call SetDBValue("user", Columna, Value, "name", CharName)

        
        Exit Sub

SetUserValue_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetUserValue", Erl)
104     Resume Next
        
End Sub

Private Sub SetCuentaValueByID(ByVal AccountId As Long, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor de la cuenta
        ' Por ID
        
        On Error GoTo SetCuentaValueByID_Err
        
100     Call SetDBValue("account", Columna, Value, "id", AccountId)

        
        Exit Sub

SetCuentaValueByID_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetCuentaValueByID", Erl)
104     Resume Next
        
End Sub

Private Sub SetUserValueByID(ByVal Id As Long, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor del char
        ' Por ID
        
        On Error GoTo SetUserValueByID_Err
        
100     Call SetDBValue("user", Columna, Value, "id", Id)

        
        Exit Sub

SetUserValueByID_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetUserValueByID", Erl)
104     Resume Next
        
End Sub

Public Function CheckUserDonatorDatabase(CuentaEmail As String) As Boolean
        
        On Error GoTo CheckUserDonatorDatabase_Err
        
100     CheckUserDonatorDatabase = GetCuentaValue(CuentaEmail, "is_donor")

        
        Exit Function

CheckUserDonatorDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.CheckUserDonatorDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetUserCreditosDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetUserCreditosDatabase_Err
        
100     GetUserCreditosDatabase = GetCuentaValue(CuentaEmail, "credits")

        
        Exit Function

GetUserCreditosDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetUserCreditosDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetUserCreditosCanjeadosDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetUserCreditosCanjeadosDatabase_Err
        
100     GetUserCreditosCanjeadosDatabase = GetCuentaValue(CuentaEmail, "credits_used")

        
        Exit Function

GetUserCreditosCanjeadosDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetUserCreditosCanjeadosDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetUserDiasDonadorDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetUserDiasDonadorDatabase_Err
        

        Dim DonadorExpire As Variant

100     DonadorExpire = SanitizeNullValue(GetCuentaValue(CuentaEmail, "donor_expire"), False)
    
102     If Not DonadorExpire Then Exit Function
104     GetUserDiasDonadorDatabase = DateDiff("d", Date, DonadorExpire)

        
        Exit Function

GetUserDiasDonadorDatabase_Err:
106     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetUserDiasDonadorDatabase", Erl)
108     Resume Next
        
End Function

Public Function GetUserComprasDonadorDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetUserComprasDonadorDatabase_Err
        
100     GetUserComprasDonadorDatabase = GetCuentaValue(CuentaEmail, "donor_purchases")

        
        Exit Function

GetUserComprasDonadorDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetUserComprasDonadorDatabase", Erl)
104     Resume Next
        
End Function

Public Function CheckUserExists(name As String) As Boolean
        
        On Error GoTo CheckUserExists_Err
        
100     CheckUserExists = GetUserValue(name, "COUNT(*)") > 0

        
        Exit Function

CheckUserExists_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.CheckUserExists", Erl)
104     Resume Next
        
End Function

Public Function CheckCuentaExiste(CuentaEmail As String) As Boolean
        
        On Error GoTo CheckCuentaExiste_Err
        
100     CheckCuentaExiste = GetCuentaValue(CuentaEmail, "COUNT(*)") > 0

        
        Exit Function

CheckCuentaExiste_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.CheckCuentaExiste", Erl)
104     Resume Next
        
End Function

Public Function BANCheckDatabase(name As String) As Boolean
        
        On Error GoTo BANCheckDatabase_Err
        
100     BANCheckDatabase = CBool(GetUserValue(name, "is_banned"))

        
        Exit Function

BANCheckDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.BANCheckDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetCodigoActivacionDatabase(name As String) As String
        
        On Error GoTo GetCodigoActivacionDatabase_Err
        
100     GetCodigoActivacionDatabase = GetCuentaValue(name, "validate_code")

        
        Exit Function

GetCodigoActivacionDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetCodigoActivacionDatabase", Erl)
104     Resume Next
        
End Function

Public Function CheckCuentaActivadaDatabase(name As String) As Boolean
        
        On Error GoTo CheckCuentaActivadaDatabase_Err
        
100     CheckCuentaActivadaDatabase = GetCuentaValue(name, "validated")

        
        Exit Function

CheckCuentaActivadaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.CheckCuentaActivadaDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetEmailDatabase(name As String) As String
        
        On Error GoTo GetEmailDatabase_Err
        
100     GetEmailDatabase = GetCuentaValue(name, "email")

        
        Exit Function

GetEmailDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetEmailDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetMacAddressDatabase(CuentaEmail As String) As String
        
        On Error GoTo GetMacAddressDatabase_Err
        
100     GetMacAddressDatabase = GetCuentaValue(CuentaEmail, "mac_address")

        
        Exit Function

GetMacAddressDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetMacAddressDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetHDSerialDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetHDSerialDatabase_Err
        
100     GetHDSerialDatabase = GetCuentaValue(CuentaEmail, "hd_serial")

        
        Exit Function

GetHDSerialDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetHDSerialDatabase", Erl)
104     Resume Next
        
End Function

Public Function CheckBanCuentaDatabase(CuentaEmail As String) As Boolean
        
        On Error GoTo CheckBanCuentaDatabase_Err
        
100     CheckBanCuentaDatabase = CBool(GetCuentaValue(CuentaEmail, "is_banned"))

        
        Exit Function

CheckBanCuentaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.CheckBanCuentaDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetMotivoBanCuentaDatabase(CuentaEmail As String) As String
        
        On Error GoTo GetMotivoBanCuentaDatabase_Err
        
100     GetMotivoBanCuentaDatabase = GetCuentaValue(CuentaEmail, "ban_reason")

        
        Exit Function

GetMotivoBanCuentaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetMotivoBanCuentaDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetQuienBanCuentaDatabase(CuentaEmail As String) As String
        
        On Error GoTo GetQuienBanCuentaDatabase_Err
        
100     GetQuienBanCuentaDatabase = GetCuentaValue(CuentaEmail, "banned_by")

        
        Exit Function

GetQuienBanCuentaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetQuienBanCuentaDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetCuentaLogeadaDatabase(CuentaEmail As String) As Boolean
        
        On Error GoTo GetCuentaLogeadaDatabase_Err
        
100     GetCuentaLogeadaDatabase = GetCuentaValue(CuentaEmail, "is_logged")

        
        Exit Function

GetCuentaLogeadaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetCuentaLogeadaDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetUserStatusDatabase(name As String) As Integer
        
        On Error GoTo GetUserStatusDatabase_Err
        
100     GetUserStatusDatabase = GetUserValue(name, "status")

        
        Exit Function

GetUserStatusDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetUserStatusDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetAccountIDDatabase(name As String) As Long
        
        On Error GoTo GetAccountIDDatabase_Err
        

        Dim temp As Variant

100     temp = GetUserValue(name, "account_id")
    
102     If VBA.IsEmpty(temp) Then
104         GetAccountIDDatabase = -1
        Else
106         GetAccountIDDatabase = val(temp)

        End If

        
        Exit Function

GetAccountIDDatabase_Err:
108     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetAccountIDDatabase", Erl)
110     Resume Next
        
End Function

Public Sub GetPasswordAndSaltDatabase(CuentaEmail As String, PasswordHash As String, Salt As String)

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("SELECT password, salt FROM account WHERE deleted = FALSE AND email = ?;", False, LCase$(CuentaEmail))

102     If QueryData Is Nothing Then Exit Sub
    
104     PasswordHash = QueryData!Password
106     Salt = QueryData!Salt
    
        Exit Sub
    
ErrorHandler:
108     Call LogDatabaseError("Error in GetPasswordAndSaltDatabase. Cuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)
    
End Sub

Public Function GetPersonajesCountDatabase(CuentaEmail As String) As Byte

        On Error GoTo ErrorHandler

        Dim Id As Long

100     Id = GetDBValue("account", "id", "email", LCase$(CuentaEmail))
    
102     GetPersonajesCountDatabase = GetPersonajesCountByIDDatabase(Id)
    
        Exit Function
    
ErrorHandler:
104     Call LogDatabaseError("Error in GetPersonajesCountDatabase. name: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)
    
End Function

Public Function GetPersonajesCountByIDDatabase(ByVal AccountId As Long) As Byte

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("SELECT COUNT(*) FROM user WHERE deleted = FALSE AND account_id = ?;", False, AccountId)
    
102     If QueryData Is Nothing Then Exit Function
    
104     GetPersonajesCountByIDDatabase = QueryData.Fields(0).Value
    
        Exit Function
    
ErrorHandler:
106     Call LogDatabaseError("Error in GetPersonajesCountByIDDatabase. AccountID: " & AccountId & ". " & Err.Number & " - " & Err.Description)
    
End Function

Public Function GetPersonajesCuentaDatabase(ByVal AccountId As Long, Personaje() As PersonajeCuenta) As Byte
        
        On Error GoTo GetPersonajesCuentaDatabase_Err
        

100     Call MakeQuery("SELECT name, head_id, class_id, body_id, pos_map, pos_x, pos_y, level, status, helmet_id, shield_id, weapon_id, guild_index, is_dead, is_sailing FROM user WHERE deleted = FALSE AND account_id = ?;", False, AccountId)

102     If QueryData Is Nothing Then Exit Function
    
104     GetPersonajesCuentaDatabase = QueryData.RecordCount
        
106     QueryData.MoveFirst
    
        Dim i As Integer

108     For i = 1 To GetPersonajesCuentaDatabase
110         Personaje(i).nombre = QueryData!name
112         Personaje(i).Cabeza = QueryData!head_id
114         Personaje(i).clase = QueryData!class_id
116         Personaje(i).cuerpo = QueryData!body_id
118         Personaje(i).Mapa = QueryData!pos_map
120         Personaje(i).posX = QueryData!pos_x
122         Personaje(i).posY = QueryData!pos_y
124         Personaje(i).nivel = QueryData!level
126         Personaje(i).Status = QueryData!Status
128         Personaje(i).Casco = QueryData!helmet_id
130         Personaje(i).Escudo = QueryData!shield_id
132         Personaje(i).Arma = QueryData!weapon_id
134         Personaje(i).ClanIndex = QueryData!Guild_Index
        
136         If EsRolesMaster(Personaje(i).nombre) Then
138             Personaje(i).Status = 3
140         ElseIf EsConsejero(Personaje(i).nombre) Then
142             Personaje(i).Status = 4
144         ElseIf EsSemiDios(Personaje(i).nombre) Then
146             Personaje(i).Status = 5
148         ElseIf EsDios(Personaje(i).nombre) Then
150             Personaje(i).Status = 6
152         ElseIf EsAdmin(Personaje(i).nombre) Then
154             Personaje(i).Status = 7

            End If

156         If val(QueryData!is_dead) = 1 Or val(QueryData!is_sailing) = 1 Then
158             Personaje(i).Cabeza = 0
            End If
        
160         QueryData.MoveNext
        Next

        
        Exit Function

GetPersonajesCuentaDatabase_Err:
162     Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetPersonajesCuentaDatabase", Erl)
164     Resume Next
        
End Function

Public Sub SetUserLoggedDatabase(ByVal Id As Long, ByVal AccountId As Long)
        
        On Error GoTo SetUserLoggedDatabase_Err
        
100     Call SetDBValue("user", "is_logged", 1, "id", Id)
102     Call MakeQuery("UPDATE account SET logged = logged + 1 WHERE id = ?;", True, AccountId)

        
        Exit Sub

SetUserLoggedDatabase_Err:
104     Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetUserLoggedDatabase", Erl)
106     Resume Next
        
End Sub

Public Sub ResetLoggedDatabase(ByVal AccountId As Long)
        
        On Error GoTo ResetLoggedDatabase_Err
        
100     Call MakeQuery("UPDATE account SET logged = 0 WHERE id = ?;", True, AccountId)

        
        Exit Sub

ResetLoggedDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.ResetLoggedDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub SetUsersLoggedDatabase(ByVal NumUsers As Long)
        
        On Error GoTo SetUsersLoggedDatabase_Err
        
100     Call MakeQuery("UPDATE statistics SET value = ? WHERE name = 'online';", True, NumUsers)
        
        Exit Sub

SetUsersLoggedDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetUsersLoggedDatabase", Erl)
104     Resume Next
        
End Sub

Public Function LeerRecordUsuariosDatabase() As Long
        
        On Error GoTo LeerRecordUsuariosDatabase_Err
        
100     Call MakeQuery("SELECT value FROM statistics WHERE name = 'record';", False)

102     If QueryData Is Nothing Then Exit Function

104     LeerRecordUsuariosDatabase = val(QueryData!Value)

        Exit Function

LeerRecordUsuariosDatabase_Err:
106     Call RegistrarError(Err.Number, Err.Description, "modDatabase.LeerRecordUsuariosDatabase", Erl)
108     Resume Next
        
End Function

Public Sub SetRecordUsersDatabase(ByVal Record As Long)
        
        On Error GoTo SetRecordUsersDatabase_Err
        
100     Call MakeQuery("UPDATE statistics SET value = ? WHERE name = 'record';", True, CStr(Record))
        
        Exit Sub

SetRecordUsersDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetRecordUsersDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub LogoutAllUsersAndAccounts()
        
        On Error GoTo LogoutAllUsersAndAccounts_Err

100     Call MakeQuery("UPDATE user SET is_logged = FALSE; UPDATE account SET logged = 0;", True)
        
        Exit Sub

LogoutAllUsersAndAccounts_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.LogoutAllUsersAndAccounts", Erl)
104     Resume Next
        
End Sub

Public Sub SaveVotoDatabase(ByVal Id As Long, ByVal Encuestas As Integer)
        
        On Error GoTo SaveVotoDatabase_Err
        
100     Call SetUserValueByID(Id, "votes_amount", Encuestas)

        
        Exit Sub

SaveVotoDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.SaveVotoDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub SaveUserBodyDatabase(UserName As String, ByVal Body As Integer)
        
        On Error GoTo SaveUserBodyDatabase_Err
        
100     Call SetUserValue(UserName, "body_id", Body)

        
        Exit Sub

SaveUserBodyDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.SaveUserBodyDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub SaveUserHeadDatabase(UserName As String, ByVal Head As Integer)
        
        On Error GoTo SaveUserHeadDatabase_Err
        
100     Call SetUserValue(UserName, "head_id", Head)

        
        Exit Sub

SaveUserHeadDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.SaveUserHeadDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub SaveUserSkillDatabase(UserName As String, ByVal Skill As Integer, ByVal Value As Integer)
        
        On Error GoTo SaveUserSkillDatabase_Err
        
100     Call MakeQuery("UPDATE skillpoints SET value = ? WHERE number = ? AND user_id = (SELECT id FROM user WHERE UPPER(name) = ?);", True, Value, Skill, UCase$(UserName))
        
        Exit Sub

SaveUserSkillDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.SaveUserSkillDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub SaveUserSkillsLibres(UserName As String, ByVal SkillsLibres As Integer)
        
        On Error GoTo SaveUserHeadDatabase_Err
        
100     Call SetUserValue(UserName, "free_skillpoints", SkillsLibres)
        
        Exit Sub

SaveUserHeadDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.SaveUserHeadDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub SaveNewAccountDatabase(CuentaEmail As String, PasswordHash As String, Salt As String, Codigo As String)

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("INSERT INTO account SET email = ?, password = ?, salt = ?, validate_code = ?, date_created = NOW();", True, LCase$(CuentaEmail), PasswordHash, Salt, Codigo)
    
        Exit Sub
        
ErrorHandler:
102     Call LogDatabaseError("Error en SaveNewAccountDatabase. Cuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub ValidarCuentaDatabase(UserCuenta As String)
        
        On Error GoTo ValidarCuentaDatabase_Err
        
100     Call SetCuentaValue(UserCuenta, "validated", 1)
        
        Exit Sub

ValidarCuentaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.ValidarCuentaDatabase", Erl)
104     Resume Next
        
End Sub

Public Function CheckUserAccount(name As String, ByVal AccountId As Long) As Boolean

100     CheckUserAccount = (val(GetUserValue(name, "account_id")) = AccountId)

End Function

Public Sub BorrarUsuarioDatabase(name As String)

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("UPDATE user SET name = CONCAT('DELETED_', name), deleted = TRUE WHERE UPPER(name) = ?;", True, UCase$(name))

        Exit Sub
    
ErrorHandler:
102     Call LogDatabaseError("Error en BorrarUsuarioDatabase borrando user de la Mysql Database: " & name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub BorrarCuentaDatabase(CuentaEmail As String)

        On Error GoTo ErrorHandler

        Dim Id As Long

100     Id = GetDBValue("account", "id", "email", LCase$(CuentaEmail))

102     Call MakeQuery("UPDATE account SET email = CONCAT('DELETED_', email), deleted = TRUE WHERE email = ?;", True, LCase$(CuentaEmail))

104     Call MakeQuery("UPDATE user SET name = CONCAT('DELETED_', name), deleted = TRUE WHERE account_id = ?;", True, Id)

        Exit Sub
    
ErrorHandler:
106     Call LogDatabaseError("Error en BorrarCuentaDatabase borrando user de la Mysql Database: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveBanDatabase(UserName As String, Reason As String, BannedBy As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Call MakeQuery("UPDATE user SET is_banned = TRUE, banned_by = ?, ban_reason = ? WHERE UPPER(name) = ?;", True, BannedBy, Reason, UCase$(UserName))

        Call SavePenaDatabase(UserName, "Baneado por: " & BannedBy & " debido a " & Reason)

        Exit Sub

ErrorHandler:
112     Call LogDatabaseError("Error in SaveBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveWarnDatabase(UserName As String, Reason As String, WarnedBy As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

    Call MakeQuery("UPDATE user SET warnings = warnings + 1 WHERE UPPER(name) = ?;", True, UCase$(UserName))
    
    Call SavePenaDatabase(UserName, "Advertencia de: " & WarnedBy & " debido a " & Reason)
    
    Exit Sub

ErrorHandler:
112     Call LogDatabaseError("Error in SaveWarnDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SavePenaDatabase(UserName As String, Reason As String)

        On Error GoTo ErrorHandler

        Dim query As String
        query = "INSERT INTO punishment(user_id, NUMBER, reason)"
        query = query & " SELECT u.id, COUNT(p.number) + 1, ? FROM user u LEFT JOIN punishment p ON p.user_id = u.id WHERE UPPER(u.name) = ?"

        Call MakeQuery(query, True, Reason, UCase$(UserName))

        Exit Sub

ErrorHandler:
110     Call LogDatabaseError("Error in SavePenaDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SilenciarUserDatabase(UserName As String, ByVal Tiempo As Integer)
    
        On Error GoTo ErrorHandler

100     Call MakeQuery("UPDATE user SET is_silenced = 1, silence_minutes_left = ?, silence_elapsed_seconds = 0 WHERE UPPER(name) = ?;", True, Tiempo, UCase$(UserName))

        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in SilenciarUserDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)
    
End Sub

Public Sub DesilenciarUserDatabase(UserName As String)

        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "is_silenced", 0)

        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in DesilenciarUserDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)
    
End Sub

Public Sub UnBanDatabase(UserName As String)

        On Error GoTo ErrorHandler

100     Call MakeQuery("UPDATE user SET is_banned = FALSE WHERE UPPER(name) = ?;", True, UCase$(UserName))

        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in UnBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveBanCuentaDatabase(ByVal AccountId As Long, Reason As String, BannedBy As String)

        On Error GoTo ErrorHandler

100     Call MakeQuery("UPDATE account SET is_banned = TRUE, banned_by = ?, ban_reason = ? WHERE id = ?;", True, BannedBy, Reason, AccountId)

        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in SaveBanCuentaDatabase: AccountId=" & AccountId & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub EcharConsejoDatabase(UserName As String)
        
        On Error GoTo EcharConsejoDatabase_Err
        
100     Call MakeQuery("UPDATE user SET pertenece_consejo_real = FALSE, pertenece_consejo_caos = FALSE WHERE UPPER(name) = ?;", True, UCase$(UserName))

        
        Exit Sub

EcharConsejoDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.EcharConsejoDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub EcharLegionDatabase(UserName As String)
        
        On Error GoTo EcharLegionDatabase_Err
        
100     Call MakeQuery("UPDATE user SET pertenece_caos = FALSE, reenlistadas = 200 WHERE UPPER(name) = ?;", True, UCase$(UserName))

        
        Exit Sub

EcharLegionDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.EcharLegionDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub EcharArmadaDatabase(UserName As String)
        
        On Error GoTo EcharArmadaDatabase_Err
        
100     Call MakeQuery("UPDATE user SET pertenece_real = FALSE, reenlistadas = 200 WHERE UPPER(name) = ?;", True, UCase$(UserName))

        
        Exit Sub

EcharArmadaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.EcharArmadaDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub CambiarPenaDatabase(UserName As String, ByVal Numero As Integer, Pena As String)
        
        On Error GoTo CambiarPenaDatabase_Err
        
100     Call MakeQuery("UPDATE punishment SET reason = ? WHERE number = ? AND user_id = (SELECT id from user WHERE UPPER(name) = ?);", True, Pena, Numero, UCase$(UserName))

        
        Exit Sub

CambiarPenaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.CambiarPenaDatabase", Erl)
104     Resume Next
        
End Sub

Public Function GetUserAmountOfPunishmentsDatabase(ByVal UserName As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call MakeQuery("SELECT COUNT(*) as punishments FROM `punishment` INNER JOIN `user` ON punishment.user_id = user.id WHERE UPPER(user.name) = ?;", False, UCase$(UserName))

102     If QueryData Is Nothing Then Exit Function

104     GetUserAmountOfPunishmentsDatabase = QueryData!punishments

        Exit Function
ErrorHandler:
106     Call LogDatabaseError("Error in GetUserAmountOfPunishmentsDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Sub SendUserPunishmentsDatabase(ByVal UserIndex As Integer, ByVal UserName As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call MakeQuery("SELECT user_id, number, reason FROM `punishment` INNER JOIN `user` ON punishment.user_id = user.id WHERE UPPER(user.name) = ?;", False, UCase$(UserName))
    
102     If QueryData Is Nothing Then Exit Sub

104     If Not QueryData.RecordCount = 0 Then
106         QueryData.MoveFirst

108         While Not QueryData.EOF
110             Call WriteConsoleMsg(UserIndex, QueryData!Number & " - " & QueryData!Reason, FontTypeNames.FONTTYPE_INFO)
            
112             QueryData.MoveNext
            Wend

        End If

        Exit Sub
ErrorHandler:
114     Call LogDatabaseError("Error in SendUserPunishmentsDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub


Public Function GetNombreCuentaDatabase(name As String) As String

        On Error GoTo ErrorHandler

        'Hacemos la query.
100     Call MakeQuery("SELECT email FROM `account` INNER JOIN `user` ON user.account_id = account.id WHERE UPPER(user.name) = ?;", False, UCase$(name))
    
        'Verificamos que la query no devuelva un resultado vacio.
102     If QueryData Is Nothing Then Exit Function
    
        'Obtenemos el nombre de la cuenta
104     GetNombreCuentaDatabase = QueryData!email

        Exit Function
    
ErrorHandler:
106     Call LogDatabaseError("Error en GetNombreCuentaDatabase leyendo user de la Mysql Database: " & name & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildIndexDatabase(UserName As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 09/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildIndexDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_index"), 0)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildMemberDatabase(UserName As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildMemberDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_member_history"), vbNullString)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildAspirantDatabase(UserName As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildAspirantDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_aspirant_index"), 0)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildRejectionReasonDatabase(UserName As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildRejectionReasonDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_rejected_because"), vbNullString)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildPedidosDatabase(UserName As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildPedidosDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_requests_history"), vbNullString)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Sub SaveUserGuildRejectionReasonDatabase(UserName As String, Reason As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_rejected_because", Reason)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildIndexDatabase(ByVal UserName As String, ByVal GuildIndex As Integer)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_index", GuildIndex)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildAspirantDatabase(ByVal UserName As String, ByVal AspirantIndex As Integer)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_aspirant_index", AspirantIndex)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildMemberDatabase(ByVal UserName As String, ByVal guilds As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_member_history", guilds)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildPedidosDatabase(ByVal UserName As String, ByVal Pedidos As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(UserName, "guild_requests_history", Pedidos)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SendCharacterInfoDatabase(ByVal UserIndex As Integer, ByVal UserName As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Dim gName       As String

        Dim Miembro     As String

        Dim GuildActual As Integer

100     Call MakeQuery("SELECT race_id, class_id, genre_id, level, gold, bank_gold, guild_requests_history, guild_index, guild_member_history, pertenece_real, pertenece_caos, ciudadanos_matados, criminales_matados FROM user WHERE UPPER(name) = ?;", False, UCase$(UserName))

102     If QueryData Is Nothing Then
104         Call WriteConsoleMsg(UserIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        ' Get the character's current guild
106     GuildActual = SanitizeNullValue(QueryData!Guild_Index, 0)

108     If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
110         gName = "<" & GuildName(GuildActual) & ">"
        Else
112         gName = "Ninguno"

        End If

        'Get previous guilds
114     Miembro = SanitizeNullValue(QueryData!guild_member_history, vbNullString)

116     If Len(Miembro) > 400 Then
118         Miembro = ".." & Right$(Miembro, 400)

        End If

120     Call WriteCharacterInfo(UserIndex, UserName, QueryData!race_id, QueryData!class_id, QueryData!genre_id, QueryData!level, QueryData!gold, QueryData!bank_gold, SanitizeNullValue(QueryData!guild_requests_history, vbNullString), gName, Miembro, QueryData!pertenece_real, QueryData!pertenece_caos, QueryData!ciudadanos_matados, QueryData!criminales_matados)

        Exit Sub
ErrorHandler:
122     Call LogDatabaseError("Error in SendCharacterInfoDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function EnterAccountDatabase(ByVal UserIndex As Integer, CuentaEmail As String, Password As String, MacAddress As String, ByVal HDserial As Long, ip As String) As Boolean

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("SELECT id, password, salt, validated, is_banned, ban_reason, banned_by FROM account WHERE email = ?;", False, LCase$(CuentaEmail))
    
102     If Database_Connection.State = adStateClosed Then
104         Call WriteShowMessageBox(UserIndex, "Ha ocurrido un error interno en el servidor. ¡Estamos tratando de resolverlo!")
            Exit Function
        End If
    
106     If QueryData Is Nothing Then
108         Call WriteShowMessageBox(UserIndex, "La cuenta no existe.")
            Exit Function
        End If
    
110     If val(QueryData!is_banned) > 0 Then
112         Call WriteShowMessageBox(UserIndex, "La cuenta se encuentra baneada debido a: " & QueryData!ban_reason & ". Esta decisión fue tomada por: " & QueryData!banned_by & ".")
            Exit Function
        End If
    
114     If Not PasswordValida(Password, QueryData!Password, QueryData!Salt) Then
116         Call WriteShowMessageBox(UserIndex, "Contraseña inválida.")
            Exit Function
        End If
    
118     If val(QueryData!validated) = 0 Then
120         Call WriteShowMessageBox(UserIndex, "¡La cuenta no ha sido validada aún!")
            Exit Function
        End If
    
122     UserList(UserIndex).AccountId = QueryData!Id
124     UserList(UserIndex).Cuenta = CuentaEmail
    
126     Call MakeQuery("UPDATE account SET mac_address = ?, hd_serial = ?, last_ip = ?, last_access = NOW() WHERE id = ?;", True, MacAddress, HDserial, ip, QueryData!Id)
    
128     EnterAccountDatabase = True
    
        Exit Function

ErrorHandler:
130     Call LogDatabaseError("Error in EnterAccountDatabase. UserCuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function PersonajePerteneceEmail(ByVal UserName As String, ByVal AccountEmail As String) As Boolean
    
100     Call MakeQuery("SELECT id FROM user INNER JOIN account ON user.account_id = account.id WHERE user.name = ? AND account.email = ?;", False, UserName, AccountEmail)
    
102     If QueryData Is Nothing Then
104         PersonajePerteneceEmail = False
            Exit Function
        End If
    
106     PersonajePerteneceEmail = True
    
End Function

Public Function PersonajePerteneceID(ByVal UserName As String, ByVal AccountId As Long) As Boolean
    
100     Call MakeQuery("SELECT id FROM user WHERE name = ? AND account_id = ?;", False, UserName, AccountId)
    
102     If QueryData Is Nothing Then
104         PersonajePerteneceID = False
            Exit Function
        End If
    
106     PersonajePerteneceID = True
    
End Function

Public Sub ChangePasswordDatabase(ByVal UserIndex As Integer, OldPassword As String, NewPassword As String)

        On Error GoTo ErrorHandler

100     If LenB(NewPassword) = 0 Then
102         Call WriteConsoleMsg(UserIndex, "Debe especificar una contraseña nueva, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
104     Call MakeQuery("SELECT password, salt FROM account WHERE id = ?;", False, UserList(UserIndex).AccountId)
    
106     If QueryData Is Nothing Then
108         Call WriteConsoleMsg(UserIndex, "No se ha podido cambiar la contraseña por un error interno. Avise a un administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
110     If Not PasswordValida(OldPassword, QueryData!Password, QueryData!Salt) Then
112         Call WriteConsoleMsg(UserIndex, "La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        Dim Salt As String * 10
114         Salt = RandomString(10) ' Alfanumerico
    
        Dim oSHA256 As CSHA256
116     Set oSHA256 = New CSHA256

        Dim PasswordHash As String * 64
118         PasswordHash = oSHA256.SHA256(NewPassword & Salt)
    
120     Set oSHA256 = Nothing
    
122     Call MakeQuery("UPDATE account SET password = ?, salt = ? WHERE id = ?;", True, PasswordHash, Salt, UserList(UserIndex).AccountId)
    
124     Call WriteConsoleMsg(UserIndex, "La contraseña de su cuenta fue cambiada con éxito.", FontTypeNames.FONTTYPE_INFO)
    
        Exit Sub

ErrorHandler:
126     Call LogDatabaseError("Error in ChangePasswordDatabase. Username: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function GetUsersLoggedAccountDatabase(ByVal AccountId As Long) As Byte

        On Error GoTo ErrorHandler

100     Call GetDBValue("account", "logged", "id", AccountId)
    
102     If QueryData Is Nothing Then Exit Function
    
104     GetUsersLoggedAccountDatabase = val(QueryData!logged)

        Exit Function

ErrorHandler:
106     Call LogDatabaseError("Error in GetUsersLoggedAccountDatabase. AccountID: " & AccountId & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function SetPositionDatabase(UserName As String, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        On Error GoTo ErrorHandler

100     Call MakeQuery("UPDATE user SET pos_map = ?, pos_x = ?, pos_y = ? WHERE UPPER(name) = ?;", True, Map, X, Y, UCase$(UserName))
    
102     SetPositionDatabase = RecordsAffected > 0

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in SetPositionDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetMapDatabase(UserName As String) As Integer
        On Error GoTo ErrorHandler

100     GetMapDatabase = val(GetUserValue(UserName, "pos_map"))

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in SetPositionDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function AddOroBancoDatabase(UserName As String, ByVal OroGanado As Long) As Boolean
        On Error GoTo ErrorHandler

100     Call MakeQuery("UPDATE user SET bank_gold = bank_gold + ? WHERE UPPER(name) = ?;", True, OroGanado, UCase$(UserName))
    
102     AddOroBancoDatabase = RecordsAffected > 0

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in AddOroBancoDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function DarLlaveAUsuarioDatabase(UserName As String, ByVal LlaveObj As Integer) As Boolean
        On Error GoTo ErrorHandler

100     Call MakeQuery("INSERT INTO house_key SET key_obj = ?, account_id = (SELECT account_id FROM user WHERE UPPER(name) = ?);", True, LlaveObj, UCase$(UserName))
    
102     DarLlaveAUsuarioDatabase = RecordsAffected > 0

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in DarLlaveAUsuarioDatabase. UserName: " & UserName & ", LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function DarLlaveACuentaDatabase(email As String, ByVal LlaveObj As Integer) As Boolean
        On Error GoTo ErrorHandler

100     Call MakeQuery("INSERT INTO house_key SET key_obj = ?, account_id = (SELECT id FROM account WHERE UPPER(email) = ?);", True, LlaveObj, UCase$(email))
    
102     DarLlaveACuentaDatabase = RecordsAffected > 0
        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in DarLlaveACuentaDatabase. Email: " & email & ", LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function SacarLlaveDatabase(ByVal LlaveObj As Integer) As Boolean
        On Error GoTo ErrorHandler

        Dim i As Integer
        Dim UserCount As Integer
        Dim Users() As String

        ' Obtengo los usuarios logueados en la cuenta del dueño de la llave
100     Call MakeQuery("SELECT name FROM `user` INNER JOIN `account` ON `user`.account_id = account.id INNER JOIN `house_key` ON `house_key`.account_id = account.id WHERE `user`.is_logged = TRUE AND `house_key`.key_obj = ?;", False, LlaveObj)
    
102     If QueryData Is Nothing Then Exit Function

        ' Los almaceno en un array
104     UserCount = QueryData.RecordCount
    
106     ReDim Users(1 To UserCount) As String
    
108     QueryData.MoveFirst

110     i = 1

112     While Not QueryData.EOF
    
114         Users(i) = QueryData!name
116         i = i + 1

118         QueryData.MoveNext
        Wend
    
        ' Intento borrar la llave de la db
120     Call MakeQuery("DELETE FROM house_key WHERE key_obj = ?;", True, LlaveObj)
    
        ' Si pudimos borrar, actualizamos los usuarios logueados
        Dim UserIndex As Integer
    
122     For i = 1 To UserCount
124         UserIndex = NameIndex(Users(i))
        
126         If UserIndex <> 0 Then
128             Call SacarLlaveDeLLavero(UserIndex, LlaveObj)
            End If
        Next
    
130     SacarLlaveDatabase = RecordsAffected > 0

        Exit Function

ErrorHandler:
132     Call LogDatabaseError("Error in SacarLlaveDatabase. LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.Description)

End Function

Public Sub VerLlavesDatabase(ByVal UserIndex As Integer)
        On Error GoTo ErrorHandler

100     Call MakeQuery("SELECT email, key_obj FROM `house_key` INNER JOIN `account` ON `house_key`.account_id = `account`.id;", False)

102     If QueryData Is Nothing Then
104         Call WriteConsoleMsg(UserIndex, "No hay llaves otorgadas por el momento.", FontTypeNames.FONTTYPE_INFO)

106     ElseIf QueryData.RecordCount = 0 Then
108         Call WriteConsoleMsg(UserIndex, "No hay llaves otorgadas por el momento.", FontTypeNames.FONTTYPE_INFO)
    
        Else
            Dim message As String
        
110         message = "Llaves usadas: " & QueryData.RecordCount & vbNewLine
    
112         QueryData.MoveFirst

114         While Not QueryData.EOF
116             message = message & "Llave: " & QueryData!key_obj & " - Cuenta: " & QueryData!email & vbNewLine

118             QueryData.MoveNext
            Wend
        
120         message = Left$(message, Len(message) - 2)
        
122         Call WriteConsoleMsg(UserIndex, message, FontTypeNames.FONTTYPE_INFO)
        End If

        Exit Sub

ErrorHandler:
124     Call LogDatabaseError("Error in VerLlavesDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function SanitizeNullValue(ByVal Value As Variant, ByVal defaultValue As Variant) As Variant
        
        On Error GoTo SanitizeNullValue_Err
        
100     SanitizeNullValue = IIf(IsNull(Value), defaultValue, Value)

        
        Exit Function

SanitizeNullValue_Err:
102     Call RegistrarError(Err.Number, Err.Description, "modDatabase.SanitizeNullValue", Erl)
104     Resume Next
        
End Function

Function adoIsConnected(adoCn As ADODB.Connection) As Boolean

        '----------------------------------------------------------------
        '#PURPOSE: Checks whether the supplied db connection is alive and
        '          hasn't had it's TCP connection forcibly closed by remote
        '          host, for example, as happens during an undock event
        '#RETURNS: True if the supplied db is connected and error-free,
        '          False otherwise
        '#AUTHOR:  Belladonna
        '----------------------------------------------------------------

        ' No sacar
        On Error Resume Next

        Dim i As Long
        Dim cmd As New ADODB.Command

        'Set up SQL command to return 1
100     cmd.CommandText = "SELECT 1"
102     cmd.ActiveConnection = adoCn

        'Run a simple query, to test the connection
        
104     i = cmd.Execute.Fields(0)
        On Error GoTo 0

        'Tidy up
106     Set cmd = Nothing

        'If i is 1, connection is open
108     If i = 1 Then
110         adoIsConnected = True
        Else
112         adoIsConnected = False
        End If

End Function
