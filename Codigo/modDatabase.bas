Attribute VB_Name = "modDatabase"
'Argentum Online Libre
'Database connection module
'Obtained from GS-Zone
'Adapted and modified by Juan Andres Dalmasso (CHOTS)
'September 2018
'Rewrited for Argentum20 by Alexis Caraballo (WyroX)
'October 2020

Option Explicit

Public Database_Enabled    As Boolean
Public Database_DataSource As String
Public Database_Host       As String
Public Database_Name       As String
Public Database_Username   As String
Public Database_Password   As String

Public Database_Connection As ADODB.Connection
Public QueryData           As ADODB.Recordset
Public RecordsAffected     As Long

Private QueryBuilder       As cStringBuilder
Private ConnectedOnce      As Boolean

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
116     Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.description)
    
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
     
100     Call Database_Connection.Close
    
102     Set Database_Connection = Nothing
     
        Exit Sub
     
ErrorHandler:
104     Call LogDatabaseError("Unable to close Mysql Database: " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveNewUserDatabase(ByVal UserIndex As Integer)

        On Error GoTo ErrorHandler
    
        'Constructor de queries.
        'Me permite concatenar strings MUCHO MAS rapido
100     Set QueryBuilder = New cStringBuilder
    
102     With UserList(UserIndex)
    
            'Basic user data
104         QueryBuilder.Append "INSERT INTO user SET "
106         QueryBuilder.Append "name = '" & .name & "', "
108         QueryBuilder.Append "account_id = " & .AccountID & ", "
110         QueryBuilder.Append "level = " & .Stats.ELV & ", "
112         QueryBuilder.Append "exp = " & .Stats.Exp & ", "
114         QueryBuilder.Append "elu = " & .Stats.ELU & ", "
116         QueryBuilder.Append "genre_id = " & .genero & ", "
118         QueryBuilder.Append "race_id = " & .raza & ", "
120         QueryBuilder.Append "class_id = " & .clase & ", "
122         QueryBuilder.Append "home_id = " & .Hogar & ", "
124         QueryBuilder.Append "description = '" & .Desc & "', "
126         QueryBuilder.Append "gold = " & .Stats.GLD & ", "
128         QueryBuilder.Append "free_skillpoints = " & .Stats.SkillPts & ", "
            'QueryBuilder.Append "assigned_skillpoints = " & .Counters.AsignedSkills & ", "
130         QueryBuilder.Append "pos_map = " & .Pos.Map & ", "
132         QueryBuilder.Append "pos_x = " & .Pos.X & ", "
134         QueryBuilder.Append "pos_y = " & .Pos.Y & ", "
136         QueryBuilder.Append "body_id = " & .Char.Body & ", "
138         QueryBuilder.Append "head_id = " & .Char.Head & ", "
140         QueryBuilder.Append "weapon_id = " & .Char.WeaponAnim & ", "
142         QueryBuilder.Append "helmet_id = " & .Char.CascoAnim & ", "
144         QueryBuilder.Append "shield_id = " & .Char.ShieldAnim & ", "
146         QueryBuilder.Append "items_Amount = " & .Invent.NroItems & ", "
148         QueryBuilder.Append "slot_armour = " & .Invent.ArmourEqpSlot & ", "
150         QueryBuilder.Append "slot_weapon = " & .Invent.WeaponEqpSlot & ", "
152         QueryBuilder.Append "slot_shield = " & .Invent.EscudoEqpSlot & ", "
154         QueryBuilder.Append "slot_helmet = " & .Invent.CascoEqpSlot & ", "
156         QueryBuilder.Append "slot_ammo = " & .Invent.MunicionEqpSlot & ", "
158         QueryBuilder.Append "slot_dm = " & .Invent.DañoMagicoEqpSlot & ", "
160         QueryBuilder.Append "slot_rm = " & .Invent.ResistenciaEqpSlot & ", "
162         QueryBuilder.Append "slot_tool = " & .Invent.HerramientaEqpSlot & ", "
164         QueryBuilder.Append "slot_magic = " & .Invent.MagicoSlot & ", "
166         QueryBuilder.Append "slot_knuckles = " & .Invent.NudilloSlot & ", "
168         QueryBuilder.Append "slot_ship = " & .Invent.BarcoSlot & ", "
170         QueryBuilder.Append "slot_mount = " & .Invent.MonturaSlot & ", "
172         QueryBuilder.Append "min_hp = " & .Stats.MinHp & ", "
174         QueryBuilder.Append "max_hp = " & .Stats.MaxHp & ", "
176         QueryBuilder.Append "min_man = " & .Stats.MinMAN & ", "
178         QueryBuilder.Append "max_man = " & .Stats.MaxMAN & ", "
180         QueryBuilder.Append "min_sta = " & .Stats.MinSta & ", "
182         QueryBuilder.Append "max_sta = " & .Stats.MaxSta & ", "
184         QueryBuilder.Append "min_ham = " & .Stats.MinHam & ", "
186         QueryBuilder.Append "max_ham = " & .Stats.MaxHam & ", "
188         QueryBuilder.Append "min_sed = " & .Stats.MinAGU & ", "
190         QueryBuilder.Append "max_sed = " & .Stats.MaxAGU & ", "
192         QueryBuilder.Append "min_hit = " & .Stats.MinHIT & ", "
194         QueryBuilder.Append "max_hit = " & .Stats.MaxHit & ", "
            'QueryBuilder.Append "rep_noble = " & .NobleRep & ", "
            'QueryBuilder.Append "rep_plebe = " & .Reputacion.PlebeRep & ", "
            'QueryBuilder.Append "rep_average = " & .Reputacion.Promedio & ", "
196         QueryBuilder.Append "is_naked = " & .flags.Desnudo & ", "
198         QueryBuilder.Append "status = " & .Faccion.Status & ", "
200         QueryBuilder.Append "is_logged = TRUE; "
        
202         Call MakeQuery(QueryBuilder.toString, True)
        
            'Borramos la query construida.
204         Call QueryBuilder.Clear
        
            ' Para recibir el ID del user
206         Call MakeQuery("SELECT LAST_INSERT_ID();")

208         If QueryData Is Nothing Then
210             .Id = 1
            Else
212             .Id = val(QueryData.Fields(0).Value)
            End If
        
            ' Comenzamos una cadena de queries (para enviar todo de una)
            Dim LoopC As Long

            'User attributes
214         QueryBuilder.Append "INSERT INTO attribute (user_id, number, value) VALUES "

216         For LoopC = 1 To NUMATRIBUTOS
        
218             QueryBuilder.Append "("
220             QueryBuilder.Append .Id & ", "
222             QueryBuilder.Append LoopC & ", "
224             QueryBuilder.Append .Stats.UserAtributos(LoopC) & ")"

226             If LoopC < NUMATRIBUTOS Then
228                 QueryBuilder.Append ", "
                Else
230                 QueryBuilder.Append "; "
                End If

232         Next LoopC

            'User spells
234         QueryBuilder.Append "INSERT INTO spell (user_id, number, spell_id) VALUES "

236         For LoopC = 1 To MAXUSERHECHIZOS
238             QueryBuilder.Append "("
240             QueryBuilder.Append .Id & ", "
242             QueryBuilder.Append LoopC & ", "
244             QueryBuilder.Append .Stats.UserHechizos(LoopC) & ")"

246             If LoopC < MAXUSERHECHIZOS Then
248                 QueryBuilder.Append ", "
                Else
250                 QueryBuilder.Append "; "
                End If

252         Next LoopC

            'User inventory
254         QueryBuilder.Append "INSERT INTO inventory_item (user_id, number, item_id, Amount, is_equipped) VALUES "

256         For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
258             QueryBuilder.Append "("
260             QueryBuilder.Append .Id & ", "
262             QueryBuilder.Append LoopC & ", "
264             QueryBuilder.Append .Invent.Object(LoopC).ObjIndex & ", "
266             QueryBuilder.Append .Invent.Object(LoopC).Amount & ", "
268             QueryBuilder.Append .Invent.Object(LoopC).Equipped & ")"

270             If LoopC < UserList(UserIndex).CurrentInventorySlots Then
272                 QueryBuilder.Append ", "
                Else
274                 QueryBuilder.Append "; "
                End If

276         Next LoopC

            'User skills
            'QueryBuilder.Append "INSERT INTO skillpoint (user_id, number, value, exp, elu) VALUES "
278         QueryBuilder.Append "INSERT INTO skillpoint (user_id, number, value) VALUES "

280         For LoopC = 1 To NUMSKILLS
282             QueryBuilder.Append "("
284             QueryBuilder.Append .Id & ", "
286             QueryBuilder.Append LoopC & ", "
288             QueryBuilder.Append .Stats.UserSkills(LoopC) & ")"
                'QueryBuilder.Append .Stats.UserSkills(LoopC) & ", "
                'QueryBuilder.Append .Stats.ExpSkills(LoopC) & ", "
                'QueryBuilder.Append .Stats.EluSkills(LoopC) & ")"

290             If LoopC < NUMSKILLS Then
292                 QueryBuilder.Append ", "
                Else
294                 QueryBuilder.Append "; "
                End If

296         Next LoopC
        
            'User friends
            'QueryBuilder.Append "INSERT INTO friend (user_id, number) VALUES "

            'For LoopC = 1 To MAXAMIGOS
        
            '    QueryBuilder.Append "("
            '    QueryBuilder.Append .ID & ", "
            '    QueryBuilder.Append LoopC & ")"

            '    If LoopC < MAXAMIGOS Then
            '        QueryBuilder.Append ", "
            '    Else
            '        QueryBuilder.Append "; "

            '    End If
        
            'Next LoopC
        
            'User quests
298         QueryBuilder.Append "INSERT INTO quest (user_id, number) VALUES "

300         For LoopC = 1 To MAXUSERQUESTS
        
302             QueryBuilder.Append "("
304             QueryBuilder.Append .Id & ", "
306             QueryBuilder.Append LoopC & ")"

308             If LoopC < MAXUSERQUESTS Then
310                 QueryBuilder.Append ", "
                Else
312                 QueryBuilder.Append "; "
                End If

314         Next LoopC
        
            'User pets
316         QueryBuilder.Append "INSERT INTO pet (user_id, number, pet_id) VALUES "

318         For LoopC = 1 To MAXMASCOTAS
        
320             QueryBuilder.Append "("
322             QueryBuilder.Append .Id & ", "
324             QueryBuilder.Append LoopC & ", 0)"

326             If LoopC < MAXMASCOTAS Then
328                 QueryBuilder.Append ", "
                Else
330                 QueryBuilder.Append "; "
                End If

332         Next LoopC

            'Enviamos todas las queries
334         Call MakeQuery(QueryBuilder.toString, True)
        
336         Set QueryBuilder = Nothing
    
        End With

        Exit Sub

ErrorHandler:
    
338     Set QueryBuilder = Nothing
    
340     Call LogDatabaseError("Error en SaveNewUserDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserDatabase(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean = False)

        On Error GoTo ErrorHandler
    
        'Constructor de queries.
        'Me permite concatenar strings MUCHO MAS rapido
100     Set QueryBuilder = New cStringBuilder

        'Basic user data
102     With UserList(UserIndex)
104         QueryBuilder.Append "UPDATE user SET "
106         QueryBuilder.Append "name = '" & .name & "', "
108         QueryBuilder.Append "level = " & .Stats.ELV & ", "
110         QueryBuilder.Append "exp = " & CLng(.Stats.Exp) & ", "
112         QueryBuilder.Append "elu = " & .Stats.ELU & ", "
114         QueryBuilder.Append "genre_id = " & .genero & ", "
116         QueryBuilder.Append "race_id = " & .raza & ", "
118         QueryBuilder.Append "class_id = " & .clase & ", "
120         QueryBuilder.Append "home_id = " & .Hogar & ", "
122         QueryBuilder.Append "description = '" & .Desc & "', "
124         QueryBuilder.Append "gold = " & .Stats.GLD & ", "
126         QueryBuilder.Append "bank_gold = " & .Stats.Banco & ", "
128         QueryBuilder.Append "free_skillpoints = " & .Stats.SkillPts & ", "
            'QueryBuilder.Append "assigned_skillpoints = " & .Counters.AsignedSkills & ", "
130         QueryBuilder.Append "pets_saved = " & .flags.MascotasGuardadas & ", "
132         QueryBuilder.Append "pos_map = " & .Pos.Map & ", "
134         QueryBuilder.Append "pos_x = " & .Pos.X & ", "
136         QueryBuilder.Append "pos_y = " & .Pos.Y & ", "
138         QueryBuilder.Append "last_map = " & .flags.lastMap & ", "
140         QueryBuilder.Append "message_info = '" & .MENSAJEINFORMACION & "', "
142         QueryBuilder.Append "body_id = " & .Char.Body & ", "
144         QueryBuilder.Append "head_id = " & .OrigChar.Head & ", "
146         QueryBuilder.Append "weapon_id = " & .Char.WeaponAnim & ", "
148         QueryBuilder.Append "helmet_id = " & .Char.CascoAnim & ", "
150         QueryBuilder.Append "shield_id = " & .Char.ShieldAnim & ", "
152         QueryBuilder.Append "heading = " & .Char.Heading & ", "
154         QueryBuilder.Append "items_Amount = " & .Invent.NroItems & ", "
156         QueryBuilder.Append "slot_armour = " & .Invent.ArmourEqpSlot & ", "
158         QueryBuilder.Append "slot_weapon = " & .Invent.WeaponEqpSlot & ", "
160         QueryBuilder.Append "slot_shield = " & .Invent.EscudoEqpSlot & ", "
162         QueryBuilder.Append "slot_helmet = " & .Invent.CascoEqpSlot & ", "
164         QueryBuilder.Append "slot_ammo = " & .Invent.MunicionEqpSlot & ", "
166         QueryBuilder.Append "slot_dm = " & .Invent.DañoMagicoEqpSlot & ", "
168         QueryBuilder.Append "slot_rm = " & .Invent.ResistenciaEqpSlot & ", "
170         QueryBuilder.Append "slot_tool = " & .Invent.HerramientaEqpSlot & ", "
172         QueryBuilder.Append "slot_magic = " & .Invent.MagicoSlot & ", "
174         QueryBuilder.Append "slot_knuckles = " & .Invent.NudilloSlot & ", "
176         QueryBuilder.Append "slot_ship = " & .Invent.BarcoSlot & ", "
178         QueryBuilder.Append "slot_mount = " & .Invent.MonturaSlot & ", "
180         QueryBuilder.Append "min_hp = " & .Stats.MinHp & ", "
182         QueryBuilder.Append "max_hp = " & .Stats.MaxHp & ", "
184         QueryBuilder.Append "min_man = " & .Stats.MinMAN & ", "
186         QueryBuilder.Append "max_man = " & .Stats.MaxMAN & ", "
188         QueryBuilder.Append "min_sta = " & .Stats.MinSta & ", "
190         QueryBuilder.Append "max_sta = " & .Stats.MaxSta & ", "
192         QueryBuilder.Append "min_ham = " & .Stats.MinHam & ", "
194         QueryBuilder.Append "max_ham = " & .Stats.MaxHam & ", "
196         QueryBuilder.Append "min_sed = " & .Stats.MinAGU & ", "
198         QueryBuilder.Append "max_sed = " & .Stats.MaxAGU & ", "
200         QueryBuilder.Append "min_hit = " & .Stats.MinHIT & ", "
202         QueryBuilder.Append "max_hit = " & .Stats.MaxHit & ", "
204         QueryBuilder.Append "killed_npcs = " & .Stats.NPCsMuertos & ", "
206         QueryBuilder.Append "killed_users = " & .Stats.UsuariosMatados & ", "
208         QueryBuilder.Append "invent_level = " & .Stats.InventLevel & ", "
            'QueryBuilder.Append "rep_asesino = " & .Reputacion.AsesinoRep & ", "
            'QueryBuilder.Append "rep_bandido = " & .Reputacion.BandidoRep & ", "
            'QueryBuilder.Append "rep_burgues = " & .Reputacion.BurguesRep & ", "
            'QueryBuilder.Append "rep_ladron = " & .Reputacion.LadronesRep & ", "
            'QueryBuilder.Append "rep_noble = " & .Reputacion.NobleRep & ", "
            'QueryBuilder.Append "rep_plebe = " & .Reputacion.PlebeRep & ", "
            'QueryBuilder.Append "rep_average = " & .Reputacion.Promedio & ", "
210         QueryBuilder.Append "is_naked = " & .flags.Desnudo & ", "
212         QueryBuilder.Append "is_poisoned = " & .flags.Envenenado & ", "
214         QueryBuilder.Append "is_hidden = " & .flags.Escondido & ", "
216         QueryBuilder.Append "is_hungry = " & .flags.Hambre & ", "
218         QueryBuilder.Append "is_thirsty = " & .flags.Sed & ", "
            'QueryBuilder.Append "is_banned = " & .flags.Ban & ", " Esto es innecesario porque se setea cuando lo baneas (creo)
220         QueryBuilder.Append "is_dead = " & .flags.Muerto & ", "
222         QueryBuilder.Append "is_sailing = " & .flags.Navegando & ", "
224         QueryBuilder.Append "is_paralyzed = " & .flags.Paralizado & ", "
226         QueryBuilder.Append "is_mounted = " & .flags.Montado & ", "
228         QueryBuilder.Append "is_silenced = " & .flags.Silenciado & ", "
230         QueryBuilder.Append "silence_minutes_left = " & .flags.MinutosRestantes & ", "
232         QueryBuilder.Append "silence_elapsed_seconds = " & .flags.SegundosPasados & ", "
234         QueryBuilder.Append "spouse = '" & .flags.Pareja & "', "
236         QueryBuilder.Append "counter_pena = " & .Counters.Pena & ", "
238         QueryBuilder.Append "deaths = " & .flags.VecesQueMoriste & ", "
240         QueryBuilder.Append "pertenece_consejo_real = " & (.flags.Privilegios And PlayerType.RoyalCouncil) & ", "
242         QueryBuilder.Append "pertenece_consejo_caos = " & (.flags.Privilegios And PlayerType.ChaosCouncil) & ", "
244         QueryBuilder.Append "pertenece_real = " & .Faccion.ArmadaReal & ", "
246         QueryBuilder.Append "pertenece_caos = " & .Faccion.FuerzasCaos & ", "
248         QueryBuilder.Append "ciudadanos_matados = " & .Faccion.CiudadanosMatados & ", "
250         QueryBuilder.Append "criminales_matados = " & .Faccion.CriminalesMatados & ", "
252         QueryBuilder.Append "recibio_armadura_real = " & .Faccion.RecibioArmaduraReal & ", "
254         QueryBuilder.Append "recibio_armadura_caos = " & .Faccion.RecibioArmaduraCaos & ", "
256         QueryBuilder.Append "recibio_exp_real = " & .Faccion.RecibioExpInicialReal & ", "
258         QueryBuilder.Append "recibio_exp_caos = " & .Faccion.RecibioExpInicialCaos & ", "
260         QueryBuilder.Append "recompensas_real = " & .Faccion.RecompensasReal & ", "
262         QueryBuilder.Append "recompensas_caos = " & .Faccion.RecompensasCaos & ", "
264         QueryBuilder.Append "reenlistadas = " & .Faccion.Reenlistadas & ", "
266         QueryBuilder.Append "fecha_ingreso = " & IIf(.Faccion.FechaIngreso <> vbNullString, "'" & .Faccion.FechaIngreso & "'", "NULL") & ", "
268         QueryBuilder.Append "nivel_ingreso = " & .Faccion.NivelIngreso & ", "
270         QueryBuilder.Append "matados_ingreso = " & .Faccion.MatadosIngreso & ", "
272         QueryBuilder.Append "siguiente_recompensa = " & .Faccion.NextRecompensa & ", "
274         QueryBuilder.Append "status = " & .Faccion.Status & ", "
276         QueryBuilder.Append "battle_points = " & .flags.BattlePuntos & ", "
278         QueryBuilder.Append "guild_index = " & .GuildIndex & ", "
280         QueryBuilder.Append "chat_combate = " & .ChatCombate & ", "
282         QueryBuilder.Append "chat_global = " & .ChatGlobal & ", "
284         QueryBuilder.Append "is_logged = " & IIf(Logout, "FALSE", "TRUE") & ", "
286         QueryBuilder.Append "warnings = " & .Stats.Advertencias
288         QueryBuilder.Append " WHERE id = " & .Id & "; "
        
            Dim LoopC As Long

            'User attributes
290         QueryBuilder.Append "INSERT INTO attribute (user_id, number, value) VALUES "

292         For LoopC = 1 To NUMATRIBUTOS
        
294             QueryBuilder.Append "("
296             QueryBuilder.Append .Id & ", "
298             QueryBuilder.Append LoopC & ", "
300             QueryBuilder.Append .Stats.UserAtributosBackUP(LoopC) & ")"

302             If LoopC < NUMATRIBUTOS Then
304                 QueryBuilder.Append ", "
                End If

306         Next LoopC
        
308         QueryBuilder.Append " ON DUPLICATE KEY UPDATE value=VALUES(value); "

            'User spells
310         QueryBuilder.Append "INSERT INTO spell (user_id, number, spell_id) VALUES "

312         For LoopC = 1 To MAXUSERHECHIZOS
        
314             QueryBuilder.Append "("
316             QueryBuilder.Append .Id & ", "
318             QueryBuilder.Append LoopC & ", "
320             QueryBuilder.Append .Stats.UserHechizos(LoopC) & ")"

322             If LoopC < MAXUSERHECHIZOS Then
324                 QueryBuilder.Append ", "
                End If

326         Next LoopC
        
328         QueryBuilder.Append " ON DUPLICATE KEY UPDATE spell_id=VALUES(spell_id); "

            'User inventory
330         QueryBuilder.Append "INSERT INTO inventory_item (user_id, number, item_id, Amount, is_equipped) VALUES "

332         For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
        
334             QueryBuilder.Append "("
336             QueryBuilder.Append .Id & ", "
338             QueryBuilder.Append LoopC & ", "
340             QueryBuilder.Append .Invent.Object(LoopC).ObjIndex & ", "
342             QueryBuilder.Append .Invent.Object(LoopC).Amount & ", "
344             QueryBuilder.Append .Invent.Object(LoopC).Equipped & ")"

346             If LoopC < UserList(UserIndex).CurrentInventorySlots Then
348                 QueryBuilder.Append ", "
                End If

350         Next LoopC
        
352         QueryBuilder.Append " ON DUPLICATE KEY UPDATE item_id=VALUES(item_id), amount=VALUES(Amount), is_equipped=VALUES(is_equipped); "

            'User bank inventory
354         QueryBuilder.Append "INSERT INTO bank_item (user_id, number, item_id, Amount) VALUES "

356         For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
        
358             QueryBuilder.Append "("
360             QueryBuilder.Append .Id & ", "
362             QueryBuilder.Append LoopC & ", "
364             QueryBuilder.Append .BancoInvent.Object(LoopC).ObjIndex & ", "
366             QueryBuilder.Append .BancoInvent.Object(LoopC).Amount & ")"

368             If LoopC < MAX_BANCOINVENTORY_SLOTS Then
370                 QueryBuilder.Append ", "
                End If

372         Next LoopC
        
374         QueryBuilder.Append " ON DUPLICATE KEY UPDATE item_id=VALUES(item_id), amount=VALUES(Amount); "

            'User skills
            'QueryBuilder.Append "INSERT INTO skillpoint (user_id, number, value, exp, elu) VALUES "
376         QueryBuilder.Append "INSERT INTO skillpoint (user_id, number, value) VALUES "

378         For LoopC = 1 To NUMSKILLS
        
380             QueryBuilder.Append "("
382             QueryBuilder.Append .Id & ", "
384             QueryBuilder.Append LoopC & ", "
386             QueryBuilder.Append .Stats.UserSkills(LoopC) & ")"
                'QueryBuilder.Append .Stats.UserSkills(LoopC) & ", "
                'Q  = Q & .Stats.ExpSkills(LoopC) & ", "
                'QueryBuilder.Append .Stats.EluSkills(LoopC) & ")"

388             If LoopC < NUMSKILLS Then
390                 QueryBuilder.Append ", "
                End If

392         Next LoopC
        
            'QueryBuilder.Append " ON DUPLICATE KEY UPDATE value=VALUES(value), exp=VALUES(exp), elu=VALUES(elu); "
394         QueryBuilder.Append " ON DUPLICATE KEY UPDATE value=VALUES(value); "

            'User pets
            Dim petType As Integer
        
396         QueryBuilder.Append "INSERT INTO pet (user_id, number, pet_id) VALUES "

398         For LoopC = 1 To MAXMASCOTAS
400             QueryBuilder.Append "("
402             QueryBuilder.Append .Id & ", "
404             QueryBuilder.Append LoopC & ", "

                'CHOTS | I got this logic from SaveUserToCharfile
406             If .MascotasIndex(LoopC) > 0 Then
            
408                 If Npclist(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
410                     petType = .MascotasType(LoopC)
                    Else
412                     petType = 0
                    End If

                Else
414                 petType = .MascotasType(LoopC)

                End If

416             QueryBuilder.Append petType & ")"

418             If LoopC < MAXMASCOTAS Then
420                 QueryBuilder.Append ", "
                End If

422         Next LoopC

424         QueryBuilder.Append " ON DUPLICATE KEY UPDATE pet_id=VALUES(pet_id); "
        
            'User friends
            'Q = "INSERT INTO friend (user_id, number, friend, ignored) VALUES "

            'For LoopC = 1 To MAXAMIGOS
        
            '    QueryBuilder.Append "("
            '    QueryBuilder.Append .ID & ", "
            '    QueryBuilder.Append LoopC & ", "
            '    QueryBuilder.Append "'" & .Amigos(LoopC).Nombre & "', "
            '    QueryBuilder.Append .Amigos(LoopC).Ignorado & ")"

            '    If LoopC < MAXAMIGOS Then
            '        QueryBuilder.Append ", "
            '    Else
            '        QueryBuilder.Append ";"

            '    End If
        
            'Next LoopC
       
            'Agrego ip del user
426         QueryBuilder.Append "INSERT INTO connection (user_id, ip, date_last_login) VALUES ("
428         QueryBuilder.Append .Id & ", "
430         QueryBuilder.Append "'" & .ip & "', "
432         QueryBuilder.Append "NOW()) "
434         QueryBuilder.Append "ON DUPLICATE KEY UPDATE "
436         QueryBuilder.Append "date_last_login = VALUES(date_last_login); "
        
            'Borro la mas vieja si hay mas de 5 (WyroX: si alguien sabe una forma mejor de hacerlo me avisa)
438         QueryBuilder.Append "DELETE FROM connection WHERE"
440         QueryBuilder.Append " user_id = " & .Id
442         QueryBuilder.Append " AND date_last_login < (SELECT min(date_last_login) FROM (SELECT date_last_login FROM connection WHERE"
444         QueryBuilder.Append " user_id = " & .Id
446         QueryBuilder.Append " ORDER BY date_last_login DESC LIMIT 5) AS d); "
        
            'User quests
448         QueryBuilder.Append "INSERT INTO quest (user_id, number, quest_id, npcs, npcstarget) VALUES "
        
            Dim Tmp As Integer, LoopK As Long

450         For LoopC = 1 To MAXUSERQUESTS
452             QueryBuilder.Append "("
454             QueryBuilder.Append .Id & ", "
456             QueryBuilder.Append LoopC & ", "
458             QueryBuilder.Append .QuestStats.Quests(LoopC).QuestIndex & ", '"
            
460             If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
462                 Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredNPCs

464                 If Tmp Then

466                     For LoopK = 1 To Tmp
468                         QueryBuilder.Append CStr(.QuestStats.Quests(LoopC).NPCsKilled(LoopK))
                        
470                         If LoopK < Tmp Then
472                             QueryBuilder.Append "-"
                            End If

474                     Next LoopK
                    

                    End If

                End If
            
476             QueryBuilder.Append "', '"
            
478             If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
                
480                 Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredTargetNPCs
                    
482                 For LoopK = 1 To Tmp

484                     QueryBuilder.Append CStr(.QuestStats.Quests(LoopC).NPCsTarget(LoopK))
                    
486                     If LoopK < Tmp Then
488                         QueryBuilder.Append "-"
                        End If
                
490                 Next LoopK
            
                End If
            
492             QueryBuilder.Append "')"

494             If LoopC < MAXUSERQUESTS Then
496                 QueryBuilder.Append ", "
                End If

498         Next LoopC
        
500         QueryBuilder.Append " ON DUPLICATE KEY UPDATE quest_id=VALUES(quest_id), npcs=VALUES(npcs); "
        
            'User completed quests
502         If .QuestStats.NumQuestsDone > 0 Then
504             QueryBuilder.Append "INSERT INTO quest_done (user_id, quest_id) VALUES "
    
506             For LoopC = 1 To .QuestStats.NumQuestsDone
508                 QueryBuilder.Append "("
510                 QueryBuilder.Append .Id & ", "
512                 QueryBuilder.Append .QuestStats.QuestsDone(LoopC) & ")"
    
514                 If LoopC < .QuestStats.NumQuestsDone Then
516                     QueryBuilder.Append ", "
                    End If
    
518             Next LoopC
            
520             QueryBuilder.Append " ON DUPLICATE KEY UPDATE quest_id=VALUES(quest_id); "

            End If
        
            'User mail
            'TODO:
        
            ' Si deslogueó, actualizo la cuenta
522         If Logout Then
524             QueryBuilder.Append "UPDATE account SET logged = logged - 1 WHERE id = " & .AccountID & ";"
            End If
526         Debug.Print
528         Call MakeQuery(QueryBuilder.toString, True)

        End With
    
530     Set QueryBuilder = Nothing
    
        Exit Sub

ErrorHandler:

532     Set QueryBuilder = Nothing
    
534     Call LogDatabaseError("Error en SaveUserDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.description)

End Sub

Sub LoadUserDatabase(ByVal UserIndex As Integer)

    On Error GoTo ErrorHandler

    'Basic user data
100 With UserList(UserIndex)

102     Call MakeQuery("SELECT *, DATE_FORMAT(fecha_ingreso, '%Y-%m-%d') as 'fecha_ingreso_format' FROM user WHERE name ='" & .name & "';")

104     If QueryData Is Nothing Then Exit Sub

        'Start setting data
106     .Id = QueryData!Id
108     .name = QueryData!name
110     .Stats.ELV = QueryData!level
112     .Stats.Exp = QueryData!Exp
114     .Stats.ELU = QueryData!ELU
116     .genero = QueryData!genre_id
118     .raza = QueryData!race_id
120     .clase = QueryData!class_id
122     .Hogar = QueryData!home_id
124     .Desc = QueryData!description
126     .Stats.GLD = QueryData!gold
128     .Stats.Banco = QueryData!bank_gold
130     .Stats.SkillPts = QueryData!free_skillpoints
        '.Counters.AsignedSkills = QueryData!assigned_skillpoints
132     .Pos.Map = QueryData!pos_map
134     .Pos.X = QueryData!pos_x
136     .Pos.Y = QueryData!pos_y
138     .flags.lastMap = QueryData!last_map
140     .MENSAJEINFORMACION = QueryData!message_info
142     .OrigChar.Body = QueryData!body_id
144     .OrigChar.Head = QueryData!head_id
146     .OrigChar.WeaponAnim = QueryData!weapon_id
148     .OrigChar.CascoAnim = QueryData!helmet_id
150     .OrigChar.ShieldAnim = QueryData!shield_id
152     .OrigChar.Heading = QueryData!Heading
154     .Invent.NroItems = QueryData!items_Amount
156     .Invent.ArmourEqpSlot = SanitizeNullValue(QueryData!slot_armour, 0)
158     .Invent.WeaponEqpSlot = SanitizeNullValue(QueryData!slot_weapon, 0)
160     .Invent.CascoEqpSlot = SanitizeNullValue(QueryData!slot_helmet, 0)
162     .Invent.EscudoEqpSlot = SanitizeNullValue(QueryData!slot_shield, 0)
164     .Invent.MunicionEqpSlot = SanitizeNullValue(QueryData!slot_ammo, 0)
166     .Invent.BarcoSlot = SanitizeNullValue(QueryData!slot_ship, 0)
168     .Invent.MonturaSlot = SanitizeNullValue(QueryData!slot_mount, 0)
170     .Invent.DañoMagicoEqpSlot = SanitizeNullValue(QueryData!slot_dm, 0)
172     .Invent.ResistenciaEqpSlot = SanitizeNullValue(QueryData!slot_rm, 0)
174     .Invent.NudilloSlot = SanitizeNullValue(QueryData!slot_knuckles, 0)
176     .Invent.HerramientaEqpSlot = SanitizeNullValue(QueryData!slot_tool, 0)
178     .Invent.MagicoSlot = SanitizeNullValue(QueryData!slot_magic, 0)
180     .Stats.MinHp = QueryData!min_hp
182     .Stats.MaxHp = QueryData!max_hp
184     .Stats.MinMAN = QueryData!min_man
186     .Stats.MaxMAN = QueryData!max_man
188     .Stats.MinSta = QueryData!min_sta
190     .Stats.MaxSta = QueryData!max_sta
192     .Stats.MinHam = QueryData!min_ham
194     .Stats.MaxHam = QueryData!max_ham
196     .Stats.MinAGU = QueryData!min_sed
198     .Stats.MaxAGU = QueryData!max_sed
200     .Stats.MinHIT = QueryData!min_hit
202     .Stats.MaxHit = QueryData!max_hit
204     .Stats.NPCsMuertos = QueryData!killed_npcs
206     .Stats.UsuariosMatados = QueryData!killed_users
208     .Stats.InventLevel = QueryData!invent_level
        '.Reputacion.AsesinoRep = QueryData!rep_asesino
        '.Reputacion.BandidoRep = QueryData!rep_bandido
        '.Reputacion.BurguesRep = QueryData!rep_burgues
        '.Reputacion.LadronesRep = QueryData!rep_ladron
        '.Reputacion.NobleRep = QueryData!rep_noble
        '.Reputacion.PlebeRep = QueryData!rep_plebe
        '.Reputacion.Promedio = QueryData!rep_average
210     .flags.Desnudo = QueryData!is_naked
212     .flags.Envenenado = QueryData!is_poisoned
214     .flags.Escondido = QueryData!is_hidden
216     .flags.Hambre = QueryData!is_hungry
218     .flags.Sed = QueryData!is_thirsty
220     .flags.Ban = QueryData!is_banned
222     .flags.Muerto = QueryData!is_dead
224     .flags.Navegando = QueryData!is_sailing
226     .flags.Paralizado = QueryData!is_paralyzed
228     .flags.VecesQueMoriste = QueryData!deaths
230     .flags.BattlePuntos = QueryData!battle_points
232     .flags.Montado = QueryData!is_mounted
234     .flags.Pareja = QueryData!spouse
236     .flags.Casado = IIf(Len(.flags.Pareja) > 0, 1, 0)
238     .flags.Silenciado = QueryData!is_silenced
240     .flags.MinutosRestantes = QueryData!silence_minutes_left
242     .flags.SegundosPasados = QueryData!silence_elapsed_seconds
244     .flags.MascotasGuardadas = QueryData!pets_saved
246     .flags.ScrollExp = 1 'TODO: sacar
248     .flags.ScrollOro = 1 'TODO: sacar
        
250     .Counters.Pena = QueryData!counter_pena
        
252     .ChatGlobal = QueryData!chat_global
254     .ChatCombate = QueryData!chat_combate

256     If QueryData!pertenece_consejo_real Then
258         .flags.Privilegios = .flags.Privilegios Or PlayerType.RoyalCouncil

        End If

260     If QueryData!pertenece_consejo_caos Then
262         .flags.Privilegios = .flags.Privilegios Or PlayerType.ChaosCouncil

        End If

264     .Faccion.ArmadaReal = QueryData!pertenece_real
266     .Faccion.FuerzasCaos = QueryData!pertenece_caos
268     .Faccion.CiudadanosMatados = QueryData!ciudadanos_matados
270     .Faccion.CriminalesMatados = QueryData!criminales_matados
272     .Faccion.RecibioArmaduraReal = QueryData!recibio_armadura_real
274     .Faccion.RecibioArmaduraCaos = QueryData!recibio_armadura_caos
276     .Faccion.RecibioExpInicialReal = QueryData!recibio_exp_real
278     .Faccion.RecibioExpInicialCaos = QueryData!recibio_exp_caos
280     .Faccion.RecompensasReal = QueryData!recompensas_real
282     .Faccion.RecompensasCaos = QueryData!recompensas_caos
284     .Faccion.Reenlistadas = QueryData!Reenlistadas
286     .Faccion.FechaIngreso = SanitizeNullValue(QueryData!fecha_ingreso_format, vbNullString)
288     .Faccion.NivelIngreso = SanitizeNullValue(QueryData!nivel_ingreso, 0)
290     .Faccion.MatadosIngreso = SanitizeNullValue(QueryData!matados_ingreso, 0)
292     .Faccion.NextRecompensa = SanitizeNullValue(QueryData!siguiente_recompensa, 0)
294     .Faccion.Status = QueryData!Status

296     .GuildIndex = SanitizeNullValue(QueryData!Guild_Index, 0)
        
298     .Stats.Advertencias = QueryData!warnings
        
        'User attributes
300     Call MakeQuery("SELECT * FROM attribute WHERE user_id = " & .Id & ";")
    
302     If Not QueryData Is Nothing Then
304         QueryData.MoveFirst

306         While Not QueryData.EOF

308             .Stats.UserAtributos(QueryData!Number) = QueryData!Value
310             .Stats.UserAtributosBackUP(QueryData!Number) = .Stats.UserAtributos(QueryData!Number)

312             QueryData.MoveNext
            Wend

        End If

        'User spells
314     Call MakeQuery("SELECT * FROM spell WHERE user_id = " & .Id & ";")

316     If Not QueryData Is Nothing Then
318         QueryData.MoveFirst

320         While Not QueryData.EOF

322             .Stats.UserHechizos(QueryData!Number) = QueryData!spell_id

324             QueryData.MoveNext
            Wend

        End If

        'User pets
326     Call MakeQuery("SELECT * FROM pet WHERE user_id = " & .Id & ";")

328     If Not QueryData Is Nothing Then
330         QueryData.MoveFirst

332         While Not QueryData.EOF

334             .MascotasType(QueryData!Number) = QueryData!pet_id
                
336             If val(QueryData!pet_id) <> 0 Then
338                 .NroMascotas = .NroMascotas + 1
                End If

340             QueryData.MoveNext
            Wend
        End If

        'User inventory
342     Call MakeQuery("SELECT * FROM inventory_item WHERE user_id = " & .Id & ";")

344     If Not QueryData Is Nothing Then
346         QueryData.MoveFirst

348         While Not QueryData.EOF

350             .Invent.Object(QueryData!Number).ObjIndex = QueryData!item_id
352             .Invent.Object(QueryData!Number).Amount = QueryData!Amount
354             .Invent.Object(QueryData!Number).Equipped = QueryData!is_equipped

356             QueryData.MoveNext
            Wend

        End If

        'User bank inventory
358     Call MakeQuery("SELECT * FROM bank_item WHERE user_id = " & .Id & ";")

360     If Not QueryData Is Nothing Then
362         QueryData.MoveFirst

364         While Not QueryData.EOF

366             .BancoInvent.Object(QueryData!Number).ObjIndex = QueryData!item_id
368             .BancoInvent.Object(QueryData!Number).Amount = QueryData!Amount

370             QueryData.MoveNext
            Wend

        End If

        'User skills
372     Call MakeQuery("SELECT * FROM skillpoint WHERE user_id = " & .Id & ";")

374     If Not QueryData Is Nothing Then
376         QueryData.MoveFirst

378         While Not QueryData.EOF

380             .Stats.UserSkills(QueryData!Number) = QueryData!Value
                '.Stats.ExpSkills(QueryData!Number) = QueryData!Exp
                '.Stats.EluSkills(QueryData!Number) = QueryData!ELU

382             QueryData.MoveNext
            Wend

        End If

        'User friends
        'Call MakeQuery("SELECT * FROM friend WHERE user_id = " & .ID & ";")

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
384     Call MakeQuery("SELECT * FROM quest WHERE user_id = " & .Id & ";")

386     If Not QueryData Is Nothing Then
388         QueryData.MoveFirst

390         While Not QueryData.EOF

392             .QuestStats.Quests(QueryData!Number).QuestIndex = QueryData!quest_id
                
394             If .QuestStats.Quests(QueryData!Number).QuestIndex > 0 Then
396                 If QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredNPCs Then

                        Dim NPCs() As String

398                     NPCs = Split(QueryData!NPCs, "-")
400                     ReDim .QuestStats.Quests(QueryData!Number).NPCsKilled(1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredNPCs)

402                     For LoopC = 1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredNPCs
404                         .QuestStats.Quests(QueryData!Number).NPCsKilled(LoopC) = val(NPCs(LoopC - 1))
406                     Next LoopC

                    End If
                    
                    
408                 If QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredTargetNPCs Then

                        Dim NPCsTarget() As String

410                     NPCsTarget = Split(QueryData!NPCsTarget, "-")
412                     ReDim .QuestStats.Quests(QueryData!Number).NPCsTarget(1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredTargetNPCs)

414                     For LoopC = 1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredTargetNPCs
416                         .QuestStats.Quests(QueryData!Number).NPCsTarget(LoopC) = val(NPCsTarget(LoopC - 1))
418                     Next LoopC

                    End If

                End If

420             QueryData.MoveNext
            Wend

        End If
        
        'User quests done
422     Call MakeQuery("SELECT * FROM quest_done WHERE user_id = " & .Id & ";")

424     If Not QueryData Is Nothing Then
426         .QuestStats.NumQuestsDone = QueryData.RecordCount
                
428         ReDim .QuestStats.QuestsDone(1 To .QuestStats.NumQuestsDone)
        
430         QueryData.MoveFirst
            
432         LoopC = 1

434         While Not QueryData.EOF
            
436             .QuestStats.QuestsDone(LoopC) = QueryData!quest_id
438             LoopC = LoopC + 1

440             QueryData.MoveNext
            Wend

        End If
        
        'User mail
        'TODO:
        
        ' Llaves
442     Call MakeQuery("SELECT key_obj FROM house_key WHERE account_id = " & .AccountID & ";")

444     If Not QueryData Is Nothing Then
446         QueryData.MoveFirst

448         LoopC = 1

450         While Not QueryData.EOF
452             .Keys(LoopC) = QueryData!key_obj
454             LoopC = LoopC + 1

456             QueryData.MoveNext
            Wend

        End If

    End With

    Exit Sub

ErrorHandler:
458 Call LogDatabaseError("Error en LoadUserDatabase: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.description & ". Línea: " & Erl)
460 Resume Next

End Sub

Private Sub MakeQuery(query As String, Optional ByVal NoResult As Boolean = False)
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Hace una unica query a la db. Asume una conexion.
        ' Si NoResult = False, el metodo lee el resultado de la query
        ' Guarda el resultado en QueryData
    
        On Error GoTo ErrorHandler

        ' Evito memory leaks
100     If Not QueryData Is Nothing Then
102         Call QueryData.Close
104         Set QueryData = Nothing

        End If
    
106     If NoResult Then
108         Call Database_Connection.Execute(query, RecordsAffected)

        Else
110         Set QueryData = Database_Connection.Execute(query, RecordsAffected)

112         If QueryData.BOF Or QueryData.EOF Then
114             Set QueryData = Nothing

            End If

        End If
    
        Exit Sub
    
ErrorHandler:

116     If Not adoIsConnected(Database_Connection) Then
118         Call LogDatabaseError("Alarma en MakeQuery: Se perdió la conexión con la DB. Reconectando.")
120         Database_Connect
122         Resume
        Else
124         Call LogDatabaseError("Error en MakeQuery: query = '" & query & "'. " & Err.Number & " - " & Err.description)
        
    On Error GoTo 0
126         Err.raise Err.Number
        End If

End Sub

Private Function GetDBValue(Tabla As String, ColumnaGet As String, ColumnaTest As String, ValueTest As Variant) As Variant
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para leer un unico valor de una unica fila

        On Error GoTo ErrorHandler
    
        'Hacemos la query segun el tipo de variable.
100     If VarType(ValueTest) = vbString Then
102         Call MakeQuery("SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = '" & ValueTest & "';")
        Else
104         Call MakeQuery("SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = " & ValueTest & ";")  ' Sin comillas

        End If
    
        'Revisamos si recibio un resultado
106     If QueryData Is Nothing Then Exit Function

        'Obtenemos la variable
108     GetDBValue = QueryData.Fields(0).Value

        Exit Function
    
ErrorHandler:
110     Call LogDatabaseError("Error en GetDBValue: SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = '" & ValueTest & "';" & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetCuentaValue(CuentaEmail As String, Columna As String) As Variant
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que leer un unico valor de la cuenta
        
        On Error GoTo GetCuentaValue_Err
        
100     GetCuentaValue = GetDBValue("account", Columna, "email", LCase$(CuentaEmail))

        
        Exit Function

GetCuentaValue_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetCuentaValue", Erl)
104     Resume Next
        
End Function

Public Function GetUserValue(CharName As String, Columna As String) As Variant
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que leer un unico valor del char
        
        On Error GoTo GetUserValue_Err
        
100     GetUserValue = GetDBValue("user", Columna, "name", CharName)

        
        Exit Function

GetUserValue_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetUserValue", Erl)
104     Resume Next
        
End Function

Private Sub SetDBValue(Tabla As String, ColumnaSet As String, ByVal ValueSet As Variant, ColumnaTest As String, ByVal ValueTest As Variant)
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para escribir un unico valor de una unica fila

        On Error GoTo ErrorHandler

        'Agregamos comillas a strings
100     If VarType(ValueSet) = vbString Then
102         ValueSet = "'" & ValueSet & "'"

        End If
    
104     If VarType(ValueTest) = vbString Then
106         ValueTest = "'" & ValueTest & "'"

        End If
    
        'Hacemos la query
108     Call MakeQuery("UPDATE " & Tabla & " SET " & ColumnaSet & " = " & ValueSet & " WHERE " & ColumnaTest & " = " & ValueTest & ";", True)

        Exit Sub
    
ErrorHandler:
110     Call LogDatabaseError("Error en SetDBValue: UPDATE " & Tabla & " SET " & ColumnaSet & " = " & ValueSet & " WHERE " & ColumnaTest & " = " & ValueTest & ";" & ". " & Err.Number & " - " & Err.description)

End Sub

Private Sub SetCuentaValue(CuentaEmail As String, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor de la cuenta
        
        On Error GoTo SetCuentaValue_Err
        
100     Call SetDBValue("account", Columna, Value, "email", LCase$(CuentaEmail))

        
        Exit Sub

SetCuentaValue_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.SetCuentaValue", Erl)
104     Resume Next
        
End Sub

Private Sub SetUserValue(CharName As String, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor del char
        
        On Error GoTo SetUserValue_Err
        
100     Call SetDBValue("user", Columna, Value, "name", CharName)

        
        Exit Sub

SetUserValue_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.SetUserValue", Erl)
104     Resume Next
        
End Sub

Private Sub SetCuentaValueByID(AccountID As Long, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor de la cuenta
        ' Por ID
        
        On Error GoTo SetCuentaValueByID_Err
        
100     Call SetDBValue("account", Columna, Value, "id", AccountID)

        
        Exit Sub

SetCuentaValueByID_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.SetCuentaValueByID", Erl)
104     Resume Next
        
End Sub

Private Sub SetUserValueByID(Id As Long, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor del char
        ' Por ID
        
        On Error GoTo SetUserValueByID_Err
        
100     Call SetDBValue("user", Columna, Value, "id", Id)

        
        Exit Sub

SetUserValueByID_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.SetUserValueByID", Erl)
104     Resume Next
        
End Sub

Public Function CheckUserDonatorDatabase(CuentaEmail As String) As Boolean
        
        On Error GoTo CheckUserDonatorDatabase_Err
        
100     CheckUserDonatorDatabase = GetCuentaValue(CuentaEmail, "is_donor")

        
        Exit Function

CheckUserDonatorDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.CheckUserDonatorDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetUserCreditosDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetUserCreditosDatabase_Err
        
100     GetUserCreditosDatabase = GetCuentaValue(CuentaEmail, "credits")

        
        Exit Function

GetUserCreditosDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetUserCreditosDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetUserCreditosCanjeadosDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetUserCreditosCanjeadosDatabase_Err
        
100     GetUserCreditosCanjeadosDatabase = GetCuentaValue(CuentaEmail, "credits_used")

        
        Exit Function

GetUserCreditosCanjeadosDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetUserCreditosCanjeadosDatabase", Erl)
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
106     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetUserDiasDonadorDatabase", Erl)
108     Resume Next
        
End Function

Public Function GetUserComprasDonadorDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetUserComprasDonadorDatabase_Err
        
100     GetUserComprasDonadorDatabase = GetCuentaValue(CuentaEmail, "donor_purchases")

        
        Exit Function

GetUserComprasDonadorDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetUserComprasDonadorDatabase", Erl)
104     Resume Next
        
End Function

Public Function CheckUserExists(name As String) As Boolean
        
        On Error GoTo CheckUserExists_Err
        
100     CheckUserExists = GetUserValue(name, "COUNT(*)") > 0

        
        Exit Function

CheckUserExists_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.CheckUserExists", Erl)
104     Resume Next
        
End Function

Public Function CheckCuentaExiste(CuentaEmail As String) As Boolean
        
        On Error GoTo CheckCuentaExiste_Err
        
100     CheckCuentaExiste = GetCuentaValue(CuentaEmail, "COUNT(*)") > 0

        
        Exit Function

CheckCuentaExiste_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.CheckCuentaExiste", Erl)
104     Resume Next
        
End Function

Public Function BANCheckDatabase(name As String) As Boolean
        
        On Error GoTo BANCheckDatabase_Err
        
100     BANCheckDatabase = CBool(GetUserValue(name, "is_banned"))

        
        Exit Function

BANCheckDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.BANCheckDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetCodigoActivacionDatabase(name As String) As String
        
        On Error GoTo GetCodigoActivacionDatabase_Err
        
100     GetCodigoActivacionDatabase = GetCuentaValue(name, "validate_code")

        
        Exit Function

GetCodigoActivacionDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetCodigoActivacionDatabase", Erl)
104     Resume Next
        
End Function

Public Function CheckCuentaActivadaDatabase(name As String) As Boolean
        
        On Error GoTo CheckCuentaActivadaDatabase_Err
        
100     CheckCuentaActivadaDatabase = GetCuentaValue(name, "validated")

        
        Exit Function

CheckCuentaActivadaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.CheckCuentaActivadaDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetEmailDatabase(name As String) As String
        
        On Error GoTo GetEmailDatabase_Err
        
100     GetEmailDatabase = GetCuentaValue(name, "email")

        
        Exit Function

GetEmailDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetEmailDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetMacAddressDatabase(CuentaEmail As String) As String
        
        On Error GoTo GetMacAddressDatabase_Err
        
100     GetMacAddressDatabase = GetCuentaValue(CuentaEmail, "mac_address")

        
        Exit Function

GetMacAddressDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetMacAddressDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetHDSerialDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetHDSerialDatabase_Err
        
100     GetHDSerialDatabase = GetCuentaValue(CuentaEmail, "hd_serial")

        
        Exit Function

GetHDSerialDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetHDSerialDatabase", Erl)
104     Resume Next
        
End Function

Public Function CheckBanCuentaDatabase(CuentaEmail As String) As Boolean
        
        On Error GoTo CheckBanCuentaDatabase_Err
        
100     CheckBanCuentaDatabase = CBool(GetCuentaValue(CuentaEmail, "is_banned"))

        
        Exit Function

CheckBanCuentaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.CheckBanCuentaDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetMotivoBanCuentaDatabase(CuentaEmail As String) As String
        
        On Error GoTo GetMotivoBanCuentaDatabase_Err
        
100     GetMotivoBanCuentaDatabase = GetCuentaValue(CuentaEmail, "ban_reason")

        
        Exit Function

GetMotivoBanCuentaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetMotivoBanCuentaDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetQuienBanCuentaDatabase(CuentaEmail As String) As String
        
        On Error GoTo GetQuienBanCuentaDatabase_Err
        
100     GetQuienBanCuentaDatabase = GetCuentaValue(CuentaEmail, "banned_by")

        
        Exit Function

GetQuienBanCuentaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetQuienBanCuentaDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetCuentaLogeadaDatabase(CuentaEmail As String) As Boolean
        
        On Error GoTo GetCuentaLogeadaDatabase_Err
        
100     GetCuentaLogeadaDatabase = GetCuentaValue(CuentaEmail, "is_logged")

        
        Exit Function

GetCuentaLogeadaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetCuentaLogeadaDatabase", Erl)
104     Resume Next
        
End Function

Public Function GetUserStatusDatabase(name As String) As Integer
        
        On Error GoTo GetUserStatusDatabase_Err
        
100     GetUserStatusDatabase = GetUserValue(name, "status")

        
        Exit Function

GetUserStatusDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetUserStatusDatabase", Erl)
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
108     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetAccountIDDatabase", Erl)
110     Resume Next
        
End Function

Public Sub GetPasswordAndSaltDatabase(CuentaEmail As String, PasswordHash As String, Salt As String)

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("SELECT password, salt FROM account WHERE deleted = FALSE AND email = '" & LCase$(CuentaEmail) & "';")

102     If QueryData Is Nothing Then Exit Sub
    
104     PasswordHash = QueryData!Password
106     Salt = QueryData!Salt
    
        Exit Sub
    
ErrorHandler:
108     Call LogDatabaseError("Error in GetPasswordAndSaltDatabase. Cuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.description)
    
End Sub

Public Function GetPersonajesCountDatabase(CuentaEmail As String) As Byte

        On Error GoTo ErrorHandler

        Dim Id As Integer

100     Id = GetDBValue("account", "id", "email", LCase$(CuentaEmail))
    
102     GetPersonajesCountDatabase = GetPersonajesCountByIDDatabase(Id)
    
        Exit Function
    
ErrorHandler:
104     Call LogDatabaseError("Error in GetPersonajesCountDatabase. name: " & CuentaEmail & ". " & Err.Number & " - " & Err.description)
    
End Function

Public Function GetPersonajesCountByIDDatabase(ByVal AccountID As Long) As Byte

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("SELECT COUNT(*) FROM user WHERE deleted = FALSE AND account_id = " & AccountID & ";")
    
102     If QueryData Is Nothing Then Exit Function
    
104     GetPersonajesCountByIDDatabase = QueryData.Fields(0).Value
    
        Exit Function
    
ErrorHandler:
106     Call LogDatabaseError("Error in GetPersonajesCountByIDDatabase. AccountID: " & AccountID & ". " & Err.Number & " - " & Err.description)
    
End Function

Public Function GetPersonajesCuentaDatabase(ByVal AccountID As Long, Personaje() As PersonajeCuenta) As Byte
        
        On Error GoTo GetPersonajesCuentaDatabase_Err
        

100     Call MakeQuery("SELECT name, head_id, class_id, body_id, pos_map, level, status, helmet_id, shield_id, weapon_id, guild_index, is_dead, is_sailing FROM user WHERE deleted = FALSE AND account_id = " & AccountID & ";")

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
120         Personaje(i).nivel = QueryData!level
122         Personaje(i).Status = QueryData!Status
124         Personaje(i).Casco = QueryData!helmet_id
126         Personaje(i).Escudo = QueryData!shield_id
128         Personaje(i).Arma = QueryData!weapon_id
130         Personaje(i).ClanIndex = QueryData!Guild_Index
        
132         If EsRolesMaster(Personaje(i).nombre) Then
134             Personaje(i).Status = 3
136         ElseIf EsConsejero(Personaje(i).nombre) Then
138             Personaje(i).Status = 4
140         ElseIf EsSemiDios(Personaje(i).nombre) Then
142             Personaje(i).Status = 5
144         ElseIf EsDios(Personaje(i).nombre) Then
146             Personaje(i).Status = 6
148         ElseIf EsAdmin(Personaje(i).nombre) Then
150             Personaje(i).Status = 7

            End If

152         If val(QueryData!is_dead) = 1 Or val(QueryData!is_sailing) = 1 Then
154             Personaje(i).Cabeza = 0
            End If
        
156         QueryData.MoveNext
        Next

        
        Exit Function

GetPersonajesCuentaDatabase_Err:
158     Call RegistrarError(Err.Number, Err.description, "modDatabase.GetPersonajesCuentaDatabase", Erl)
160     Resume Next
        
End Function

Public Sub SetUserLoggedDatabase(ByVal Id As Long, ByVal AccountID As Long)
        
        On Error GoTo SetUserLoggedDatabase_Err
        
100     Call SetDBValue("user", "is_logged", 1, "id", Id)
102     Call MakeQuery("UPDATE account SET logged = logged + 1 WHERE id = " & AccountID & ";", True)

        
        Exit Sub

SetUserLoggedDatabase_Err:
104     Call RegistrarError(Err.Number, Err.description, "modDatabase.SetUserLoggedDatabase", Erl)
106     Resume Next
        
End Sub

Public Sub ResetLoggedDatabase(ByVal AccountID As Long)
        
        On Error GoTo ResetLoggedDatabase_Err
        
100     Call MakeQuery("UPDATE account SET logged = 0 WHERE id = " & AccountID & ";", True)

        
        Exit Sub

ResetLoggedDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.ResetLoggedDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub SetUsersLoggedDatabase(ByVal NumUsers As Long)
        
        On Error GoTo SetUsersLoggedDatabase_Err
        
100     Call MakeQuery("UPDATE statistics SET value = '" & NumUsers & "' WHERE name = 'online';", True)
        
        Exit Sub

SetUsersLoggedDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.SetUsersLoggedDatabase", Erl)
104     Resume Next
        
End Sub

Public Function LeerRecordUsuariosDatabase() As Long
        
        On Error GoTo LeerRecordUsuariosDatabase_Err
        
100     Call MakeQuery("SELECT value FROM statistics WHERE name = 'record';")

102     If QueryData Is Nothing Then Exit Function

104     LeerRecordUsuariosDatabase = val(QueryData!Value)

        Exit Function

LeerRecordUsuariosDatabase_Err:
106     Call RegistrarError(Err.Number, Err.description, "modDatabase.LeerRecordUsuariosDatabase", Erl)
108     Resume Next
        
End Function

Public Sub SetRecordUsersDatabase(ByVal Record As Long)
        
        On Error GoTo SetRecordUsersDatabase_Err
        
100     Call MakeQuery("UPDATE statistics SET value = '" & Record & "' WHERE name = 'record';", True)
        
        Exit Sub

SetRecordUsersDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.SetRecordUsersDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub LogoutAllUsersAndAccounts()
        
        On Error GoTo LogoutAllUsersAndAccounts_Err
        

        Dim query As String

100     query = "UPDATE user SET is_logged = FALSE; "
102     query = query & "UPDATE account SET logged = 0;"

104     Call MakeQuery(query, True)

        
        Exit Sub

LogoutAllUsersAndAccounts_Err:
106     Call RegistrarError(Err.Number, Err.description, "modDatabase.LogoutAllUsersAndAccounts", Erl)
108     Resume Next
        
End Sub

Public Sub SaveBattlePointsDatabase(ByVal Id As Long, ByVal BattlePoints As Long)
        
        On Error GoTo SaveBattlePointsDatabase_Err
        
100     Call SetDBValue("user", "battle_points", BattlePoints, "id", Id)

        
        Exit Sub

SaveBattlePointsDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.SaveBattlePointsDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub SaveVotoDatabase(ByVal Id As Long, ByVal Encuestas As Integer)
        
        On Error GoTo SaveVotoDatabase_Err
        
100     Call SetUserValueByID(Id, "votes_amount", Encuestas)

        
        Exit Sub

SaveVotoDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.SaveVotoDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub SaveUserBodyDatabase(UserName As String, ByVal Body As Integer)
        
        On Error GoTo SaveUserBodyDatabase_Err
        
100     Call SetUserValue(UserName, "body_id", Body)

        
        Exit Sub

SaveUserBodyDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.SaveUserBodyDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub SaveUserHeadDatabase(UserName As String, ByVal Head As Integer)
        
        On Error GoTo SaveUserHeadDatabase_Err
        
100     Call SetUserValue(UserName, "head_id", Head)

        
        Exit Sub

SaveUserHeadDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.SaveUserHeadDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub SaveUserSkillDatabase(UserName As String, ByVal Skill As Integer, ByVal Value As Integer)
        
        On Error GoTo SaveUserSkillDatabase_Err
        
100     Call MakeQuery("UPDATE skillpoints SET value = " & Value & " WHERE number = " & Skill & " AND user_id = (SELECT id FROM user WHERE name = '" & UserName & "');", True)

        
        Exit Sub

SaveUserSkillDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.SaveUserSkillDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub SaveUserSkillsLibres(UserName As String, ByVal SkillsLibres As Integer)
        
        On Error GoTo SaveUserHeadDatabase_Err
        
100     Call SetUserValue(UserName, "SaveUserSkillsLibres", SkillsLibres)

        
        Exit Sub

SaveUserHeadDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.SaveUserHeadDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub SaveNewAccountDatabase(CuentaEmail As String, PasswordHash As String, Salt As String, Codigo As String)

        On Error GoTo ErrorHandler

        Dim q As String

        'Basic user data
100     q = "INSERT INTO account SET "
102     q = q & "email = '" & LCase$(CuentaEmail) & "', "
104     q = q & "password = '" & PasswordHash & "', "
106     q = q & "salt = '" & Salt & "', "
108     q = q & "validate_code = '" & Codigo & "', "
110     q = q & "date_created = NOW();"

112     Call MakeQuery(q, True)
    
        Exit Sub
        
ErrorHandler:
114     Call LogDatabaseError("Error en SaveNewAccountDatabase. Cuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub ValidarCuentaDatabase(UserCuenta As String)
        
        On Error GoTo ValidarCuentaDatabase_Err
        
100     Call SetCuentaValue(UserCuenta, "validated", 1)

        
        Exit Sub

ValidarCuentaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.ValidarCuentaDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub BorrarUsuarioDatabase(name As String)

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("UPDATE user SET name = CONCAT('DELETED_', name), deleted = TRUE WHERE name = '" & name & "';", True)

        Exit Sub
    
ErrorHandler:
102     Call LogDatabaseError("Error en BorrarUsuarioDatabase borrando user de la Mysql Database: " & name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub BorrarCuentaDatabase(CuentaEmail As String)

        On Error GoTo ErrorHandler

        Dim Id As Integer

100     Id = GetDBValue("account", "id", "email", LCase$(CuentaEmail))

        Dim query As String
    
102     query = "UPDATE account SET email = CONCAT('DELETED_', email), deleted = TRUE WHERE email = '" & LCase$(CuentaEmail) & "'; "
    
104     query = query & "UPDATE user SET name = CONCAT('DELETED_', name), deleted = TRUE WHERE account_id = '" & Id & "';"
    
106     Call MakeQuery(query, True)

        Exit Sub
    
ErrorHandler:
108     Call LogDatabaseError("Error en BorrarCuentaDatabase borrando user de la Mysql Database: " & CuentaEmail & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveBanDatabase(ByVal UserName As String, ByVal Reason As String, ByVal BannedBy As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Dim query As String

100     query = "UPDATE user SET is_banned = TRUE WHERE name = '" & UserName & "'; "

102     query = query & "INSERT INTO punishment SET "
104     query = query & "user_id = (SELECT id from user WHERE name = '" & UserName & "'), "
106     query = query & "number = number + 1, "
108     query = query & "reason = '" & BannedBy & ": " & LCase$(Reason) & " " & Date & " " & Time & "';"

110     Call MakeQuery(query, True)

        Exit Sub

ErrorHandler:
112     Call LogDatabaseError("Error in SaveBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveWarnDatabase(ByVal UserName As String, ByVal Reason As String, ByVal WarnedBy As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Dim query As String

100     query = "UPDATE user SET warnings = warnings + 1 WHERE name = '" & UserName & "'; "

102     query = query & "INSERT INTO punishment SET "
104     query = query & "user_id = (SELECT id from user WHERE name = '" & UserName & "'), "
106     query = query & "number = number + 1, "
108     query = query & "reason = '" & WarnedBy & ": " & LCase$(Reason) & " " & Date & " " & Time & "';"

110     Call MakeQuery(query, True)

        Exit Sub

ErrorHandler:
112     Call LogDatabaseError("Error in SaveWarnDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SavePenaDatabase(UserName As String, Reason As String)

        On Error GoTo ErrorHandler

        Dim query As String

100     query = query & "INSERT INTO punishment SET "
102     query = query & "user_id = (SELECT id from user WHERE name = '" & UserName & "'), "
104     query = query & "number = number + 1, "
106     query = query & "reason = '" & Reason & "';"

108     Call MakeQuery(query, True)

        Exit Sub

ErrorHandler:
110     Call LogDatabaseError("Error in SavePenaDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub UnBanDatabase(UserName As String)

        On Error GoTo ErrorHandler

100     Call MakeQuery("UPDATE user SET is_banned = FALSE WHERE name = '" & UserName & "';", True)

        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in UnBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub EcharConsejoDatabase(UserName As String)
        
        On Error GoTo EcharConsejoDatabase_Err
        
100     Call MakeQuery("UPDATE user SET pertenece_consejo_real = FALSE, pertenece_consejo_caos = FALSE WHERE name = '" & UserName & "';", True)

        
        Exit Sub

EcharConsejoDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.EcharConsejoDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub EcharLegionDatabase(UserName As String)
        
        On Error GoTo EcharLegionDatabase_Err
        
100     Call MakeQuery("UPDATE user SET pertenece_caos = FALSE, reenlistadas = 200 WHERE name = '" & UserName & "';", True)

        
        Exit Sub

EcharLegionDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.EcharLegionDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub EcharArmadaDatabase(UserName As String)
        
        On Error GoTo EcharArmadaDatabase_Err
        
100     Call MakeQuery("UPDATE user SET pertenece_real = FALSE, reenlistadas = 200 WHERE name = '" & UserName & "';", True)

        
        Exit Sub

EcharArmadaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.EcharArmadaDatabase", Erl)
104     Resume Next
        
End Sub

Public Sub CambiarPenaDatabase(UserName As String, ByVal Numero As Integer, Pena As String)
        
        On Error GoTo CambiarPenaDatabase_Err
        
100     Call MakeQuery("UPDATE punishment SET reason = '" & Pena & "' WHERE number = " & Numero & " AND user_id = (SELECT id from user WHERE name = '" & UserName & "');", True)

        
        Exit Sub

CambiarPenaDatabase_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.CambiarPenaDatabase", Erl)
104     Resume Next
        
End Sub

Public Function GetUserAmountOfPunishmentsDatabase(ByVal UserName As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call MakeQuery("SELECT COUNT(*) as punishments FROM punishment WHERE user_id = (SELECT id from user WHERE name = '" & UserName & "');")

102     If QueryData Is Nothing Then Exit Function

104     GetUserAmountOfPunishmentsDatabase = QueryData!punishments

        Exit Function
ErrorHandler:
106     Call LogDatabaseError("Error in GetUserAmountOfPunishmentsDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SendUserPunishmentsDatabase(ByVal UserIndex As Integer, ByVal UserName As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call MakeQuery("SELECT * FROM punishment WHERE user_id = (SELECT id from user WHERE UPPER(name) = '" & UCase$(UserName) & "');")
    
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
114     Call LogDatabaseError("Error in SendUserPunishmentsDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub


Public Function GetNombreCuentaDatabase(name As String) As String

        On Error GoTo ErrorHandler

        'Hacemos la query.
100     Call MakeQuery("SELECT email FROM account WHERE id = (SELECT account_id FROM user WHERE name = '" & name & "');")
    
        'Verificamos que la query no devuelva un resultado vacio.
102     If QueryData Is Nothing Then Exit Function
    
        'Obtenemos el nombre de la cuenta
104     GetNombreCuentaDatabase = QueryData!name

        Exit Function
    
ErrorHandler:
106     Call LogDatabaseError("Error en GetNombreCuentaDatabase leyendo user de la Mysql Database: " & name & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildIndexDatabase(ByVal UserName As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 09/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call MakeQuery("SELECT guild_index FROM user WHERE UPPER(name) = '" & UCase$(UserName) & "';")

102     If QueryData Is Nothing Then Exit Function

104     GetUserGuildIndexDatabase = SanitizeNullValue(QueryData!Guild_Index, 0)

        Exit Function

ErrorHandler:
106     Call LogDatabaseError("Error in GetUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildMemberDatabase(ByVal UserName As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call MakeQuery("SELECT guild_member_history FROM user WHERE name = '" & UserName & "';")

102     If QueryData Is Nothing Then Exit Function

104     GetUserGuildMemberDatabase = SanitizeNullValue(QueryData!guild_member_history, vbNullString)

        Exit Function

ErrorHandler:
106     Call LogDatabaseError("Error in GetUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildAspirantDatabase(ByVal UserName As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call MakeQuery("SELECT guild_aspirant_index FROM user WHERE name = '" & UserName & "';")

102     If QueryData Is Nothing Then Exit Function

104     GetUserGuildAspirantDatabase = SanitizeNullValue(QueryData!guild_aspirant_index, 0)

        Exit Function

ErrorHandler:
106     Call LogDatabaseError("Error in GetUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildRejectionReasonDatabase(ByVal UserName As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler
    
100     Call MakeQuery("SELECT guild_rejected_because FROM user WHERE name = '" & UserName & "';")

102     If QueryData Is Nothing Then Exit Function

104     GetUserGuildRejectionReasonDatabase = SanitizeNullValue(QueryData!guild_rejected_because, vbNullString)

        Exit Function

ErrorHandler:
106     Call LogDatabaseError("Error in GetUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildPedidosDatabase(ByVal UserName As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call MakeQuery("SELECT guild_requests_history FROM user WHERE name = '" & UserName & "';")

102     If QueryData Is Nothing Then Exit Function

104     GetUserGuildPedidosDatabase = SanitizeNullValue(QueryData!guild_requests_history, vbNullString)

        Exit Function

ErrorHandler:
106     Call LogDatabaseError("Error in GetUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserGuildRejectionReasonDatabase(ByVal UserName As String, ByVal Reason As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Dim query As String

100     query = "UPDATE user SET "
102     query = query & "guild_rejected_because = '" & Reason & "' "
104     query = query & "WHERE name = '" & UserName & "';"

106     Call MakeQuery(query, True)

        Exit Sub
ErrorHandler:
108     Call LogDatabaseError("Error in SaveUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildIndexDatabase(ByVal UserName As String, ByVal GuildIndex As Integer)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Dim query As String

100     query = "UPDATE user SET "
102     query = query & "guild_index = " & GuildIndex & " "
104     query = query & "WHERE name = '" & UserName & "';"
    
106     Call MakeQuery(query, True)

        Exit Sub
ErrorHandler:
108     Call LogDatabaseError("Error in SaveUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildAspirantDatabase(ByVal UserName As String, ByVal AspirantIndex As Integer)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Dim query As String

100     query = "UPDATE user SET "
102     query = query & "guild_aspirant_index = " & AspirantIndex & " "
104     query = query & "WHERE name = '" & UserName & "';"

106     Call MakeQuery(query, True)

        Exit Sub
ErrorHandler:
108     Call LogDatabaseError("Error in SaveUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildMemberDatabase(ByVal UserName As String, ByVal guilds As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Dim query As String

100     query = "UPDATE user SET "
102     query = query & "guild_member_history = '" & guilds & "' "
104     query = query & "WHERE name = '" & UserName & "';"

106     Call MakeQuery(query, True)

        Exit Sub
ErrorHandler:
108     Call LogDatabaseError("Error in SaveUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildPedidosDatabase(ByVal UserName As String, ByVal Pedidos As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Dim query As String

100     query = "UPDATE user SET "
102     query = query & "guild_requests_history = '" & Pedidos & "' "
104     query = query & "WHERE name = '" & UserName & "';"

106     Call MakeQuery(query, True)

        Exit Sub
ErrorHandler:
108     Call LogDatabaseError("Error in SaveUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

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

100     Call MakeQuery("SELECT race_id, class_id, genre_id, level, gold, bank_gold, guild_requests_history, guild_index, guild_member_history, pertenece_real, pertenece_caos, ciudadanos_matados, criminales_matados FROM user WHERE name = '" & UserName & "';")

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

120     Call Protocol.WriteCharacterInfo(UserIndex, UserName, QueryData!race_id, QueryData!class_id, QueryData!genre_id, QueryData!level, QueryData!gold, QueryData!bank_gold, SanitizeNullValue(QueryData!guild_requests_history, vbNullString), gName, Miembro, QueryData!pertenece_real, QueryData!pertenece_caos, QueryData!ciudadanos_matados, QueryData!criminales_matados)

        Exit Sub
ErrorHandler:
122     Call LogDatabaseError("Error in SendCharacterInfoDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function EnterAccountDatabase(ByVal UserIndex As Integer, CuentaEmail As String, Password As String, MacAddress As String, ByVal HDserial As Long, ip As String) As Boolean

        On Error GoTo ErrorHandler
    
100     Call MakeQuery("SELECT id, password, salt, validated, is_banned, ban_reason, banned_by FROM account WHERE email = '" & LCase$(CuentaEmail) & "';")
    
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
    
122     UserList(UserIndex).AccountID = QueryData!Id
124     UserList(UserIndex).Cuenta = CuentaEmail
    
126     Call MakeQuery("UPDATE account SET mac_address = '" & MacAddress & "', hd_serial = " & HDserial & ", last_ip = '" & ip & "', last_access = NOW() WHERE id = " & QueryData!Id & ";", True)
    
128     EnterAccountDatabase = True
    
        Exit Function

ErrorHandler:
130     Call LogDatabaseError("Error in EnterAccountDatabase. UserCuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub ChangePasswordDatabase(ByVal UserIndex As Integer, OldPassword As String, NewPassword As String)

        On Error GoTo ErrorHandler

100     If LenB(NewPassword) = 0 Then
102         Call WriteConsoleMsg(UserIndex, "Debe especificar una contraseña nueva, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
104     Call MakeQuery("SELECT password, salt FROM account WHERE id = " & UserList(UserIndex).AccountID & ";")
    
106     If QueryData Is Nothing Then
108         Call WriteConsoleMsg(UserIndex, "No se ha podido cambiar la contraseña por un error interno. Avise a un administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
110     If Not PasswordValida(OldPassword, QueryData!Password, QueryData!Salt) Then
112         Call WriteConsoleMsg(UserIndex, "La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        Dim Salt As String * 10

114     Salt = RandomString(10) ' Alfanumerico
    
        Dim oSHA256 As CSHA256

116     Set oSHA256 = New CSHA256

        Dim PasswordHash As String * 64

118     PasswordHash = oSHA256.SHA256(NewPassword & Salt)
    
120     Set oSHA256 = Nothing
    
122     Call MakeQuery("UPDATE account SET password = '" & PasswordHash & "', salt = '" & Salt & "' WHERE id = " & UserList(UserIndex).AccountID & ";", True)
    
124     Call WriteConsoleMsg(UserIndex, "La contraseña de su cuenta fue cambiada con éxito.", FontTypeNames.FONTTYPE_INFO)
    
        Exit Sub

ErrorHandler:
126     Call LogDatabaseError("Error in ChangePasswordDatabase. Username: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUsersLoggedAccountDatabase(ByVal AccountID As Integer) As Byte

        On Error GoTo ErrorHandler

100     Call GetDBValue("account", "logged", "id", AccountID)
    
102     If QueryData Is Nothing Then Exit Function
    
104     GetUsersLoggedAccountDatabase = val(QueryData!logged)

        Exit Function

ErrorHandler:
106     Call LogDatabaseError("Error in GetUsersLoggedAccountDatabase. AccountID: " & AccountID & ". " & Err.Number & " - " & Err.description)

End Function

Public Function SetPositionDatabase(UserName As String, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        On Error GoTo ErrorHandler

100     Call MakeQuery("UPDATE user SET pos_map = " & Map & ", pos_x = " & X & ", pos_y = " & X & " WHERE UPPER(name) = '" & UCase$(UserName) & "';", True)
    
102     SetPositionDatabase = RecordsAffected > 0

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in SetPositionDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function AddOroBancoDatabase(UserName As String, ByVal OroGanado As Long) As Boolean
        On Error GoTo ErrorHandler

100     Call MakeQuery("UPDATE user SET bank_gold = bank_gold + " & OroGanado & " WHERE UPPER(name) = '" & UCase$(UserName) & "';", True)
    
102     AddOroBancoDatabase = RecordsAffected > 0

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in AddOroBancoDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function DarLlaveAUsuarioDatabase(UserName As String, ByVal LlaveObj As Integer) As Boolean
        On Error GoTo ErrorHandler

100     Call MakeQuery("INSERT INTO house_key SET key_obj = " & LlaveObj & ", account_id = (SELECT account_id FROM user WHERE UPPER(name) = '" & UCase$(UserName) & "');", True)
    
102     DarLlaveAUsuarioDatabase = RecordsAffected > 0

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in DarLlaveAUsuarioDatabase. UserName: " & UserName & ", LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.description)

End Function

Public Function DarLlaveACuentaDatabase(email As String, ByVal LlaveObj As Integer) As Boolean
        On Error GoTo ErrorHandler

100     Call MakeQuery("INSERT INTO house_key SET key_obj = " & LlaveObj & ", account_id = (SELECT id FROM account WHERE UPPER(email) = '" & UCase$(email) & "');", True)
    
102     DarLlaveACuentaDatabase = RecordsAffected > 0
        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in DarLlaveACuentaDatabase. Email: " & email & ", LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.description)

End Function

Public Function SacarLlaveDatabase(ByVal LlaveObj As Integer) As Boolean
        On Error GoTo ErrorHandler

        Dim i As Integer
        Dim UserCount As Integer
        Dim Users() As String

        ' Obtengo los usuarios logueados en la cuenta del dueño de la llave
100     Call MakeQuery("SELECT name FROM user WHERE is_logged = TRUE AND account_id = (SELECT account_id FROM house_key WHERE key_obj = " & LlaveObj & ");")
    
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
120     Call MakeQuery("DELETE FROM house_key WHERE key_obj = " & LlaveObj & ";", True)
    
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
132     Call LogDatabaseError("Error in SacarLlaveDatabase. LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub VerLlavesDatabase(ByVal UserIndex As Integer)
        On Error GoTo ErrorHandler

100     Call MakeQuery("SELECT (SELECT email FROM account WHERE id = K.account_id) as email, key_obj FROM house_key AS K;")

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
124     Call LogDatabaseError("Error in VerLlavesDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function SanitizeNullValue(ByVal Value As Variant, ByVal defaultValue As Variant) As Variant
        
        On Error GoTo SanitizeNullValue_Err
        
100     SanitizeNullValue = IIf(IsNull(Value), defaultValue, Value)

        
        Exit Function

SanitizeNullValue_Err:
102     Call RegistrarError(Err.Number, Err.description, "modDatabase.SanitizeNullValue", Erl)
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
        Dim Cmd As New ADODB.Command

        'Set up SQL command to return 1
100     Cmd.CommandText = "SELECT 1"
102     Cmd.ActiveConnection = adoCn

        'Run a simple query, to test the connection
        
104     i = Cmd.Execute.Fields(0)
        On Error GoTo 0

        'Tidy up
106     Set Cmd = Nothing

        'If i is 1, connection is open
108     If i = 1 Then
110         adoIsConnected = True
        Else
112         adoIsConnected = False
        End If

End Function
