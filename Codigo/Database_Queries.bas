Attribute VB_Name = "Database_Queries"
Option Explicit

'Constructor de queries.
'Me permite concatenar strings MUCHO MAS rapido
Private QueryBuilder As cStringBuilder

' DYNAMIC QUERIES
Public QUERY_SAVE_MAINPJ As String
Public QUERY_SAVE_SPELLS As String
Public QUERY_SAVE_INVENTORY As String
Public QUERY_SAVE_BANCOINV As String
Public QUERY_SAVE_SKILLS As String
Public QUERY_SAVE_QUESTS As String
Public QUERY_SAVE_PETS As String

Public QUERY_UPDATE_MAINPJ As String
Public QUERY_UPSERT_SPELLS As String
Public QUERY_UPSERT_INVENTORY As String
Public QUERY_UPSERT_SKILLS As String
Public QUERY_UPSERT_PETS As String

' CONSTANT QUERIES (NEW)
Public Const QUERY_INSERT_ATTRIBUTES As String = "INSERT INTO attribute VALUES (?, ?, ?, ?, ?, ?)"
Public Const QUERY_UPDATE_ATTRIBUTES As String = "UPDATE attribute SET strength = ?, agility = ?,  intelligence = ?, constitution = ?, charisma = ? WHERE user_id = ?"

Public Sub Contruir_Querys()
100     Call ConstruirQuery_CrearPersonaje
102     Call ConstruirQuery_GuardarPersonaje
End Sub

Private Sub ConstruirQuery_CrearPersonaje()
        Dim LoopC As Long
    
100     Set QueryBuilder = New cStringBuilder
    
        ' ************************** Basic user data ********************************
102     QueryBuilder.Append "INSERT INTO user SET "
104     QueryBuilder.Append "name = ?, "
106     QueryBuilder.Append "account_id = ?, "
108     QueryBuilder.Append "level = ?, "
110     QueryBuilder.Append "exp = ?, "
112     QueryBuilder.Append "genre_id = ?, "
114     QueryBuilder.Append "race_id = ?, "
116     QueryBuilder.Append "class_id = ?, "
118     QueryBuilder.Append "home_id = ?, "
120     QueryBuilder.Append "description = ?, "
122     QueryBuilder.Append "gold = ?, "
124     QueryBuilder.Append "free_skillpoints = ?, "
126     QueryBuilder.Append "pos_map = ?, "
128     QueryBuilder.Append "pos_x = ?, "
130     QueryBuilder.Append "pos_y = ?, "
132     QueryBuilder.Append "body_id = ?, "
134     QueryBuilder.Append "head_id = ?, "
136     QueryBuilder.Append "weapon_id = ?, "
138     QueryBuilder.Append "helmet_id = ?, "
140     QueryBuilder.Append "shield_id = ?, "
144     QueryBuilder.Append "slot_armour = ?, "
146     QueryBuilder.Append "slot_weapon = ?, "
148     QueryBuilder.Append "slot_shield = ?, "
150     QueryBuilder.Append "slot_helmet = ?, "
152     QueryBuilder.Append "slot_ammo = ?, "
154     QueryBuilder.Append "slot_dm = ?, "
156     QueryBuilder.Append "slot_rm = ?, "
158     QueryBuilder.Append "slot_tool = ?, "
160     QueryBuilder.Append "slot_magic = ?, "
162     QueryBuilder.Append "slot_knuckles = ?, "
164     QueryBuilder.Append "slot_ship = ?, "
166     QueryBuilder.Append "slot_mount = ?, "
168     QueryBuilder.Append "min_hp = ?, "
170     QueryBuilder.Append "max_hp = ?, "
172     QueryBuilder.Append "min_man = ?, "
174     QueryBuilder.Append "max_man = ?, "
176     QueryBuilder.Append "min_sta = ?, "
178     QueryBuilder.Append "max_sta = ?, "
180     QueryBuilder.Append "min_ham = ?, "
182     QueryBuilder.Append "max_ham = ?, "
184     QueryBuilder.Append "min_sed = ?, "
186     QueryBuilder.Append "max_sed = ?, "
188     QueryBuilder.Append "min_hit = ?, "
190     QueryBuilder.Append "max_hit = ?, "
192     QueryBuilder.Append "is_naked = ?, "
194     QueryBuilder.Append "status = ?, "
196     QueryBuilder.Append "is_logged = TRUE;"
    
        ' Guardo la query ensamblada
198     QUERY_SAVE_MAINPJ = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
200     Call QueryBuilder.Clear

        ' ************************** User spells ************************************
218     QueryBuilder.Append "INSERT INTO spell (user_id, number, spell_id) VALUES "

220     For LoopC = 1 To MAXUSERHECHIZOS
222         QueryBuilder.Append "(?, ?, ?)"

224         If LoopC < MAXUSERHECHIZOS Then
226             QueryBuilder.Append ", "
            End If

228     Next LoopC
    
        ' Guardo la query ensamblada
230     QUERY_SAVE_SPELLS = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
232     Call QueryBuilder.Clear
    
        ' ******************* INVENTORY *******************
234     QueryBuilder.Append "INSERT INTO inventory_item (user_id, number, item_id, Amount, is_equipped) VALUES "

236     For LoopC = 1 To MAX_INVENTORY_SLOTS
238         QueryBuilder.Append "(?, ?, ?, ?, ?)"

240         If LoopC < MAX_INVENTORY_SLOTS Then
242             QueryBuilder.Append ", "
            End If

244     Next LoopC
    
        ' Guardo la query ensamblada
246     QUERY_SAVE_INVENTORY = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
248     Call QueryBuilder.Clear

        ' ************************** User skills ************************************
250     QueryBuilder.Append "INSERT INTO skillpoint (user_id, number, value) VALUES "

252     For LoopC = 1 To NUMSKILLS
254         QueryBuilder.Append "(?, ?, ?)"

256         If LoopC < NUMSKILLS Then
258             QueryBuilder.Append ", "
            End If

260     Next LoopC

        ' Guardo la query ensamblada
262     QUERY_SAVE_SKILLS = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
264     Call QueryBuilder.Clear

        ' ************************** User quests ************************************
266     QueryBuilder.Append "INSERT INTO quest (user_id, number) VALUES "

268     For LoopC = 1 To MAXUSERQUESTS
270         QueryBuilder.Append "(?, ?)"

272         If LoopC < MAXUSERQUESTS Then
274             QueryBuilder.Append ", "
            End If

276     Next LoopC
    
        ' Guardo la query ensamblada
278     QUERY_SAVE_QUESTS = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
280     Call QueryBuilder.Clear
    
        ' ************************** User pets **************************************
282     QueryBuilder.Append "INSERT INTO pet (user_id, number, pet_id) VALUES "

284     For LoopC = 1 To MAXMASCOTAS
286         QueryBuilder.Append "(?, ?, ?)"

288         If LoopC < MAXMASCOTAS Then
290             QueryBuilder.Append ", "
            End If

292     Next LoopC
    
        ' Guardo la query ensamblada
294     QUERY_SAVE_PETS = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
296     Call QueryBuilder.Clear
    
End Sub

Private Sub ConstruirQuery_GuardarPersonaje()

        Dim LoopC As Long
    
100     QueryBuilder.Append "UPDATE user SET "
102     QueryBuilder.Append "name = ?, "
104     QueryBuilder.Append "level = ?, "
106     QueryBuilder.Append "exp = ?, "
108     QueryBuilder.Append "genre_id = ?, "
110     QueryBuilder.Append "race_id = ?, "
112     QueryBuilder.Append "class_id = ?, "
114     QueryBuilder.Append "home_id = ?, "
116     QueryBuilder.Append "description = ?, "
118     QueryBuilder.Append "gold = ?, "
120     QueryBuilder.Append "bank_gold = ?, "
122     QueryBuilder.Append "free_skillpoints = ?, "
124     QueryBuilder.Append "pets_saved = ?, "
126     QueryBuilder.Append "pos_map = ?, "
128     QueryBuilder.Append "pos_x = ?, "
130     QueryBuilder.Append "pos_y = ?, "
134     QueryBuilder.Append "message_info = ?, "
136     QueryBuilder.Append "body_id = ?, "
138     QueryBuilder.Append "head_id = ?, "
140     QueryBuilder.Append "weapon_id = ?, "
142     QueryBuilder.Append "helmet_id = ?, "
144     QueryBuilder.Append "shield_id = ?, "
146     QueryBuilder.Append "heading = ?, "
150     QueryBuilder.Append "slot_armour = ?, "
152     QueryBuilder.Append "slot_weapon = ?, "
154     QueryBuilder.Append "slot_shield = ?, "
156     QueryBuilder.Append "slot_helmet = ?, "
158     QueryBuilder.Append "slot_ammo = ?, "
160     QueryBuilder.Append "slot_dm = ?, "
162     QueryBuilder.Append "slot_rm = ?, "
164     QueryBuilder.Append "slot_tool = ?, "
166     QueryBuilder.Append "slot_magic = ?, "
168     QueryBuilder.Append "slot_knuckles = ?, "
170     QueryBuilder.Append "slot_ship = ?, "
172     QueryBuilder.Append "slot_mount = ?, "
174     QueryBuilder.Append "min_hp = ?, "
176     QueryBuilder.Append "max_hp = ?, "
178     QueryBuilder.Append "min_man = ?, "
180     QueryBuilder.Append "max_man = ?, "
182     QueryBuilder.Append "min_sta = ?, "
184     QueryBuilder.Append "max_sta = ?, "
186     QueryBuilder.Append "min_ham = ?, "
188     QueryBuilder.Append "max_ham = ?, "
190     QueryBuilder.Append "min_sed = ?, "
192     QueryBuilder.Append "max_sed = ?, "
194     QueryBuilder.Append "min_hit = ?, "
196     QueryBuilder.Append "max_hit = ?, "
198     QueryBuilder.Append "killed_npcs = ?, "
200     QueryBuilder.Append "killed_users = ?, "
202     QueryBuilder.Append "invent_level = ?, "
203     QueryBuilder.Append "elo = ?, "
206     QueryBuilder.Append "is_naked = ?, "
207     QueryBuilder.Append "is_poisoned = ?, "
208     QueryBuilder.Append "is_incinerated = ?, "
214     QueryBuilder.Append "is_dead = ?, "
216     QueryBuilder.Append "is_sailing = ?, "
218     QueryBuilder.Append "is_paralyzed = ?, "
220     QueryBuilder.Append "is_mounted = ?, "
222     QueryBuilder.Append "is_silenced = ?, "
224     QueryBuilder.Append "silence_minutes_left = ?, "
226     QueryBuilder.Append "silence_elapsed_seconds = ?, "
228     QueryBuilder.Append "spouse = ?, "
230     QueryBuilder.Append "counter_pena = ?, "
232     QueryBuilder.Append "deaths = ?, "
234     QueryBuilder.Append "pertenece_consejo_real = ?, "
236     QueryBuilder.Append "pertenece_consejo_caos = ?, "
238     QueryBuilder.Append "pertenece_real = ?, "
240     QueryBuilder.Append "pertenece_caos = ?, "
242     QueryBuilder.Append "ciudadanos_matados = ?, "
244     QueryBuilder.Append "criminales_matados = ?, "
246     QueryBuilder.Append "recibio_armadura_real = ?, "
248     QueryBuilder.Append "recibio_armadura_caos = ?, "
250     QueryBuilder.Append "recibio_exp_real = ?, "
252     QueryBuilder.Append "recibio_exp_caos = ?, "
254     QueryBuilder.Append "recompensas_real = ?, "
256     QueryBuilder.Append "recompensas_caos = ?, "
258     QueryBuilder.Append "reenlistadas = ?, "
260     QueryBuilder.Append "nivel_ingreso = ?, "
262     QueryBuilder.Append "matados_ingreso = ?, "
264     QueryBuilder.Append "siguiente_recompensa = ?, "
266     QueryBuilder.Append "status = ?, "
268     QueryBuilder.Append "guild_index = ?, "
270     QueryBuilder.Append "chat_combate = ?, "
272     QueryBuilder.Append "chat_global = ?, "
274     QueryBuilder.Append "is_logged = ?, "
276     QueryBuilder.Append "warnings = ?,"
        QueryBuilder.Append "return_map = ?,"
        QueryBuilder.Append "return_x = ?,"
        QueryBuilder.Append "return_y = ? "
278     QueryBuilder.Append " WHERE id = ?"
    
        ' Guardo la query ensamblada
280     QUERY_UPDATE_MAINPJ = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
282     Call QueryBuilder.Clear
    
        ' ************************** User bank inventory **************************************
284     QueryBuilder.Append "INSERT INTO bank_item (user_id, number, item_id, amount) VALUES "

286     For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
288         QueryBuilder.Append "(?, ?, ?, ?)"

290         If LoopC < MAX_BANCOINVENTORY_SLOTS Then
292             QueryBuilder.Append ", "

            End If

294     Next LoopC
        
296     QueryBuilder.Append " ON DUPLICATE KEY UPDATE item_id=VALUES(item_id), amount=VALUES(Amount); "
    
        ' Guardo la query ensamblada
298     QUERY_SAVE_BANCOINV = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
300     Call QueryBuilder.Clear
    
        ' ************************** UPSERT QUERIES **************************************
    
304     QUERY_UPSERT_SPELLS = QUERY_SAVE_SPELLS & " ON DUPLICATE KEY UPDATE spell_id=VALUES(spell_id); "
306     QUERY_UPSERT_INVENTORY = QUERY_SAVE_INVENTORY & " ON DUPLICATE KEY UPDATE item_id=VALUES(item_id), amount=VALUES(Amount), is_equipped=VALUES(is_equipped); "
308     QUERY_UPSERT_SKILLS = QUERY_SAVE_SKILLS & " ON DUPLICATE KEY UPDATE value=VALUES(value); "
310     QUERY_UPSERT_PETS = QUERY_SAVE_PETS & " ON DUPLICATE KEY UPDATE pet_id=VALUES(pet_id); "
    
End Sub
