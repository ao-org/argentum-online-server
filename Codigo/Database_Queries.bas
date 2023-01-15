Attribute VB_Name = "Database_Queries"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit

'Constructor de queries.
'Me permite concatenar strings MUCHO MAS rapido
Private QueryBuilder As cStringBuilder

Public QUERY_LOAD_MAINPJ As String

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


Public Sub Contruir_Querys()
        Call ConstruirQuery_CargarPersonaje
100     Call ConstruirQuery_CrearPersonaje
102     Call ConstruirQuery_GuardarPersonaje
End Sub

Private Sub ConstruirQuery_CargarPersonaje()
        Dim LoopC As Long
    
100     Set QueryBuilder = New cStringBuilder
    
        ' ************************** Basic user data ********************************
102     QueryBuilder.Append "SELECT "
        QueryBuilder.Append "account_id,"
        QueryBuilder.Append "ID,"
        QueryBuilder.Append "name,"
        QueryBuilder.Append "level,"
        QueryBuilder.Append "Exp,"
        QueryBuilder.Append "genre_id,"
        QueryBuilder.Append "race_id,"
        QueryBuilder.Append "class_id,"
        QueryBuilder.Append "home_id,"
        QueryBuilder.Append "description,"
        QueryBuilder.Append "gold,"
        QueryBuilder.Append "bank_gold,"
        QueryBuilder.Append "free_skillpoints,"
        QueryBuilder.Append "pos_map,"
        QueryBuilder.Append "pos_x,"
        QueryBuilder.Append "pos_y,"
        QueryBuilder.Append "message_info,"
        QueryBuilder.Append "body_id,"
        QueryBuilder.Append "head_id,"
        QueryBuilder.Append "weapon_id,"
        QueryBuilder.Append "helmet_id,"
        QueryBuilder.Append "shield_id,"
        QueryBuilder.Append "Heading,"
        QueryBuilder.Append "slot_armour,"
        QueryBuilder.Append "slot_weapon,"
        QueryBuilder.Append "slot_helmet,"
        QueryBuilder.Append "slot_shield,"
        QueryBuilder.Append "slot_ammo,"
        QueryBuilder.Append "slot_ship,"
        QueryBuilder.Append "slot_mount,"
        QueryBuilder.Append "slot_dm,"
        QueryBuilder.Append "slot_rm,"
        QueryBuilder.Append "slot_knuckles,"
        QueryBuilder.Append "slot_tool,"
        QueryBuilder.Append "slot_magic,"
        QueryBuilder.Append "min_hp,"
        QueryBuilder.Append "max_hp,"
        QueryBuilder.Append "min_man,"
        QueryBuilder.Append "max_man,"
        QueryBuilder.Append "min_sta,"
        QueryBuilder.Append "max_sta,"
        QueryBuilder.Append "min_ham,"
        QueryBuilder.Append "max_ham,"
        QueryBuilder.Append "min_sed,"
        QueryBuilder.Append "max_sed,"
        QueryBuilder.Append "min_hit,"
        QueryBuilder.Append "max_hit,"
        QueryBuilder.Append "killed_npcs,"
        QueryBuilder.Append "killed_users,"
        QueryBuilder.Append "puntos_pesca,"
        QueryBuilder.Append "invent_level,"
        QueryBuilder.Append "ELO,"
        QueryBuilder.Append "is_naked,"
        QueryBuilder.Append "is_poisoned,"
        QueryBuilder.Append "is_incinerated,"
        QueryBuilder.Append "is_banned,"
        QueryBuilder.Append "banned_by,"
        QueryBuilder.Append "ban_reason,"
        QueryBuilder.Append "is_dead,"
        QueryBuilder.Append "is_sailing,"
        QueryBuilder.Append "is_paralyzed,"
        QueryBuilder.Append "deaths,"
        QueryBuilder.Append "is_mounted,"
        QueryBuilder.Append "spouse,"
        QueryBuilder.Append "is_silenced,"
        QueryBuilder.Append "silence_minutes_left,"
        QueryBuilder.Append "silence_elapsed_seconds,"
        QueryBuilder.Append "pets_saved,"
        QueryBuilder.Append "return_map,"
        QueryBuilder.Append "return_x,"
        QueryBuilder.Append "return_y,"
        QueryBuilder.Append "counter_pena,"
        QueryBuilder.Append "chat_global,"
        QueryBuilder.Append "chat_combate,"
        QueryBuilder.Append "ciudadanos_matados,"
        QueryBuilder.Append "criminales_matados,"
        QueryBuilder.Append "recibio_armadura_real,"
        QueryBuilder.Append "recibio_armadura_caos,"
        QueryBuilder.Append "recibio_exp_real,"
        QueryBuilder.Append "recibio_exp_caos,"
        QueryBuilder.Append "recompensas_real,"
        QueryBuilder.Append "recompensas_caos,"
        QueryBuilder.Append "Reenlistadas,"
        QueryBuilder.Append "nivel_ingreso,"
        QueryBuilder.Append "matados_ingreso,"
        QueryBuilder.Append "siguiente_recompensa,"
        QueryBuilder.Append "Status,"
        QueryBuilder.Append "Guild_Index,"
        QueryBuilder.Append "guild_rejected_because,"
        QueryBuilder.Append "warnings,"
        QueryBuilder.Append "last_logout,"
        QueryBuilder.Append "credits,"
        QueryBuilder.Append "is_reset,"
        QueryBuilder.Append "quest_belthor,"
        QueryBuilder.Append "is_locked_in_mao"
        'QueryBuilder.Append ",DATE_FORMAT(fecha_ingreso, '%Y-%m-%d') as 'fecha_ingreso_format'"
        QueryBuilder.Append " FROM user WHERE name= ?"
    
        ' Guardo la query ensamblada
198     QUERY_LOAD_MAINPJ = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
200     Call QueryBuilder.Clear

End Sub

Private Sub ConstruirQuery_CrearPersonaje()
        Dim LoopC As Long
    
100     Set QueryBuilder = New cStringBuilder
    
        ' ************************** Basic user data ********************************
102     QueryBuilder.Append "INSERT INTO user ("
104     QueryBuilder.Append "name, "
106     QueryBuilder.Append "account_id, "
108     QueryBuilder.Append "level, "
110     QueryBuilder.Append "exp, "
112     QueryBuilder.Append "genre_id, "
114     QueryBuilder.Append "race_id, "
116     QueryBuilder.Append "class_id, "
118     QueryBuilder.Append "home_id, "
120     QueryBuilder.Append "description, "
122     QueryBuilder.Append "gold, "
124     QueryBuilder.Append "free_skillpoints, "
126     QueryBuilder.Append "pos_map, "
128     QueryBuilder.Append "pos_x, "
130     QueryBuilder.Append "pos_y, "
132     QueryBuilder.Append "body_id, "
134     QueryBuilder.Append "head_id, "
136     QueryBuilder.Append "weapon_id, "
138     QueryBuilder.Append "helmet_id, "
140     QueryBuilder.Append "shield_id, "
144     QueryBuilder.Append "slot_armour, "
146     QueryBuilder.Append "slot_weapon, "
148     QueryBuilder.Append "slot_shield, "
150     QueryBuilder.Append "slot_helmet, "
152     QueryBuilder.Append "slot_ammo, "
154     QueryBuilder.Append "slot_dm, "
156     QueryBuilder.Append "slot_rm, "
158     QueryBuilder.Append "slot_tool, "
160     QueryBuilder.Append "slot_magic, "
162     QueryBuilder.Append "slot_knuckles, "
164     QueryBuilder.Append "slot_ship, "
166     QueryBuilder.Append "slot_mount, "
168     QueryBuilder.Append "min_hp, "
170     QueryBuilder.Append "max_hp, "
172     QueryBuilder.Append "min_man, "
174     QueryBuilder.Append "max_man, "
176     QueryBuilder.Append "min_sta, "
178     QueryBuilder.Append "max_sta, "
180     QueryBuilder.Append "min_ham, "
182     QueryBuilder.Append "max_ham, "
184     QueryBuilder.Append "min_sed, "
186     QueryBuilder.Append "max_sed, "
188     QueryBuilder.Append "min_hit, "
190     QueryBuilder.Append "max_hit, "
192     QueryBuilder.Append "is_naked, "
'193     QueryBuilder.Append "status, "
'195     QueryBuilder.Append "is_reset, "
'194     QueryBuilder.Append "quest_belthor) VALUES ("
194     QueryBuilder.Append "status) VALUES ("
        Dim i As Long
        For i = 0 To 43
            QueryBuilder.Append "?,"
        Next i
        QueryBuilder.Append "?)"

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
201     QueryBuilder.Append "puntos_pesca = ?, "
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
276     QueryBuilder.Append "warnings = ?,"
        QueryBuilder.Append "return_map = ?,"
        QueryBuilder.Append "return_x = ?,"
        QueryBuilder.Append "return_y = ?, "
        QueryBuilder.Append "last_logout = strftime('%s','now') "
278     QueryBuilder.Append "WHERE id = ?"
    
        ' Guardo la query ensamblada
280     QUERY_UPDATE_MAINPJ = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
282     Call QueryBuilder.Clear
    
        ' ************************** User bank inventory **************************************
284     QueryBuilder.Append "REPLACE INTO bank_item (user_id, number, item_id, amount) VALUES "

286     For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
288         QueryBuilder.Append "(?, ?, ?, ?)"

290         If LoopC < MAX_BANCOINVENTORY_SLOTS Then
292             QueryBuilder.Append ", "

            End If

294     Next LoopC

        ' Guardo la query ensamblada
298     QUERY_SAVE_BANCOINV = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
300     Call QueryBuilder.Clear
    
        ' ************************** User spells ************************************
        QueryBuilder.Append "REPLACE INTO spell (user_id, number, spell_id) VALUES "

        For LoopC = 1 To MAXUSERHECHIZOS
            QueryBuilder.Append "(?, ?, ?)"

            If LoopC < MAXUSERHECHIZOS Then
                QueryBuilder.Append ", "

            End If

        Next LoopC

        ' Guardo la query ensamblada
        QUERY_UPSERT_SPELLS = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
        Call QueryBuilder.Clear
    
        ' ******************* INVENTORY *******************
        QueryBuilder.Append "REPLACE INTO inventory_item (user_id, number, item_id, Amount, is_equipped) VALUES "

        For LoopC = 1 To MAX_INVENTORY_SLOTS
            QueryBuilder.Append "(?, ?, ?, ?, ?)"

            If LoopC < MAX_INVENTORY_SLOTS Then
                QueryBuilder.Append ", "

            End If

        Next LoopC

        ' Guardo la query ensamblada
        QUERY_UPSERT_INVENTORY = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
        Call QueryBuilder.Clear
   
        ' ************************** User skills ************************************
        QueryBuilder.Append "REPLACE INTO skillpoint (user_id, number, value) VALUES "

        For LoopC = 1 To NUMSKILLS
            QueryBuilder.Append "(?, ?, ?)"

            If LoopC < NUMSKILLS Then
                QueryBuilder.Append ", "

            End If

        Next LoopC

        ' Guardo la query ensamblada
        QUERY_UPSERT_SKILLS = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
        Call QueryBuilder.Clear

        ' ************************** User pets **************************************
        QueryBuilder.Append "REPLACE INTO pet (user_id, number, pet_id) VALUES "

        For LoopC = 1 To MAXMASCOTAS
            QueryBuilder.Append "(?, ?, ?)"

            If LoopC < MAXMASCOTAS Then
                QueryBuilder.Append ", "

            End If

        Next LoopC

        ' Guardo la query ensamblada
        QUERY_UPSERT_PETS = QueryBuilder.ToString
    
        ' Limpio el constructor de querys
        Call QueryBuilder.Clear

End Sub
