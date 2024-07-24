Attribute VB_Name = "Database_Queries"
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
        QueryBuilder.Append "max_hp,"
        QueryBuilder.Append "min_hp,"
        QueryBuilder.Append "min_man,"
        QueryBuilder.Append "min_sta,"
        QueryBuilder.Append "min_ham,"
        QueryBuilder.Append "min_sed,"
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
        QueryBuilder.Append "recompensas_real,"
        QueryBuilder.Append "recompensas_caos,"
        QueryBuilder.Append "faction_score,"
        QueryBuilder.Append "Reenlistadas,"
        QueryBuilder.Append "nivel_ingreso,"
        QueryBuilder.Append "matados_ingreso,"
        QueryBuilder.Append "Status,"
        QueryBuilder.Append "Guild_Index,"
        QueryBuilder.Append "guild_rejected_because,"
        QueryBuilder.Append "warnings,"
        QueryBuilder.Append "last_logout,"
        QueryBuilder.Append "is_reset,"
        QueryBuilder.Append "is_locked_in_mao,"
        QueryBuilder.Append "user_key"
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
167     QueryBuilder.Append "max_hp, "
168     QueryBuilder.Append "min_hp, "
172     QueryBuilder.Append "min_man, "
176     QueryBuilder.Append "min_sta, "
180     QueryBuilder.Append "min_ham, "
184     QueryBuilder.Append "min_sed, "
192     QueryBuilder.Append "is_naked, "
194     QueryBuilder.Append "status, "
195     QueryBuilder.Append "user_key) VALUES ("
        Dim i As Long
        For i = 0 To 26
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
173     QueryBuilder.Append "max_hp = ?, "
174     QueryBuilder.Append "min_hp = ?, "
178     QueryBuilder.Append "min_man = ?, "
182     QueryBuilder.Append "min_sta = ?, "
186     QueryBuilder.Append "min_ham = ?, "
190     QueryBuilder.Append "min_sed = ?, "
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
254     QueryBuilder.Append "recompensas_real = ?, "
255     QueryBuilder.Append "faction_score = ?, "
256     QueryBuilder.Append "recompensas_caos = ?, "
258     QueryBuilder.Append "reenlistadas = ?, "
260     QueryBuilder.Append "nivel_ingreso = ?, "
262     QueryBuilder.Append "matados_ingreso = ?, "
266     QueryBuilder.Append "status = ?, "
268     QueryBuilder.Append "guild_index = ?, "
270     QueryBuilder.Append "chat_combate = ?, "
272     QueryBuilder.Append "chat_global = ?, "
276     QueryBuilder.Append "warnings = ?,"
        QueryBuilder.Append "return_map = ?,"
        QueryBuilder.Append "return_x = ?,"
        QueryBuilder.Append "return_y = ?, "
        QueryBuilder.Append "last_logout = strftime('%s','now'), "
        QueryBuilder.Append "user_key = ? "
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
