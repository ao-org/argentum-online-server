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
Private QueryBuilder            As cStringBuilder

Public QUERY_LOAD_MAINPJ        As String

' DYNAMIC QUERIES
Public QUERY_SAVE_MAINPJ        As String
Public QUERY_SAVE_SPELLS        As String
Public QUERY_SAVE_INVENTORY     As String
Public QUERY_SAVE_BANCOINV      As String
Public QUERY_SAVE_SKILLS        As String
Public QUERY_SAVE_QUESTS        As String
Public QUERY_SAVE_PETS          As String

Public QUERY_UPDATE_MAINPJ      As String
Public QUERY_UPSERT_SPELLS      As String
Public QUERY_UPSERT_INVENTORY   As String
Public QUERY_UPSERT_SKILLS      As String
Public QUERY_UPSERT_PETS        As String

Public Sub Contruir_Querys()

        Call ConstruirQuery_CargarPersonaje
100     Call ConstruirQuery_CrearPersonaje
102     Call ConstruirQuery_GuardarPersonaje

End Sub

Private Sub ConstruirQuery_CargarPersonaje()
Dim LoopC                       As Long

10  On Error GoTo ConstruirQuery_CargarPersonaje_Error

20  Set QueryBuilder = New cStringBuilder

    ' ************************** Basic user data ********************************
30  QueryBuilder.Append "SELECT "
40  QueryBuilder.Append "account_id,"
50  QueryBuilder.Append "ID,"
60  QueryBuilder.Append "name,"
70  QueryBuilder.Append "level,"
80  QueryBuilder.Append "Exp,"
90  QueryBuilder.Append "genre_id,"
100 QueryBuilder.Append "race_id,"
110 QueryBuilder.Append "class_id,"
120 QueryBuilder.Append "home_id,"
130 QueryBuilder.Append "description,"
140 QueryBuilder.Append "gold,"
150 QueryBuilder.Append "bank_gold,"
160 QueryBuilder.Append "free_skillpoints,"
170 QueryBuilder.Append "pos_map,"
180 QueryBuilder.Append "pos_x,"
190 QueryBuilder.Append "pos_y,"
200 QueryBuilder.Append "message_info,"
210 QueryBuilder.Append "body_id,"
220 QueryBuilder.Append "head_id,"
230 QueryBuilder.Append "weapon_id,"
240 QueryBuilder.Append "helmet_id,"
250 QueryBuilder.Append "shield_id,"
260 QueryBuilder.Append "Heading,"
270 QueryBuilder.Append "max_hp,"
280 QueryBuilder.Append "min_hp,"
290 QueryBuilder.Append "min_man,"
300 QueryBuilder.Append "min_sta,"
310 QueryBuilder.Append "min_ham,"
320 QueryBuilder.Append "min_sed,"
330 QueryBuilder.Append "killed_npcs,"
340 QueryBuilder.Append "killed_users,"
350 QueryBuilder.Append "puntos_pesca,"
360 QueryBuilder.Append "ELO,"
370 QueryBuilder.Append "is_naked,"
380 QueryBuilder.Append "is_poisoned,"
390 QueryBuilder.Append "is_incinerated,"
400 QueryBuilder.Append "is_banned,"
410 QueryBuilder.Append "banned_by,"
420 QueryBuilder.Append "ban_reason,"
430 QueryBuilder.Append "is_dead,"
440 QueryBuilder.Append "is_sailing,"
450 QueryBuilder.Append "is_paralyzed,"
460 QueryBuilder.Append "deaths,"
470 QueryBuilder.Append "is_mounted,"
480 QueryBuilder.Append "spouse,"
490 QueryBuilder.Append "is_silenced,"
500 QueryBuilder.Append "silence_minutes_left,"
510 QueryBuilder.Append "silence_elapsed_seconds,"
520 QueryBuilder.Append "pets_saved,"
530 QueryBuilder.Append "return_map,"
540 QueryBuilder.Append "return_x,"
550 QueryBuilder.Append "return_y,"
560 QueryBuilder.Append "counter_pena,"
570 QueryBuilder.Append "chat_global,"
580 QueryBuilder.Append "chat_combate,"
590 QueryBuilder.Append "ciudadanos_matados,"
600 QueryBuilder.Append "criminales_matados,"
610 QueryBuilder.Append "recibio_armadura_real,"
620 QueryBuilder.Append "recibio_armadura_caos,"
630 QueryBuilder.Append "recompensas_real,"
640 QueryBuilder.Append "recompensas_caos,"
650 QueryBuilder.Append "faction_score,"
660 QueryBuilder.Append "Reenlistadas,"
670 QueryBuilder.Append "nivel_ingreso,"
680 QueryBuilder.Append "matados_ingreso,"
690 QueryBuilder.Append "Status,"
700 QueryBuilder.Append "Guild_Index,"
710 QueryBuilder.Append "guild_rejected_because,"
720 QueryBuilder.Append "warnings,"
730 QueryBuilder.Append "last_logout,"
740 QueryBuilder.Append "is_reset,"
750 QueryBuilder.Append "is_locked_in_mao,"
760 QueryBuilder.Append "jinete_level,"
770 QueryBuilder.Append "backpack_id"
780 QueryBuilder.Append " FROM user WHERE name= ?"

    ' Guardo la query ensamblada
790 QUERY_LOAD_MAINPJ = QueryBuilder.ToString

    ' Limpio el constructor de querys
800 Call QueryBuilder.Clear

810 On Error GoTo 0
820 Exit Sub

ConstruirQuery_CargarPersonaje_Error:

830 Call Logging.TraceError(Err.Number, Err.Description, "Database_Queries.ConstruirQuery_CargarPersonaje", Erl())

End Sub

Private Sub ConstruirQuery_CrearPersonaje()
Dim LoopC                       As Long

10  Set QueryBuilder = New cStringBuilder

    ' ************************** Basic user data ********************************
20  QueryBuilder.Append "INSERT INTO user ("
30  QueryBuilder.Append "name, "
40  QueryBuilder.Append "account_id, "
50  QueryBuilder.Append "level, "
60  QueryBuilder.Append "exp, "
70  QueryBuilder.Append "genre_id, "
80  QueryBuilder.Append "race_id, "
90  QueryBuilder.Append "class_id, "
100 QueryBuilder.Append "home_id, "
110 QueryBuilder.Append "description, "
120 QueryBuilder.Append "gold, "
130 QueryBuilder.Append "free_skillpoints, "
140 QueryBuilder.Append "pos_map, "
150 QueryBuilder.Append "pos_x, "
160 QueryBuilder.Append "pos_y, "
170 QueryBuilder.Append "body_id, "
180 QueryBuilder.Append "head_id, "
190 QueryBuilder.Append "weapon_id, "
200 QueryBuilder.Append "helmet_id, "
210 QueryBuilder.Append "shield_id, "
220 QueryBuilder.Append "max_hp, "
230 QueryBuilder.Append "min_hp, "
240 QueryBuilder.Append "min_man, "
250 QueryBuilder.Append "min_sta, "
260 QueryBuilder.Append "min_ham, "
270 QueryBuilder.Append "min_sed, "
280 QueryBuilder.Append "is_naked, "
290 QueryBuilder.Append "status) VALUES ( "
    Dim i                       As Long
300 For i = 0 To 25
310     QueryBuilder.Append "?,"
320 Next i
330 QueryBuilder.Append "?)"

    ' Guardo la query ensamblada
340 QUERY_SAVE_MAINPJ = QueryBuilder.ToString

    ' Limpio el constructor de querys
350 Call QueryBuilder.Clear

    ' ************************** User spells ************************************
360 QueryBuilder.Append "INSERT INTO spell (user_id, number, spell_id) VALUES "

370 For LoopC = 1 To MAXUSERHECHIZOS
380     QueryBuilder.Append "(?, ?, ?)"

390     If LoopC < MAXUSERHECHIZOS Then
400         QueryBuilder.Append ", "
410     End If

420 Next LoopC

    ' Guardo la query ensamblada
430 QUERY_SAVE_SPELLS = QueryBuilder.ToString

    ' Limpio el constructor de querys
440 Call QueryBuilder.Clear

    ' ******************* INVENTORY *******************
450 QueryBuilder.Append "INSERT INTO inventory_item (user_id, number, item_id, Amount, is_equipped, elemental_tags) VALUES "

460 For LoopC = 1 To MAX_INVENTORY_SLOTS
470     QueryBuilder.Append "(?, ?, ?, ?, ?, ?)"
480     If LoopC < MAX_INVENTORY_SLOTS Then
490         QueryBuilder.Append ", "
500     End If
510 Next LoopC

    ' Guardo la query ensamblada
520 QUERY_SAVE_INVENTORY = QueryBuilder.ToString

    ' Limpio el constructor de querys
530 Call QueryBuilder.Clear

    ' ************************** User skills ************************************
540 QueryBuilder.Append "INSERT INTO skillpoint (user_id, number, value) VALUES "

550 For LoopC = 1 To NUMSKILLS
560     QueryBuilder.Append "(?, ?, ?)"
570     If LoopC < NUMSKILLS Then
580         QueryBuilder.Append ", "
590     End If
600 Next LoopC

    ' Guardo la query ensamblada
610 QUERY_SAVE_SKILLS = QueryBuilder.ToString

    ' Limpio el constructor de querys
620 Call QueryBuilder.Clear

    ' ************************** User quests ************************************
630 QueryBuilder.Append "INSERT INTO quest (user_id, number) VALUES "

640 For LoopC = 1 To MAXUSERQUESTS
650     QueryBuilder.Append "(?, ?)"
660     If LoopC < MAXUSERQUESTS Then
670         QueryBuilder.Append ", "
680     End If
690 Next LoopC

    ' Guardo la query ensamblada
700 QUERY_SAVE_QUESTS = QueryBuilder.ToString

    ' Limpio el constructor de querys
710 Call QueryBuilder.Clear

    ' ************************** User pets **************************************
720 QueryBuilder.Append "INSERT INTO pet (user_id, number, pet_id) VALUES "

730 For LoopC = 1 To MAXMASCOTAS
740     QueryBuilder.Append "(?, ?, ?)"
750     If LoopC < MAXMASCOTAS Then
760         QueryBuilder.Append ", "
770     End If
780 Next LoopC

    ' Guardo la query ensamblada
790 QUERY_SAVE_PETS = QueryBuilder.ToString

    ' Limpio el constructor de querys
800 Call QueryBuilder.Clear

End Sub

Private Sub ConstruirQuery_GuardarPersonaje()

Dim LoopC                       As Long

10  QueryBuilder.Append "UPDATE user SET "
20  QueryBuilder.Append "name = ?, "
30  QueryBuilder.Append "level = ?, "
40  QueryBuilder.Append "exp = ?, "
50  QueryBuilder.Append "genre_id = ?, "
60  QueryBuilder.Append "race_id = ?, "
70  QueryBuilder.Append "class_id = ?, "
80  QueryBuilder.Append "home_id = ?, "
90  QueryBuilder.Append "description = ?, "
100 QueryBuilder.Append "gold = ?, "
110 QueryBuilder.Append "bank_gold = ?, "
120 QueryBuilder.Append "free_skillpoints = ?, "
130 QueryBuilder.Append "pets_saved = ?, "
140 QueryBuilder.Append "pos_map = ?, "
150 QueryBuilder.Append "pos_x = ?, "
160 QueryBuilder.Append "pos_y = ?, "
170 QueryBuilder.Append "message_info = ?, "
180 QueryBuilder.Append "body_id = ?, "
190 QueryBuilder.Append "head_id = ?, "
200 QueryBuilder.Append "weapon_id = ?, "
210 QueryBuilder.Append "helmet_id = ?, "
220 QueryBuilder.Append "shield_id = ?, "
230 QueryBuilder.Append "heading = ?, "
240 QueryBuilder.Append "max_hp = ?, "
250 QueryBuilder.Append "min_hp = ?, "
260 QueryBuilder.Append "min_man = ?, "
270 QueryBuilder.Append "min_sta = ?, "
280 QueryBuilder.Append "min_ham = ?, "
290 QueryBuilder.Append "min_sed = ?, "
300 QueryBuilder.Append "killed_npcs = ?, "
310 QueryBuilder.Append "killed_users = ?, "
320 QueryBuilder.Append "puntos_pesca = ?, "
330 QueryBuilder.Append "elo = ?, "
340 QueryBuilder.Append "is_naked = ?, "
350 QueryBuilder.Append "is_poisoned = ?, "
360 QueryBuilder.Append "is_incinerated = ?, "
370 QueryBuilder.Append "is_dead = ?, "
380 QueryBuilder.Append "is_sailing = ?, "
390 QueryBuilder.Append "is_paralyzed = ?, "
400 QueryBuilder.Append "is_mounted = ?, "
410 QueryBuilder.Append "is_silenced = ?, "
420 QueryBuilder.Append "silence_minutes_left = ?, "
430 QueryBuilder.Append "silence_elapsed_seconds = ?, "
440 QueryBuilder.Append "spouse = ?, "
450 QueryBuilder.Append "counter_pena = ?, "
460 QueryBuilder.Append "deaths = ?, "
470 QueryBuilder.Append "ciudadanos_matados = ?, "
480 QueryBuilder.Append "criminales_matados = ?, "
490 QueryBuilder.Append "recibio_armadura_real = ?, "
500 QueryBuilder.Append "recibio_armadura_caos = ?, "
510 QueryBuilder.Append "recompensas_real = ?, "
520 QueryBuilder.Append "faction_score = ?, "
530 QueryBuilder.Append "recompensas_caos = ?, "
540 QueryBuilder.Append "reenlistadas = ?, "
550 QueryBuilder.Append "nivel_ingreso = ?, "
560 QueryBuilder.Append "matados_ingreso = ?, "
570 QueryBuilder.Append "status = ?, "
580 QueryBuilder.Append "guild_index = ?, "
590 QueryBuilder.Append "chat_combate = ?, "
600 QueryBuilder.Append "chat_global = ?, "
610 QueryBuilder.Append "warnings = ?, "
620 QueryBuilder.Append "return_map = ?, "
630 QueryBuilder.Append "return_x = ?, "
640 QueryBuilder.Append "return_y = ?, "
650 QueryBuilder.Append "jinete_level = ?, "
660 QueryBuilder.Append "backpack_id = ?, "
670 QueryBuilder.Append "last_logout = strftime('%s','now') "
680 QueryBuilder.Append "WHERE id = ?"

    ' Guardo la query ensamblada
690 QUERY_UPDATE_MAINPJ = QueryBuilder.ToString

    ' Limpio el constructor de querys
700 Call QueryBuilder.Clear

    ' ************************** User bank inventory **************************************
710 QueryBuilder.Append "REPLACE INTO bank_item (user_id, number, item_id, amount, elemental_tags) VALUES "

720 For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
730     QueryBuilder.Append "(?, ?, ?, ?, ?)"
740     If LoopC < MAX_BANCOINVENTORY_SLOTS Then
750         QueryBuilder.Append ", "
760     End If
770 Next LoopC

    ' Guardo la query ensamblada
780 QUERY_SAVE_BANCOINV = QueryBuilder.ToString

    ' Limpio el constructor de querys
790 Call QueryBuilder.Clear

    ' ************************** User spells ************************************
800 QueryBuilder.Append "REPLACE INTO spell (user_id, number, spell_id) VALUES "

810 For LoopC = 1 To MAXUSERHECHIZOS
820     QueryBuilder.Append "(?, ?, ?)"
830     If LoopC < MAXUSERHECHIZOS Then
840         QueryBuilder.Append ", "
850     End If
860 Next LoopC

    ' Guardo la query ensamblada
870 QUERY_UPSERT_SPELLS = QueryBuilder.ToString

    ' Limpio el constructor de querys
880 Call QueryBuilder.Clear

    ' ******************* INVENTORY *******************
890 QueryBuilder.Append "REPLACE INTO inventory_item (user_id, number, item_id, Amount, is_equipped, elemental_tags) VALUES "

900 For LoopC = 1 To MAX_INVENTORY_SLOTS
910     QueryBuilder.Append "(?, ?, ?, ?, ?, ?)"
920     If LoopC < MAX_INVENTORY_SLOTS Then
930         QueryBuilder.Append ", "
940     End If
950 Next LoopC

    ' Guardo la query ensamblada
960 QUERY_UPSERT_INVENTORY = QueryBuilder.ToString

    ' Limpio el constructor de querys
970 Call QueryBuilder.Clear

    ' ************************** User skills ************************************
980 QueryBuilder.Append "REPLACE INTO skillpoint (user_id, number, value) VALUES "

990 For LoopC = 1 To NUMSKILLS
1000    QueryBuilder.Append "(?, ?, ?)"
1010    If LoopC < NUMSKILLS Then
1020        QueryBuilder.Append ", "
1030    End If
1040 Next LoopC

    ' Guardo la query ensamblada
1050 QUERY_UPSERT_SKILLS = QueryBuilder.ToString

    ' Limpio el constructor de querys
1060 Call QueryBuilder.Clear

    ' ************************** User pets **************************************
1070 QueryBuilder.Append "REPLACE INTO pet (user_id, number, pet_id) VALUES "

1080 For LoopC = 1 To MAXMASCOTAS
1090    QueryBuilder.Append "(?, ?, ?)"
1100    If LoopC < MAXMASCOTAS Then
1110        QueryBuilder.Append ", "
1120    End If

1130 Next LoopC

    ' Guardo la query ensamblada
1140 QUERY_UPSERT_PETS = QueryBuilder.ToString

    ' Limpio el constructor de querys
1150 Call QueryBuilder.Clear

End Sub
