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
Private QueryBuilder          As cStringBuilder
Public QUERY_LOAD_MAINPJ      As String
' DYNAMIC QUERIES
Public QUERY_SAVE_MAINPJ      As String
Public QUERY_SAVE_SPELLS      As String
Public QUERY_SAVE_INVENTORY   As String
Public QUERY_SAVE_BANCOINV    As String
Public QUERY_SAVE_SKILLS      As String
Public QUERY_SAVE_QUESTS      As String
Public QUERY_SAVE_PETS        As String
Public QUERY_UPDATE_MAINPJ    As String
Public QUERY_UPSERT_SPELLS    As String
Public QUERY_UPSERT_INVENTORY As String
Public QUERY_UPSERT_SKILLS    As String
Public QUERY_UPSERT_PETS      As String

Public Sub Contruir_Querys()
    Call ConstruirQuery_CargarPersonaje
    Call ConstruirQuery_CrearPersonaje
    Call ConstruirQuery_GuardarPersonaje
End Sub

Private Sub ConstruirQuery_CargarPersonaje()
    Dim LoopC As Long
    Set QueryBuilder = New cStringBuilder
    ' ************************** Basic user data ********************************
    QueryBuilder.Append "SELECT "
    QueryBuilder.Append "account_id,"
    QueryBuilder.Append "ID,"
    QueryBuilder.Append "name,"
    QueryBuilder.Append "alias,"
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
    QueryBuilder.Append "is_reset,"
    QueryBuilder.Append "is_locked_in_mao,"
    QueryBuilder.Append "jinete_level,"
    QueryBuilder.Append "backpack_id"
    QueryBuilder.Append " FROM user WHERE name= ?"
    ' Guardo la query ensamblada
    QUERY_LOAD_MAINPJ = QueryBuilder.ToString
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
End Sub

Private Sub ConstruirQuery_CrearPersonaje()
    Dim LoopC As Long
    Set QueryBuilder = New cStringBuilder
    ' ************************** Basic user data ********************************
    QueryBuilder.Append "INSERT INTO user ("
    QueryBuilder.Append "name, "
    QueryBuilder.Append "account_id, "
    QueryBuilder.Append "level, "
    QueryBuilder.Append "exp, "
    QueryBuilder.Append "genre_id, "
    QueryBuilder.Append "race_id, "
    QueryBuilder.Append "class_id, "
    QueryBuilder.Append "home_id, "
    QueryBuilder.Append "description, "
    QueryBuilder.Append "gold, "
    QueryBuilder.Append "free_skillpoints, "
    QueryBuilder.Append "pos_map, "
    QueryBuilder.Append "pos_x, "
    QueryBuilder.Append "pos_y, "
    QueryBuilder.Append "body_id, "
    QueryBuilder.Append "head_id, "
    QueryBuilder.Append "weapon_id, "
    QueryBuilder.Append "helmet_id, "
    QueryBuilder.Append "shield_id, "
    QueryBuilder.Append "max_hp, "
    QueryBuilder.Append "min_hp, "
    QueryBuilder.Append "min_man, "
    QueryBuilder.Append "min_sta, "
    QueryBuilder.Append "min_ham, "
    QueryBuilder.Append "min_sed, "
    QueryBuilder.Append "is_naked, "
    QueryBuilder.Append "status) VALUES ( "
    Dim i As Long
    For i = 0 To 25
        QueryBuilder.Append "?,"
    Next i
    QueryBuilder.Append "?)"
    ' Guardo la query ensamblada
    QUERY_SAVE_MAINPJ = QueryBuilder.ToString
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    ' ************************** User spells ************************************
    QueryBuilder.Append "INSERT INTO spell (user_id, number, spell_id) VALUES "
    For LoopC = 1 To MAXUSERHECHIZOS
        QueryBuilder.Append "(?, ?, ?)"
        If LoopC < MAXUSERHECHIZOS Then
            QueryBuilder.Append ", "
        End If
    Next LoopC
    ' Guardo la query ensamblada
    QUERY_SAVE_SPELLS = QueryBuilder.ToString
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    ' ******************* INVENTORY *******************
    QueryBuilder.Append "INSERT INTO inventory_item (user_id, number, item_id, Amount, is_equipped, elemental_tags) VALUES "
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        QueryBuilder.Append "(?, ?, ?, ?, ?, ?)"
        If LoopC < MAX_INVENTORY_SLOTS Then
            QueryBuilder.Append ", "
        End If
    Next LoopC
    ' Guardo la query ensamblada
    QUERY_SAVE_INVENTORY = QueryBuilder.ToString
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    ' ************************** User skills ************************************
    QueryBuilder.Append "INSERT INTO skillpoint (user_id, number, value) VALUES "
    For LoopC = 1 To NUMSKILLS
        QueryBuilder.Append "(?, ?, ?)"
        If LoopC < NUMSKILLS Then
            QueryBuilder.Append ", "
        End If
    Next LoopC
    ' Guardo la query ensamblada
    QUERY_SAVE_SKILLS = QueryBuilder.ToString
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    ' ************************** User quests ************************************
    QueryBuilder.Append "INSERT INTO quest (user_id, number) VALUES "
    For LoopC = 1 To MAXUSERQUESTS
        QueryBuilder.Append "(?, ?)"
        If LoopC < MAXUSERQUESTS Then
            QueryBuilder.Append ", "
        End If
    Next LoopC
    ' Guardo la query ensamblada
    QUERY_SAVE_QUESTS = QueryBuilder.ToString
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    ' ************************** User pets **************************************
    QueryBuilder.Append "INSERT INTO pet (user_id, number, pet_id) VALUES "
    For LoopC = 1 To MAXMASCOTAS
        QueryBuilder.Append "(?, ?, ?)"
        If LoopC < MAXMASCOTAS Then
            QueryBuilder.Append ", "
        End If
    Next LoopC
    ' Guardo la query ensamblada
    QUERY_SAVE_PETS = QueryBuilder.ToString
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
End Sub

Private Sub ConstruirQuery_GuardarPersonaje()
    Dim LoopC As Long
    QueryBuilder.Append "UPDATE user SET "
    QueryBuilder.Append "name = ?, "
    QueryBuilder.Append "level = ?, "
    QueryBuilder.Append "exp = ?, "
    QueryBuilder.Append "genre_id = ?, "
    QueryBuilder.Append "race_id = ?, "
    QueryBuilder.Append "class_id = ?, "
    QueryBuilder.Append "home_id = ?, "
    QueryBuilder.Append "description = ?, "
    QueryBuilder.Append "gold = ?, "
    QueryBuilder.Append "bank_gold = ?, "
    QueryBuilder.Append "free_skillpoints = ?, "
    QueryBuilder.Append "pets_saved = ?, "
    QueryBuilder.Append "pos_map = ?, "
    QueryBuilder.Append "pos_x = ?, "
    QueryBuilder.Append "pos_y = ?, "
    QueryBuilder.Append "message_info = ?, "
    QueryBuilder.Append "body_id = ?, "
    QueryBuilder.Append "head_id = ?, "
    QueryBuilder.Append "weapon_id = ?, "
    QueryBuilder.Append "helmet_id = ?, "
    QueryBuilder.Append "shield_id = ?, "
    QueryBuilder.Append "heading = ?, "
    QueryBuilder.Append "max_hp = ?, "
    QueryBuilder.Append "min_hp = ?, "
    QueryBuilder.Append "min_man = ?, "
    QueryBuilder.Append "min_sta = ?, "
    QueryBuilder.Append "min_ham = ?, "
    QueryBuilder.Append "min_sed = ?, "
    QueryBuilder.Append "killed_npcs = ?, "
    QueryBuilder.Append "killed_users = ?, "
    QueryBuilder.Append "puntos_pesca = ?, "
    QueryBuilder.Append "elo = ?, "
    QueryBuilder.Append "is_naked = ?, "
    QueryBuilder.Append "is_poisoned = ?, "
    QueryBuilder.Append "is_incinerated = ?, "
    QueryBuilder.Append "is_dead = ?, "
    QueryBuilder.Append "is_sailing = ?, "
    QueryBuilder.Append "is_paralyzed = ?, "
    QueryBuilder.Append "is_mounted = ?, "
    QueryBuilder.Append "is_silenced = ?, "
    QueryBuilder.Append "silence_minutes_left = ?, "
    QueryBuilder.Append "silence_elapsed_seconds = ?, "
    QueryBuilder.Append "spouse = ?, "
    QueryBuilder.Append "counter_pena = ?, "
    QueryBuilder.Append "deaths = ?, "
    QueryBuilder.Append "ciudadanos_matados = ?, "
    QueryBuilder.Append "criminales_matados = ?, "
    QueryBuilder.Append "recibio_armadura_real = ?, "
    QueryBuilder.Append "recibio_armadura_caos = ?, "
    QueryBuilder.Append "recompensas_real = ?, "
    QueryBuilder.Append "faction_score = ?, "
    QueryBuilder.Append "recompensas_caos = ?, "
    QueryBuilder.Append "reenlistadas = ?, "
    QueryBuilder.Append "nivel_ingreso = ?, "
    QueryBuilder.Append "matados_ingreso = ?, "
    QueryBuilder.Append "status = ?, "
    QueryBuilder.Append "guild_index = ?, "
    QueryBuilder.Append "chat_combate = ?, "
    QueryBuilder.Append "chat_global = ?, "
    QueryBuilder.Append "warnings = ?, "
    QueryBuilder.Append "return_map = ?, "
    QueryBuilder.Append "return_x = ?, "
    QueryBuilder.Append "return_y = ?, "
    QueryBuilder.Append "jinete_level = ?, "
    QueryBuilder.Append "backpack_id = ? "
    QueryBuilder.Append "WHERE id = ?"
    ' Guardo la query ensamblada
    QUERY_UPDATE_MAINPJ = QueryBuilder.ToString
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    ' ************************** User bank inventory **************************************
    QueryBuilder.Append "REPLACE INTO bank_item (user_id, number, item_id, amount, elemental_tags) VALUES "
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
        QueryBuilder.Append "(?, ?, ?, ?, ?)"
        If LoopC < MAX_BANCOINVENTORY_SLOTS Then
            QueryBuilder.Append ", "
        End If
    Next LoopC
    ' Guardo la query ensamblada
    QUERY_SAVE_BANCOINV = QueryBuilder.ToString
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
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
    QueryBuilder.Append "REPLACE INTO inventory_item (user_id, number, item_id, Amount, is_equipped, elemental_tags) VALUES "
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        QueryBuilder.Append "(?, ?, ?, ?, ?, ?)"
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

Function Exists(ByRef sTable As String, ByRef sField As String, ByRef sValue As String, _
                Optional ByRef sExtraField = vbNullString, Optional ByRef sExtraValue = vbNullString) As Boolean

Dim RS                          As ADODB.Recordset
Dim SQL                         As String

   On Error GoTo Exists_Error

    If sExtraField <> vbNullString And sExtraValue <> vbNullString Then
        SQL = "SELECT " & sField & " FROM " & sTable & " WHERE " & sField & " = ? AND " & sExtraField & " = ?"
        Set RS = Query(SQL, sValue, sExtraValue)
    Else
        SQL = "SELECT " & sField & " FROM " & sTable & " WHERE " & sField & " = ?"
        Set RS = Query(SQL, sValue)
    End If

    If RS Is Nothing Then
        Exists = False
    Else
        Exists = Not RS.EOF
        RS.Close
    End If

    Set RS = Nothing

   On Error GoTo 0
   Exit Function

Exists_Error:
    Exists = False
    Set RS = Nothing
    Call Logging.TraceError(Err.Number, Err.Description, "Database_Queries.Exists", Erl())

End Function
