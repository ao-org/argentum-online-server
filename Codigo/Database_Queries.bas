Attribute VB_Name = "Database_Queries"
Option Explicit

'Constructor de queries.
'Me permite concatenar strings MUCHO MAS rapido
Private QueryBuilder As cStringBuilder

' DYNAMIC QUERIES
Public QUERY_SAVE_MAINPJ As String
Public QUERY_SAVE_ATTRIBUTES As String
Public QUERY_SAVE_SPELLS As String
Public QUERY_SAVE_INVENTORY As String
Public QUERY_SAVE_BANCOINV As String
Public QUERY_SAVE_SKILLS As String
Public QUERY_SAVE_QUESTS As String
Public QUERY_SAVE_PETS As String

Public QUERY_UPDATE_MAINPJ As String
Public QUERY_UPSERT_ATTRIBUTES As String
Public QUERY_UPSERT_SPELLS As String
Public QUERY_UPSERT_INVENTORY As String
Public QUERY_UPSERT_SKILLS As String
Public QUERY_UPSERT_PETS As String

' CONSTANT QUERIES
Public Const QUERY_SAVE_CONNECTION As String = "INSERT INTO connection (user_id, ip) VALUES (? , ?) ON DUPLICATE KEY UPDATE date_last_login = VALUES(date_last_login);"
Public Const QUERY_DELETE_LAST_CONNECTIONS As String = "DELETE FROM connection WHERE user_id = ? AND date_last_login < (SELECT min(date_last_login) FROM (SELECT date_last_login FROM connection WHERE user_id = ? ORDER BY date_last_login DESC LIMIT 5) AS d);"

Public Sub Contruir_Querys()
    Call ConstruirQuery_CrearPersonaje
    Call ConstruirQuery_GuardarPersonaje
End Sub

Private Sub ConstruirQuery_CrearPersonaje()
    Dim LoopC As Long
    
    Set QueryBuilder = New cStringBuilder
    
    ' ************************** Basic user data ********************************
    QueryBuilder.Append "INSERT INTO user SET "
    QueryBuilder.Append "name = ?, "
    QueryBuilder.Append "account_id = ?, "
    QueryBuilder.Append "level = ?, "
    QueryBuilder.Append "exp = ?, "
    QueryBuilder.Append "genre_id = ?, "
    QueryBuilder.Append "race_id = ?, "
    QueryBuilder.Append "class_id = ?, "
    QueryBuilder.Append "home_id = ?, "
    QueryBuilder.Append "description = ?, "
    QueryBuilder.Append "gold = ?, "
    QueryBuilder.Append "free_skillpoints = ?, "
    QueryBuilder.Append "pos_map = ?, "
    QueryBuilder.Append "pos_x = ?, "
    QueryBuilder.Append "pos_y = ?, "
    QueryBuilder.Append "body_id = ?, "
    QueryBuilder.Append "head_id = ?, "
    QueryBuilder.Append "weapon_id = ?, "
    QueryBuilder.Append "helmet_id = ?, "
    QueryBuilder.Append "shield_id = ?, "
    QueryBuilder.Append "items_Amount = ?, "
    QueryBuilder.Append "slot_armour = ?, "
    QueryBuilder.Append "slot_weapon = ?, "
    QueryBuilder.Append "slot_shield = ?, "
    QueryBuilder.Append "slot_helmet = ?, "
    QueryBuilder.Append "slot_ammo = ?, "
    QueryBuilder.Append "slot_dm = ?, "
    QueryBuilder.Append "slot_rm = ?, "
    QueryBuilder.Append "slot_tool = ?, "
    QueryBuilder.Append "slot_magic = ?, "
    QueryBuilder.Append "slot_knuckles = ?, "
    QueryBuilder.Append "slot_ship = ?, "
    QueryBuilder.Append "slot_mount = ?, "
    QueryBuilder.Append "min_hp = ?, "
    QueryBuilder.Append "max_hp = ?, "
    QueryBuilder.Append "min_man = ?, "
    QueryBuilder.Append "max_man = ?, "
    QueryBuilder.Append "min_sta = ?, "
    QueryBuilder.Append "max_sta = ?, "
    QueryBuilder.Append "min_ham = ?, "
    QueryBuilder.Append "max_ham = ?, "
    QueryBuilder.Append "min_sed = ?, "
    QueryBuilder.Append "max_sed = ?, "
    QueryBuilder.Append "min_hit = ?, "
    QueryBuilder.Append "max_hit = ?, "
    QueryBuilder.Append "is_naked = ?, "
    QueryBuilder.Append "status = ?, "
    QueryBuilder.Append "is_logged = TRUE;"
    
    ' Guardo la query ensamblada
    QUERY_SAVE_MAINPJ = QueryBuilder.ToString
    
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    
    ' ************************** User attributes ********************************
    QueryBuilder.Append "INSERT INTO attribute (user_id, number, value) VALUES "

    For LoopC = 1 To NUMATRIBUTOS
        QueryBuilder.Append "(?, ?, ?)"

        If LoopC < NUMATRIBUTOS Then
            QueryBuilder.Append ", "
        End If

    Next LoopC
    
    ' Guardo la query ensamblada
    QUERY_SAVE_ATTRIBUTES = QueryBuilder.ToString
    
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
    QueryBuilder.Append "INSERT INTO inventory_item (user_id, number, item_id, Amount, is_equipped) VALUES "

    For LoopC = 1 To MAX_INVENTORY_SLOTS
        QueryBuilder.Append "(?, ?, ?, ?, ?)"

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
    QueryBuilder.Append "last_map = ?, "
    QueryBuilder.Append "message_info = ?, "
    QueryBuilder.Append "body_id = ?, "
    QueryBuilder.Append "head_id = ?, "
    QueryBuilder.Append "weapon_id = ?, "
    QueryBuilder.Append "helmet_id = ?, "
    QueryBuilder.Append "shield_id = ?, "
    QueryBuilder.Append "heading = ?, "
    QueryBuilder.Append "items_Amount = ?, "
    QueryBuilder.Append "slot_armour = ?, "
    QueryBuilder.Append "slot_weapon = ?, "
    QueryBuilder.Append "slot_shield = ?, "
    QueryBuilder.Append "slot_helmet = ?, "
    QueryBuilder.Append "slot_ammo = ?, "
    QueryBuilder.Append "slot_dm = ?, "
    QueryBuilder.Append "slot_rm = ?, "
    QueryBuilder.Append "slot_tool = ?, "
    QueryBuilder.Append "slot_magic = ?, "
    QueryBuilder.Append "slot_knuckles = ?, "
    QueryBuilder.Append "slot_ship = ?, "
    QueryBuilder.Append "slot_mount = ?, "
    QueryBuilder.Append "min_hp = ?, "
    QueryBuilder.Append "max_hp = ?, "
    QueryBuilder.Append "min_man = ?, "
    QueryBuilder.Append "max_man = ?, "
    QueryBuilder.Append "min_sta = ?, "
    QueryBuilder.Append "max_sta = ?, "
    QueryBuilder.Append "min_ham = ?, "
    QueryBuilder.Append "max_ham = ?, "
    QueryBuilder.Append "min_sed = ?, "
    QueryBuilder.Append "max_sed = ?, "
    QueryBuilder.Append "min_hit = ?, "
    QueryBuilder.Append "max_hit = ?, "
    QueryBuilder.Append "killed_npcs = ?, "
    QueryBuilder.Append "killed_users = ?, "
    QueryBuilder.Append "invent_level = ?, "
    QueryBuilder.Append "is_naked = ?, "
    QueryBuilder.Append "is_poisoned = ?, "
    QueryBuilder.Append "is_hidden = ?, "
    QueryBuilder.Append "is_hungry = ?, "
    QueryBuilder.Append "is_thirsty = ?, "
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
    QueryBuilder.Append "pertenece_consejo_real = ?, "
    QueryBuilder.Append "pertenece_consejo_caos = ?, "
    QueryBuilder.Append "pertenece_real = ?, "
    QueryBuilder.Append "pertenece_caos = ?, "
    QueryBuilder.Append "ciudadanos_matados = ?, "
    QueryBuilder.Append "criminales_matados = ?, "
    QueryBuilder.Append "recibio_armadura_real = ?, "
    QueryBuilder.Append "recibio_armadura_caos = ?, "
    QueryBuilder.Append "recibio_exp_real = ?, "
    QueryBuilder.Append "recibio_exp_caos = ?, "
    QueryBuilder.Append "recompensas_real = ?, "
    QueryBuilder.Append "recompensas_caos = ?, "
    QueryBuilder.Append "reenlistadas = ?, "
    QueryBuilder.Append "fecha_ingreso = ?, "
    QueryBuilder.Append "nivel_ingreso = ?, "
    QueryBuilder.Append "matados_ingreso = ?, "
    QueryBuilder.Append "siguiente_recompensa = ?, "
    QueryBuilder.Append "status = ?, "
    QueryBuilder.Append "guild_index = ?, "
    QueryBuilder.Append "chat_combate = ?, "
    QueryBuilder.Append "chat_global = ?, "
    QueryBuilder.Append "is_logged = ?, "
    QueryBuilder.Append "warnings = ? "
    QueryBuilder.Append " WHERE id = ?"
    
    ' Guardo la query ensamblada
    QUERY_UPDATE_MAINPJ = QueryBuilder.ToString
    
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    
    ' ************************** User bank inventory **************************************
    QueryBuilder.Append "INSERT INTO bank_item (user_id, number, item_id, amount) VALUES "

    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
        QueryBuilder.Append "(?, ?, ?, ?)"

        If LoopC < MAX_BANCOINVENTORY_SLOTS Then
            QueryBuilder.Append ", "

        End If

    Next LoopC
        
    QueryBuilder.Append " ON DUPLICATE KEY UPDATE item_id=VALUES(item_id), amount=VALUES(Amount); "
    
    ' Guardo la query ensamblada
    QUERY_SAVE_BANCOINV = QueryBuilder.ToString
    
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    
    ' ************************** UPSERT QUERIES **************************************
    
    QUERY_UPSERT_ATTRIBUTES = QUERY_SAVE_ATTRIBUTES & " ON DUPLICATE KEY UPDATE value=VALUES(value); "
    QUERY_UPSERT_SPELLS = QUERY_SAVE_SPELLS & " ON DUPLICATE KEY UPDATE spell_id=VALUES(spell_id); "
    QUERY_UPSERT_INVENTORY = QUERY_SAVE_INVENTORY & " ON DUPLICATE KEY UPDATE item_id=VALUES(item_id), amount=VALUES(Amount), is_equipped=VALUES(is_equipped); "
    QUERY_UPSERT_SKILLS = QUERY_SAVE_SKILLS & " ON DUPLICATE KEY UPDATE value=VALUES(value); "
    QUERY_UPSERT_PETS = QUERY_SAVE_PETS & " ON DUPLICATE KEY UPDATE pet_id=VALUES(pet_id); "
    
End Sub
