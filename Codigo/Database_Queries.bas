Attribute VB_Name = "Database_Queries"
Option Explicit

'Constructor de queries.
'Me permite concatenar strings MUCHO MAS rapido
Private QueryBuilder As cStringBuilder

' Queries
Public QUERY_CREARPJ_MAIN As String
Public QUERY_CREARPJ_ATRIBUTOS As String
Public QUERY_CREARPJ_SPELLS As String
Public QUERY_CREARPJ_INVENTORY As String
Public QUERY_CREARPJ_SKILLS As String
Public QUERY_CREARPJ_QUESTS As String
Public QUERY_CREARPJ_PETS As String

Public Sub Contruir_Querys()
    Call ConstruirQuery_CrearPersonaje
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
    QUERY_CREARPJ_MAIN = QueryBuilder.ToString
    
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    
    ' ************************** User attributes ********************************
    QueryBuilder.Append "INSERT INTO attribute (user_id, number, value) VALUES "

    For LoopC = 1 To NUMATRIBUTOS
        QueryBuilder.Append "(?, ?, ?)"

        If LoopC < NUMATRIBUTOS Then
            QueryBuilder.Append ", "
        Else
            QueryBuilder.Append "; "
        End If

    Next LoopC
    
    ' Guardo la query ensamblada
    QUERY_CREARPJ_ATRIBUTOS = QueryBuilder.ToString
    
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    
    ' ************************** User spells ************************************
    QueryBuilder.Append "INSERT INTO spell (user_id, number, spell_id) VALUES "

    For LoopC = 1 To MAXUSERHECHIZOS
        QueryBuilder.Append "(?, ?, ?)"

        If LoopC < MAXUSERHECHIZOS Then
            QueryBuilder.Append ", "
        Else
            QueryBuilder.Append "; "
        End If

    Next LoopC
    
    ' Guardo la query ensamblada
    QUERY_CREARPJ_SPELLS = QueryBuilder.ToString
    
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    
    ' ******************* INVENTORY *******************
    QueryBuilder.Append "INSERT INTO inventory_item (user_id, number, item_id, Amount, is_equipped) VALUES "

    For LoopC = 1 To MAX_INVENTORY_SLOTS
        QueryBuilder.Append "(?, ?, ?, ?, ?)"

        If LoopC < MAX_INVENTORY_SLOTS Then
            QueryBuilder.Append ", "
        Else
            QueryBuilder.Append "; "
        End If

    Next LoopC
    
    ' Guardo la query ensamblada
    QUERY_CREARPJ_INVENTORY = QueryBuilder.ToString
    
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear

    ' ************************** User skills ************************************
    QueryBuilder.Append "INSERT INTO skillpoint (user_id, number, value) VALUES "

    For LoopC = 1 To NUMSKILLS
        QueryBuilder.Append "(?, ?, ?)"

        If LoopC < NUMSKILLS Then
            QueryBuilder.Append ", "
        Else
            QueryBuilder.Append "; "
        End If

    Next LoopC

    ' Guardo la query ensamblada
    QUERY_CREARPJ_SKILLS = QueryBuilder.ToString
    
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear

    ' ************************** User quests ************************************
    QueryBuilder.Append "INSERT INTO quest (user_id, number) VALUES "

    For LoopC = 1 To MAXUSERQUESTS
        QueryBuilder.Append "(?, ?)"

        If LoopC < MAXUSERQUESTS Then
            QueryBuilder.Append ", "
        Else
            QueryBuilder.Append "; "
        End If

    Next LoopC
    
    ' Guardo la query ensamblada
    QUERY_CREARPJ_QUESTS = QueryBuilder.ToString
    
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    
    ' ************************** User pets **************************************
    QueryBuilder.Append "INSERT INTO pet (user_id, number, pet_id) VALUES "

    For LoopC = 1 To MAXMASCOTAS
        QueryBuilder.Append "(?, ?, 0)"

        If LoopC < MAXMASCOTAS Then
            QueryBuilder.Append ", "
        Else
            QueryBuilder.Append "; "
        End If

    Next LoopC
    
    ' Guardo la query ensamblada
    QUERY_CREARPJ_PETS = QueryBuilder.ToString
    
    ' Limpio el constructor de querys
    Call QueryBuilder.Clear
    
End Sub
