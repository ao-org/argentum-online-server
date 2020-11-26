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
 
    Set Database_Connection = New ADODB.Connection
    
    If Len(Database_DataSource) <> 0 Then
    
        Database_Connection.ConnectionString = "DATA SOURCE=" & Database_DataSource & ";"
        
    Else
    
        Database_Connection.ConnectionString = "DRIVER={MySQL ODBC 8.0 ANSI Driver};" & _
                                               "SERVER=" & Database_Host & ";" & _
                                               "DATABASE=" & Database_Name & ";" & _
                                               "USER=" & Database_Username & ";" & _
                                               "PASSWORD=" & Database_Password & ";" & _
                                               "OPTION=3;MULTI_STATEMENTS=1"
                                               
    End If
    
    Debug.Print Database_Connection.ConnectionString
    
    Database_Connection.CursorLocation = adUseClient
    
    Call Database_Connection.Open
    
    ConnectedOnce = True

    Exit Sub
    
ErrorHandler:
    Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.description)
    
    If Not ConnectedOnce Then
        Call MsgBox("No se pudo conectar a la base de datos. Mas información en logs/Database.log", vbCritical, "OBDC - Error")
        Call CerrarServidor
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
     
    Call Database_Connection.Close
    
    Set Database_Connection = Nothing
     
    Exit Sub
     
ErrorHandler:
    Call LogDatabaseError("Unable to close Mysql Database: " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveNewUserDatabase(ByVal UserIndex As Integer)

    On Error GoTo ErrorHandler
    
    'Constructor de queries.
    'Me permite concatenar strings MUCHO MAS rapido
    Set QueryBuilder = New cStringBuilder
    
    With UserList(UserIndex)
    
        'Basic user data
        QueryBuilder.Append "INSERT INTO user SET "
        QueryBuilder.Append "name = '" & .name & "', "
        QueryBuilder.Append "account_id = " & .AccountID & ", "
        QueryBuilder.Append "level = " & .Stats.ELV & ", "
        QueryBuilder.Append "exp = " & .Stats.Exp & ", "
        QueryBuilder.Append "elu = " & .Stats.ELU & ", "
        QueryBuilder.Append "genre_id = " & .genero & ", "
        QueryBuilder.Append "race_id = " & .raza & ", "
        QueryBuilder.Append "class_id = " & .clase & ", "
        QueryBuilder.Append "home_id = " & .Hogar & ", "
        QueryBuilder.Append "description = '" & .Desc & "', "
        QueryBuilder.Append "gold = " & .Stats.GLD & ", "
        QueryBuilder.Append "free_skillpoints = " & .Stats.SkillPts & ", "
        'QueryBuilder.Append "assigned_skillpoints = " & .Counters.AsignedSkills & ", "
        QueryBuilder.Append "pos_map = " & .Pos.Map & ", "
        QueryBuilder.Append "pos_x = " & .Pos.X & ", "
        QueryBuilder.Append "pos_y = " & .Pos.Y & ", "
        QueryBuilder.Append "body_id = " & .Char.Body & ", "
        QueryBuilder.Append "head_id = " & .Char.Head & ", "
        QueryBuilder.Append "weapon_id = " & .Char.WeaponAnim & ", "
        QueryBuilder.Append "helmet_id = " & .Char.CascoAnim & ", "
        QueryBuilder.Append "shield_id = " & .Char.ShieldAnim & ", "
        QueryBuilder.Append "items_Amount = " & .Invent.NroItems & ", "
        QueryBuilder.Append "slot_armour = " & .Invent.ArmourEqpSlot & ", "
        QueryBuilder.Append "slot_weapon = " & .Invent.WeaponEqpSlot & ", "
        QueryBuilder.Append "slot_shield = " & .Invent.EscudoEqpSlot & ", "
        QueryBuilder.Append "slot_helmet = " & .Invent.CascoEqpSlot & ", "
        QueryBuilder.Append "slot_ammo = " & .Invent.MunicionEqpSlot & ", "
        QueryBuilder.Append "slot_ring = " & .Invent.AnilloEqpSlot & ", "
        QueryBuilder.Append "slot_tool = " & .Invent.HerramientaEqpSlot & ", "
        QueryBuilder.Append "slot_magic = " & .Invent.MagicoSlot & ", "
        QueryBuilder.Append "slot_knuckles = " & .Invent.NudilloSlot & ", "
        QueryBuilder.Append "slot_ship = " & .Invent.BarcoSlot & ", "
        QueryBuilder.Append "slot_mount = " & .Invent.MonturaSlot & ", "
        QueryBuilder.Append "min_hp = " & .Stats.MinHp & ", "
        QueryBuilder.Append "max_hp = " & .Stats.MaxHp & ", "
        QueryBuilder.Append "min_man = " & .Stats.MinMAN & ", "
        QueryBuilder.Append "max_man = " & .Stats.MaxMAN & ", "
        QueryBuilder.Append "min_sta = " & .Stats.MinSta & ", "
        QueryBuilder.Append "max_sta = " & .Stats.MaxSta & ", "
        QueryBuilder.Append "min_ham = " & .Stats.MinHam & ", "
        QueryBuilder.Append "max_ham = " & .Stats.MaxHam & ", "
        QueryBuilder.Append "min_sed = " & .Stats.MinAGU & ", "
        QueryBuilder.Append "max_sed = " & .Stats.MaxAGU & ", "
        QueryBuilder.Append "min_hit = " & .Stats.MinHIT & ", "
        QueryBuilder.Append "max_hit = " & .Stats.MaxHit & ", "
        'QueryBuilder.Append "rep_noble = " & .NobleRep & ", "
        'QueryBuilder.Append "rep_plebe = " & .Reputacion.PlebeRep & ", "
        'QueryBuilder.Append "rep_average = " & .Reputacion.Promedio & ", "
        QueryBuilder.Append "is_naked = " & .flags.Desnudo & ", "
        QueryBuilder.Append "status = " & .Faccion.Status & ", "
        QueryBuilder.Append "is_logged = TRUE; "
        
        Call MakeQuery(QueryBuilder.toString, True)
        
        'Borramos la query construida.
        Call QueryBuilder.Clear
        
        ' Para recibir el ID del user
        Call MakeQuery("SELECT LAST_INSERT_ID();")

        If QueryData Is Nothing Then
            .Id = 1
        Else
            .Id = val(QueryData.Fields(0).Value)
        End If
        
        ' Comenzamos una cadena de queries (para enviar todo de una)
        Dim LoopC As Long

        'User attributes
        QueryBuilder.Append "INSERT INTO attribute (user_id, number, value) VALUES "

        For LoopC = 1 To NUMATRIBUTOS
        
            QueryBuilder.Append "("
            QueryBuilder.Append .Id & ", "
            QueryBuilder.Append LoopC & ", "
            QueryBuilder.Append .Stats.UserAtributos(LoopC) & ")"

            If LoopC < NUMATRIBUTOS Then
                QueryBuilder.Append ", "
            Else
                QueryBuilder.Append "; "
            End If

        Next LoopC

        'User spells
        QueryBuilder.Append "INSERT INTO spell (user_id, number, spell_id) VALUES "

        For LoopC = 1 To MAXUSERHECHIZOS
            QueryBuilder.Append "("
            QueryBuilder.Append .Id & ", "
            QueryBuilder.Append LoopC & ", "
            QueryBuilder.Append .Stats.UserHechizos(LoopC) & ")"

            If LoopC < MAXUSERHECHIZOS Then
                QueryBuilder.Append ", "
            Else
                QueryBuilder.Append "; "
            End If

        Next LoopC

        'User inventory
        QueryBuilder.Append "INSERT INTO inventory_item (user_id, number, item_id, Amount, is_equipped) VALUES "

        For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
            QueryBuilder.Append "("
            QueryBuilder.Append .Id & ", "
            QueryBuilder.Append LoopC & ", "
            QueryBuilder.Append .Invent.Object(LoopC).ObjIndex & ", "
            QueryBuilder.Append .Invent.Object(LoopC).Amount & ", "
            QueryBuilder.Append .Invent.Object(LoopC).Equipped & ")"

            If LoopC < UserList(UserIndex).CurrentInventorySlots Then
                QueryBuilder.Append ", "
            Else
                QueryBuilder.Append "; "
            End If

        Next LoopC

        'User skills
        'QueryBuilder.Append "INSERT INTO skillpoint (user_id, number, value, exp, elu) VALUES "
        QueryBuilder.Append "INSERT INTO skillpoint (user_id, number, value) VALUES "

        For LoopC = 1 To NUMSKILLS
            QueryBuilder.Append "("
            QueryBuilder.Append .Id & ", "
            QueryBuilder.Append LoopC & ", "
            QueryBuilder.Append .Stats.UserSkills(LoopC) & ")"
            'QueryBuilder.Append .Stats.UserSkills(LoopC) & ", "
            'QueryBuilder.Append .Stats.ExpSkills(LoopC) & ", "
            'QueryBuilder.Append .Stats.EluSkills(LoopC) & ")"

            If LoopC < NUMSKILLS Then
                QueryBuilder.Append ", "
            Else
                QueryBuilder.Append "; "
            End If

        Next LoopC
        
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
        QueryBuilder.Append "INSERT INTO quest (user_id, number) VALUES "

        For LoopC = 1 To MAXUSERQUESTS
        
            QueryBuilder.Append "("
            QueryBuilder.Append .Id & ", "
            QueryBuilder.Append LoopC & ")"

            If LoopC < MAXUSERQUESTS Then
                QueryBuilder.Append ", "
            Else
                QueryBuilder.Append "; "
            End If

        Next LoopC
        
        'User pets
        QueryBuilder.Append "INSERT INTO pet (user_id, number, pet_id) VALUES "

        For LoopC = 1 To MAXMASCOTAS
        
            QueryBuilder.Append "("
            QueryBuilder.Append .Id & ", "
            QueryBuilder.Append LoopC & ", 0)"

            If LoopC < MAXMASCOTAS Then
                QueryBuilder.Append ", "
            Else
                QueryBuilder.Append "; "
            End If

        Next LoopC

        'Enviamos todas las queries
        Call MakeQuery(QueryBuilder.toString, True)
        
        Set QueryBuilder = Nothing
    
    End With

    Exit Sub

ErrorHandler:
    
    Set QueryBuilder = Nothing
    
    Call LogDatabaseError("Error en SaveNewUserDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserDatabase(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean = False)

    On Error GoTo ErrorHandler
    
    'Constructor de queries.
    'Me permite concatenar strings MUCHO MAS rapido
    Set QueryBuilder = New cStringBuilder

    'Basic user data
    With UserList(UserIndex)
        QueryBuilder.Append "UPDATE user SET "
        QueryBuilder.Append "name = '" & .name & "', "
        QueryBuilder.Append "level = " & .Stats.ELV & ", "
        QueryBuilder.Append "exp = " & CLng(.Stats.Exp) & ", "
        QueryBuilder.Append "elu = " & .Stats.ELU & ", "
        QueryBuilder.Append "genre_id = " & .genero & ", "
        QueryBuilder.Append "race_id = " & .raza & ", "
        QueryBuilder.Append "class_id = " & .clase & ", "
        QueryBuilder.Append "home_id = " & .Hogar & ", "
        QueryBuilder.Append "description = '" & .Desc & "', "
        QueryBuilder.Append "gold = " & .Stats.GLD & ", "
        QueryBuilder.Append "bank_gold = " & .Stats.Banco & ", "
        QueryBuilder.Append "free_skillpoints = " & .Stats.SkillPts & ", "
        'QueryBuilder.Append "assigned_skillpoints = " & .Counters.AsignedSkills & ", "
        'QueryBuilder.Append "pet_Amount = " & .NroMascotas & ", "
        QueryBuilder.Append "pos_map = " & .Pos.Map & ", "
        QueryBuilder.Append "pos_x = " & .Pos.X & ", "
        QueryBuilder.Append "pos_y = " & .Pos.Y & ", "
        QueryBuilder.Append "message_info = '" & .MENSAJEINFORMACION & "', "
        QueryBuilder.Append "body_id = " & .Char.Body & ", "
        QueryBuilder.Append "head_id = " & .OrigChar.Head & ", "
        QueryBuilder.Append "weapon_id = " & .Char.WeaponAnim & ", "
        QueryBuilder.Append "helmet_id = " & .Char.CascoAnim & ", "
        QueryBuilder.Append "shield_id = " & .Char.ShieldAnim & ", "
        QueryBuilder.Append "heading = " & .Char.Heading & ", "
        QueryBuilder.Append "items_Amount = " & .Invent.NroItems & ", "
        QueryBuilder.Append "slot_armour = " & .Invent.ArmourEqpSlot & ", "
        QueryBuilder.Append "slot_weapon = " & .Invent.WeaponEqpSlot & ", "
        QueryBuilder.Append "slot_shield = " & .Invent.EscudoEqpSlot & ", "
        QueryBuilder.Append "slot_helmet = " & .Invent.CascoEqpSlot & ", "
        QueryBuilder.Append "slot_ammo = " & .Invent.MunicionEqpSlot & ", "
        QueryBuilder.Append "slot_ring = " & .Invent.AnilloEqpSlot & ", "
        QueryBuilder.Append "slot_tool = " & .Invent.HerramientaEqpSlot & ", "
        QueryBuilder.Append "slot_magic = " & .Invent.MagicoSlot & ", "
        QueryBuilder.Append "slot_knuckles = " & .Invent.NudilloSlot & ", "
        QueryBuilder.Append "slot_ship = " & .Invent.BarcoSlot & ", "
        QueryBuilder.Append "slot_mount = " & .Invent.MonturaSlot & ", "
        QueryBuilder.Append "min_hp = " & .Stats.MinHp & ", "
        QueryBuilder.Append "max_hp = " & .Stats.MaxHp & ", "
        QueryBuilder.Append "min_man = " & .Stats.MinMAN & ", "
        QueryBuilder.Append "max_man = " & .Stats.MaxMAN & ", "
        QueryBuilder.Append "min_sta = " & .Stats.MinSta & ", "
        QueryBuilder.Append "max_sta = " & .Stats.MaxSta & ", "
        QueryBuilder.Append "min_ham = " & .Stats.MinHam & ", "
        QueryBuilder.Append "max_ham = " & .Stats.MaxHam & ", "
        QueryBuilder.Append "min_sed = " & .Stats.MinAGU & ", "
        QueryBuilder.Append "max_sed = " & .Stats.MaxAGU & ", "
        QueryBuilder.Append "min_hit = " & .Stats.MinHIT & ", "
        QueryBuilder.Append "max_hit = " & .Stats.MaxHit & ", "
        QueryBuilder.Append "killed_npcs = " & .Stats.NPCsMuertos & ", "
        QueryBuilder.Append "killed_users = " & .Stats.UsuariosMatados & ", "
        QueryBuilder.Append "invent_level = " & .Stats.InventLevel & ", "
        'QueryBuilder.Append "rep_asesino = " & .Reputacion.AsesinoRep & ", "
        'QueryBuilder.Append "rep_bandido = " & .Reputacion.BandidoRep & ", "
        'QueryBuilder.Append "rep_burgues = " & .Reputacion.BurguesRep & ", "
        'QueryBuilder.Append "rep_ladron = " & .Reputacion.LadronesRep & ", "
        'QueryBuilder.Append "rep_noble = " & .Reputacion.NobleRep & ", "
        'QueryBuilder.Append "rep_plebe = " & .Reputacion.PlebeRep & ", "
        'QueryBuilder.Append "rep_average = " & .Reputacion.Promedio & ", "
        QueryBuilder.Append "is_naked = " & .flags.Desnudo & ", "
        QueryBuilder.Append "is_poisoned = " & .flags.Envenenado & ", "
        QueryBuilder.Append "is_hidden = " & .flags.Escondido & ", "
        QueryBuilder.Append "is_hungry = " & .flags.Hambre & ", "
        QueryBuilder.Append "is_thirsty = " & .flags.Sed & ", "
        'QueryBuilder.Append "is_banned = " & .flags.Ban & ", " Esto es innecesario porque se setea cuando lo baneas (creo)
        QueryBuilder.Append "is_dead = " & .flags.Muerto & ", "
        QueryBuilder.Append "is_sailing = " & .flags.Navegando & ", "
        QueryBuilder.Append "is_paralyzed = " & .flags.Paralizado & ", "
        QueryBuilder.Append "is_mounted = " & .flags.Montado & ", "
        QueryBuilder.Append "is_silenced = " & .flags.Silenciado & ", "
        QueryBuilder.Append "silence_minutes_left = " & .flags.MinutosRestantes & ", "
        QueryBuilder.Append "silence_elapsed_seconds = " & .flags.SegundosPasados & ", "
        QueryBuilder.Append "spouse = '" & .flags.Pareja & "', "
        QueryBuilder.Append "counter_pena = " & .Counters.Pena & ", "
        QueryBuilder.Append "deaths = " & .flags.VecesQueMoriste & ", "
        QueryBuilder.Append "pertenece_consejo_real = " & (.flags.Privilegios And PlayerType.RoyalCouncil) & ", "
        QueryBuilder.Append "pertenece_consejo_caos = " & (.flags.Privilegios And PlayerType.ChaosCouncil) & ", "
        QueryBuilder.Append "pertenece_real = " & .Faccion.ArmadaReal & ", "
        QueryBuilder.Append "pertenece_caos = " & .Faccion.FuerzasCaos & ", "
        QueryBuilder.Append "ciudadanos_matados = " & .Faccion.CiudadanosMatados & ", "
        QueryBuilder.Append "criminales_matados = " & .Faccion.CriminalesMatados & ", "
        QueryBuilder.Append "recibio_armadura_real = " & .Faccion.RecibioArmaduraReal & ", "
        QueryBuilder.Append "recibio_armadura_caos = " & .Faccion.RecibioArmaduraCaos & ", "
        QueryBuilder.Append "recibio_exp_real = " & .Faccion.RecibioExpInicialReal & ", "
        QueryBuilder.Append "recibio_exp_caos = " & .Faccion.RecibioExpInicialCaos & ", "
        QueryBuilder.Append "recompensas_real = " & .Faccion.RecompensasReal & ", "
        QueryBuilder.Append "recompensas_caos = " & .Faccion.RecompensasCaos & ", "
        QueryBuilder.Append "reenlistadas = " & .Faccion.Reenlistadas & ", "
        QueryBuilder.Append "fecha_ingreso = " & IIf(.Faccion.FechaIngreso <> vbNullString, "'" & .Faccion.FechaIngreso & "'", "NULL") & ", "
        QueryBuilder.Append "nivel_ingreso = " & .Faccion.NivelIngreso & ", "
        QueryBuilder.Append "matados_ingreso = " & .Faccion.MatadosIngreso & ", "
        QueryBuilder.Append "siguiente_recompensa = " & .Faccion.NextRecompensa & ", "
        QueryBuilder.Append "status = " & .Faccion.Status & ", "
        QueryBuilder.Append "battle_points = " & .flags.BattlePuntos & ", "
        QueryBuilder.Append "guild_index = " & .GuildIndex & ", "
        QueryBuilder.Append "chat_combate = " & .ChatCombate & ", "
        QueryBuilder.Append "chat_global = " & .ChatGlobal & ", "
        QueryBuilder.Append "is_logged = " & IIf(Logout, "FALSE", "TRUE")
        QueryBuilder.Append " WHERE id = " & .Id & "; "
        
        Dim LoopC As Long

        'User attributes
        QueryBuilder.Append "INSERT INTO attribute (user_id, number, value) VALUES "

        For LoopC = 1 To NUMATRIBUTOS
        
            QueryBuilder.Append "("
            QueryBuilder.Append .Id & ", "
            QueryBuilder.Append LoopC & ", "
            QueryBuilder.Append .Stats.UserAtributosBackUP(LoopC) & ")"

            If LoopC < NUMATRIBUTOS Then
                QueryBuilder.Append ", "
            End If

        Next LoopC
        
        QueryBuilder.Append " ON DUPLICATE KEY UPDATE value=VALUES(value); "

        'User spells
        QueryBuilder.Append "INSERT INTO spell (user_id, number, spell_id) VALUES "

        For LoopC = 1 To MAXUSERHECHIZOS
        
            QueryBuilder.Append "("
            QueryBuilder.Append .Id & ", "
            QueryBuilder.Append LoopC & ", "
            QueryBuilder.Append .Stats.UserHechizos(LoopC) & ")"

            If LoopC < MAXUSERHECHIZOS Then
                QueryBuilder.Append ", "
            End If

        Next LoopC
        
        QueryBuilder.Append " ON DUPLICATE KEY UPDATE spell_id=VALUES(spell_id); "

        'User inventory
        QueryBuilder.Append "INSERT INTO inventory_item (user_id, number, item_id, Amount, is_equipped) VALUES "

        For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
        
            QueryBuilder.Append "("
            QueryBuilder.Append .Id & ", "
            QueryBuilder.Append LoopC & ", "
            QueryBuilder.Append .Invent.Object(LoopC).ObjIndex & ", "
            QueryBuilder.Append .Invent.Object(LoopC).Amount & ", "
            QueryBuilder.Append .Invent.Object(LoopC).Equipped & ")"

            If LoopC < UserList(UserIndex).CurrentInventorySlots Then
                QueryBuilder.Append ", "
            End If

        Next LoopC
        
        QueryBuilder.Append " ON DUPLICATE KEY UPDATE item_id=VALUES(item_id), amount=VALUES(Amount), is_equipped=VALUES(is_equipped); "

        'User bank inventory
        QueryBuilder.Append "INSERT INTO bank_item (user_id, number, item_id, Amount) VALUES "

        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
        
            QueryBuilder.Append "("
            QueryBuilder.Append .Id & ", "
            QueryBuilder.Append LoopC & ", "
            QueryBuilder.Append .BancoInvent.Object(LoopC).ObjIndex & ", "
            QueryBuilder.Append .BancoInvent.Object(LoopC).Amount & ")"

            If LoopC < MAX_BANCOINVENTORY_SLOTS Then
                QueryBuilder.Append ", "
            End If

        Next LoopC
        
        QueryBuilder.Append " ON DUPLICATE KEY UPDATE item_id=VALUES(item_id), amount=VALUES(Amount); "

        'User skills
        'QueryBuilder.Append "INSERT INTO skillpoint (user_id, number, value, exp, elu) VALUES "
        QueryBuilder.Append "INSERT INTO skillpoint (user_id, number, value) VALUES "

        For LoopC = 1 To NUMSKILLS
        
            QueryBuilder.Append "("
            QueryBuilder.Append .Id & ", "
            QueryBuilder.Append LoopC & ", "
            QueryBuilder.Append .Stats.UserSkills(LoopC) & ")"
            'QueryBuilder.Append .Stats.UserSkills(LoopC) & ", "
            'Q  = Q & .Stats.ExpSkills(LoopC) & ", "
            'QueryBuilder.Append .Stats.EluSkills(LoopC) & ")"

            If LoopC < NUMSKILLS Then
                QueryBuilder.Append ", "
            End If

        Next LoopC
        
        'QueryBuilder.Append " ON DUPLICATE KEY UPDATE value=VALUES(value), exp=VALUES(exp), elu=VALUES(elu); "
        QueryBuilder.Append " ON DUPLICATE KEY UPDATE value=VALUES(value); "

        'User pets
        Dim petType As Integer
        
        QueryBuilder.Append "INSERT INTO pet (user_id, number, pet_id) VALUES "

        For LoopC = 1 To MAXMASCOTAS
            QueryBuilder.Append "("
            QueryBuilder.Append .Id & ", "
            QueryBuilder.Append LoopC & ", "

            'CHOTS | I got this logic from SaveUserToCharfile
            If .MascotasIndex(LoopC) > 0 Then
            
                If Npclist(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
                    petType = .MascotasType(LoopC)
                Else
                    petType = 0
                End If

            Else
                petType = .MascotasType(LoopC)

            End If

            QueryBuilder.Append petType & ")"

            If LoopC < MAXMASCOTAS Then
                QueryBuilder.Append ", "
            End If

        Next LoopC

        QueryBuilder.Append " ON DUPLICATE KEY UPDATE pet_id=VALUES(pet_id); "
        
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
        QueryBuilder.Append "INSERT INTO connection (user_id, ip, date_last_login) VALUES ("
        QueryBuilder.Append .Id & ", "
        QueryBuilder.Append "'" & .ip & "', "
        QueryBuilder.Append "NOW()) "
        QueryBuilder.Append "ON DUPLICATE KEY UPDATE "
        QueryBuilder.Append "date_last_login = VALUES(date_last_login); "
        
        'Borro la mas vieja si hay mas de 5 (WyroX: si alguien sabe una forma mejor de hacerlo me avisa)
        QueryBuilder.Append "DELETE FROM connection WHERE"
        QueryBuilder.Append " user_id = " & .Id
        QueryBuilder.Append " AND date_last_login < (SELECT min(date_last_login) FROM (SELECT date_last_login FROM connection WHERE"
        QueryBuilder.Append " user_id = " & .Id
        QueryBuilder.Append " ORDER BY date_last_login DESC LIMIT 5) AS d); "
        
        'User quests
        QueryBuilder.Append "INSERT INTO quest (user_id, number, quest_id, npcs) VALUES "
        
        Dim Tmp As Integer, LoopK As Long

        For LoopC = 1 To MAXUSERQUESTS
            QueryBuilder.Append "("
            QueryBuilder.Append .Id & ", "
            QueryBuilder.Append LoopC & ", "
            QueryBuilder.Append .QuestStats.Quests(LoopC).QuestIndex & ", '"
            
            If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
                Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredNPCs

                If Tmp Then

                    For LoopK = 1 To Tmp
                        QueryBuilder.Append CStr(.QuestStats.Quests(LoopC).NPCsKilled(LoopK))
                        
                        If LoopK < Tmp Then
                            QueryBuilder.Append "-"
                        End If

                    Next LoopK

                End If

            End If
            
            QueryBuilder.Append "')"

            If LoopC < MAXUSERQUESTS Then
                QueryBuilder.Append ", "
            End If

        Next LoopC
        
        QueryBuilder.Append " ON DUPLICATE KEY UPDATE quest_id=VALUES(quest_id), npcs=VALUES(npcs); "
        
        'User completed quests
        If .QuestStats.NumQuestsDone > 0 Then
            QueryBuilder.Append "INSERT INTO quest_done (user_id, quest_id) VALUES "
    
            For LoopC = 1 To .QuestStats.NumQuestsDone
                QueryBuilder.Append "("
                QueryBuilder.Append .Id & ", "
                QueryBuilder.Append .QuestStats.QuestsDone(LoopC) & ")"
    
                If LoopC < .QuestStats.NumQuestsDone Then
                    QueryBuilder.Append ", "
                End If
    
            Next LoopC
            
            QueryBuilder.Append " ON DUPLICATE KEY UPDATE quest_id=VALUES(quest_id); "

        End If
        
        'User mail
        'TODO:
        
        ' Si deslogueó, actualizo la cuenta
        If Logout Then
            QueryBuilder.Append "UPDATE account SET logged = logged - 1 WHERE id = " & .AccountID & ";"
        End If

        Call MakeQuery(QueryBuilder.toString, True)

    End With
    
    Set QueryBuilder = Nothing
    
    Exit Sub

ErrorHandler:

    Set QueryBuilder = Nothing
    
    Call LogDatabaseError("Error en SaveUserDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.description)

End Sub

Sub LoadUserDatabase(ByVal UserIndex As Integer)

    On Error GoTo ErrorHandler

    'Basic user data
    With UserList(UserIndex)

        Call MakeQuery("SELECT *, DATE_FORMAT(fecha_ingreso, '%Y-%m-%d') as 'fecha_ingreso_format' FROM user WHERE name ='" & .name & "';")

        If QueryData Is Nothing Then Exit Sub

        'Start setting data
        .Id = QueryData!Id
        .name = QueryData!name
        .Stats.ELV = QueryData!level
        .Stats.Exp = QueryData!Exp
        .Stats.ELU = QueryData!ELU
        .genero = QueryData!genre_id
        .raza = QueryData!race_id
        .clase = QueryData!class_id
        .Hogar = QueryData!home_id
        .Desc = QueryData!description
        .Stats.GLD = QueryData!gold
        .Stats.Banco = QueryData!bank_gold
        .Stats.SkillPts = QueryData!free_skillpoints
        '.Counters.AsignedSkills = QueryData!assigned_skillpoints
        '.NroMascotas = QueryData!pet_Amount
        .Pos.Map = QueryData!pos_map
        .Pos.X = QueryData!pos_x
        .Pos.Y = QueryData!pos_y
        .MENSAJEINFORMACION = QueryData!message_info
        .OrigChar.Body = QueryData!body_id
        .OrigChar.Head = QueryData!head_id
        .OrigChar.WeaponAnim = QueryData!weapon_id
        .OrigChar.CascoAnim = QueryData!helmet_id
        .OrigChar.ShieldAnim = QueryData!shield_id
        .OrigChar.Heading = QueryData!Heading
        .Invent.NroItems = QueryData!items_Amount
        .Invent.ArmourEqpSlot = SanitizeNullValue(QueryData!slot_armour, 0)
        .Invent.WeaponEqpSlot = SanitizeNullValue(QueryData!slot_weapon, 0)
        .Invent.CascoEqpSlot = SanitizeNullValue(QueryData!slot_helmet, 0)
        .Invent.EscudoEqpSlot = SanitizeNullValue(QueryData!slot_shield, 0)
        .Invent.MunicionEqpSlot = SanitizeNullValue(QueryData!slot_ammo, 0)
        .Invent.BarcoSlot = SanitizeNullValue(QueryData!slot_ship, 0)
        .Invent.MonturaSlot = SanitizeNullValue(QueryData!slot_mount, 0)
        .Invent.AnilloEqpSlot = SanitizeNullValue(QueryData!slot_ring, 0)
        .Invent.NudilloSlot = SanitizeNullValue(QueryData!slot_knuckles, 0)
        .Invent.HerramientaEqpSlot = SanitizeNullValue(QueryData!slot_tool, 0)
        .Invent.MagicoSlot = SanitizeNullValue(QueryData!slot_magic, 0)
        .Stats.MinHp = QueryData!min_hp
        .Stats.MaxHp = QueryData!max_hp
        .Stats.MinMAN = QueryData!min_man
        .Stats.MaxMAN = QueryData!max_man
        .Stats.MinSta = QueryData!min_sta
        .Stats.MaxSta = QueryData!max_sta
        .Stats.MinHam = QueryData!min_ham
        .Stats.MaxHam = QueryData!max_ham
        .Stats.MinAGU = QueryData!min_sed
        .Stats.MaxAGU = QueryData!max_sed
        .Stats.MinHIT = QueryData!min_hit
        .Stats.MaxHit = QueryData!max_hit
        .Stats.NPCsMuertos = QueryData!killed_npcs
        .Stats.UsuariosMatados = QueryData!killed_users
        .Stats.InventLevel = QueryData!invent_level
        '.Reputacion.AsesinoRep = QueryData!rep_asesino
        '.Reputacion.BandidoRep = QueryData!rep_bandido
        '.Reputacion.BurguesRep = QueryData!rep_burgues
        '.Reputacion.LadronesRep = QueryData!rep_ladron
        '.Reputacion.NobleRep = QueryData!rep_noble
        '.Reputacion.PlebeRep = QueryData!rep_plebe
        '.Reputacion.Promedio = QueryData!rep_average
        .flags.Desnudo = QueryData!is_naked
        .flags.Envenenado = QueryData!is_poisoned
        .flags.Escondido = QueryData!is_hidden
        .flags.Hambre = QueryData!is_hungry
        .flags.Sed = QueryData!is_thirsty
        .flags.Ban = QueryData!is_banned
        .flags.Muerto = QueryData!is_dead
        .flags.Navegando = QueryData!is_sailing
        .flags.Paralizado = QueryData!is_paralyzed
        .flags.VecesQueMoriste = QueryData!deaths
        .flags.BattlePuntos = QueryData!battle_points
        .flags.Montado = QueryData!is_mounted
        .flags.Pareja = QueryData!spouse
        .flags.Casado = IIf(Len(.flags.Pareja) > 0, 1, 0)
        .flags.Silenciado = QueryData!is_silenced
        .flags.MinutosRestantes = QueryData!silence_minutes_left
        .flags.SegundosPasados = QueryData!silence_elapsed_seconds
        .flags.ScrollExp = 1 'TODO: sacar
        .flags.ScrollOro = 1 'TODO: sacar
        
        .Counters.Pena = QueryData!counter_pena
        
        .ChatGlobal = QueryData!chat_global
        .ChatCombate = QueryData!chat_combate

        If QueryData!pertenece_consejo_real Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.RoyalCouncil

        End If

        If QueryData!pertenece_consejo_caos Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.ChaosCouncil

        End If

        .Faccion.ArmadaReal = QueryData!pertenece_real
        .Faccion.FuerzasCaos = QueryData!pertenece_caos
        .Faccion.CiudadanosMatados = QueryData!ciudadanos_matados
        .Faccion.CriminalesMatados = QueryData!criminales_matados
        .Faccion.RecibioArmaduraReal = QueryData!recibio_armadura_real
        .Faccion.RecibioArmaduraCaos = QueryData!recibio_armadura_caos
        .Faccion.RecibioExpInicialReal = QueryData!recibio_exp_real
        .Faccion.RecibioExpInicialCaos = QueryData!recibio_exp_caos
        .Faccion.RecompensasReal = QueryData!recompensas_real
        .Faccion.RecompensasCaos = QueryData!recompensas_caos
        .Faccion.Reenlistadas = QueryData!Reenlistadas
        .Faccion.FechaIngreso = SanitizeNullValue(QueryData!fecha_ingreso_format, vbNullString)
        .Faccion.NivelIngreso = SanitizeNullValue(QueryData!nivel_ingreso, 0)
        .Faccion.MatadosIngreso = SanitizeNullValue(QueryData!matados_ingreso, 0)
        .Faccion.NextRecompensa = SanitizeNullValue(QueryData!siguiente_recompensa, 0)
        .Faccion.Status = QueryData!Status

        .GuildIndex = SanitizeNullValue(QueryData!Guild_Index, 0)

        'User attributes
        Call MakeQuery("SELECT * FROM attribute WHERE user_id = " & .Id & ";")
    
        If Not QueryData Is Nothing Then
            QueryData.MoveFirst

            While Not QueryData.EOF

                .Stats.UserAtributos(QueryData!Number) = QueryData!Value
                .Stats.UserAtributosBackUP(QueryData!Number) = .Stats.UserAtributos(QueryData!Number)

                QueryData.MoveNext
            Wend

        End If

        'User spells
        Call MakeQuery("SELECT * FROM spell WHERE user_id = " & .Id & ";")

        If Not QueryData Is Nothing Then
            QueryData.MoveFirst

            While Not QueryData.EOF

                .Stats.UserHechizos(QueryData!Number) = QueryData!spell_id

                QueryData.MoveNext
            Wend

        End If

        'User pets
        Call MakeQuery("SELECT * FROM pet WHERE user_id = " & .Id & ";")

        If Not QueryData Is Nothing Then
            QueryData.MoveFirst

            While Not QueryData.EOF

                .MascotasType(QueryData!Number) = QueryData!pet_id

                QueryData.MoveNext
            Wend
        End If

        'User inventory
        Call MakeQuery("SELECT * FROM inventory_item WHERE user_id = " & .Id & ";")

        If Not QueryData Is Nothing Then
            QueryData.MoveFirst

            While Not QueryData.EOF

                .Invent.Object(QueryData!Number).ObjIndex = QueryData!item_id
                .Invent.Object(QueryData!Number).Amount = QueryData!Amount
                .Invent.Object(QueryData!Number).Equipped = QueryData!is_equipped

                QueryData.MoveNext
            Wend

        End If

        'User bank inventory
        Call MakeQuery("SELECT * FROM bank_item WHERE user_id = " & .Id & ";")

        If Not QueryData Is Nothing Then
            QueryData.MoveFirst

            While Not QueryData.EOF

                .BancoInvent.Object(QueryData!Number).ObjIndex = QueryData!item_id
                .BancoInvent.Object(QueryData!Number).Amount = QueryData!Amount

                QueryData.MoveNext
            Wend

        End If

        'User skills
        Call MakeQuery("SELECT * FROM skillpoint WHERE user_id = " & .Id & ";")

        If Not QueryData Is Nothing Then
            QueryData.MoveFirst

            While Not QueryData.EOF

                .Stats.UserSkills(QueryData!Number) = QueryData!Value
                '.Stats.ExpSkills(QueryData!Number) = QueryData!Exp
                '.Stats.EluSkills(QueryData!Number) = QueryData!ELU

                QueryData.MoveNext
            Wend

        End If

        'User friends
        'Call MakeQuery("SELECT * FROM friend WHERE user_id = " & .ID & ";")

        'If Not QueryData Is Nothing Then
        '    QueryData.MoveFirst

        '    While Not QueryData.EOF

        '        .Amigos(QueryData!Number).Nombre = QueryData!friend
        '        .Amigos(QueryData!Number).Ignorado = QueryData!Ignored

        '        QueryData.MoveNext
        '    Wend
        'End If
        
        Dim LoopC As Byte
        
        'User quests
        Call MakeQuery("SELECT * FROM quest WHERE user_id = " & .Id & ";")

        If Not QueryData Is Nothing Then
            QueryData.MoveFirst

            While Not QueryData.EOF

                .QuestStats.Quests(QueryData!Number).QuestIndex = QueryData!quest_id
                
                If .QuestStats.Quests(QueryData!Number).QuestIndex > 0 Then
                    If QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredNPCs Then

                        Dim NPCs() As String

                        NPCs = Split(QueryData!NPCs, "-")
                        ReDim .QuestStats.Quests(QueryData!Number).NPCsKilled(1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredNPCs)

                        For LoopC = 1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredNPCs
                            .QuestStats.Quests(QueryData!Number).NPCsKilled(LoopC) = val(NPCs(LoopC - 1))
                        Next LoopC

                    End If

                End If

                QueryData.MoveNext
            Wend

        End If
        
        'User quests done
        Call MakeQuery("SELECT * FROM quest_done WHERE user_id = " & .Id & ";")

        If Not QueryData Is Nothing Then
            .QuestStats.NumQuestsDone = QueryData.RecordCount
                
            ReDim .QuestStats.QuestsDone(1 To .QuestStats.NumQuestsDone)
        
            QueryData.MoveFirst
            
            LoopC = 1

            While Not QueryData.EOF
            
                .QuestStats.QuestsDone(LoopC) = QueryData!quest_id
                LoopC = LoopC + 1

                QueryData.MoveNext
            Wend

        End If
        
        'User mail
        'TODO:
        
        ' Llaves
        Call MakeQuery("SELECT key_obj FROM house_key WHERE account_id = " & .AccountID & ";")

        If Not QueryData Is Nothing Then
            QueryData.MoveFirst

            LoopC = 1

            While Not QueryData.EOF
                .Keys(LoopC) = QueryData!key_obj
                LoopC = LoopC + 1

                QueryData.MoveNext
            Wend

        End If

    End With

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to LOAD User from Mysql Database: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.description)

End Sub

Private Sub MakeQuery(Query As String, Optional ByVal NoResult As Boolean = False)
    ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Hace una unica query a la db. Asume una conexion.
    ' Si NoResult = False, el metodo lee el resultado de la query
    ' Guarda el resultado en QueryData
    
    On Error GoTo ErrorHandler

    ' Evito memory leaks
    If Not QueryData Is Nothing Then
        Call QueryData.Close
        Set QueryData = Nothing

    End If
    
    If NoResult Then
        Call Database_Connection.Execute(Query)

    Else
        Set QueryData = Database_Connection.Execute(Query)

        If QueryData.BOF Or QueryData.EOF Then
            Set QueryData = Nothing

        End If

    End If
    
    Exit Sub
    
ErrorHandler:

    If Database_Connection.State = adStateClosed Then
        Call LogDatabaseError("Alarma en MakeQuery: Se perdió la conexión con la DB. Reconectando.")
        Database_Connect
        Resume
    Else
        Call LogDatabaseError("Error en MakeQuery: query = '" & Query & "'. " & Err.Number & " - " & Err.description)
        
On Error GoTo 0
        Err.raise Err.Number
    End If

End Sub

Private Function GetDBValue(Tabla As String, ColumnaGet As String, ColumnaTest As String, ValueTest As Variant) As Variant
    ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para leer un unico valor de una unica fila

    On Error GoTo ErrorHandler
    
    'Hacemos la query segun el tipo de variable.
    If VarType(ValueTest) = vbString Then
        Call MakeQuery("SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = '" & ValueTest & "';")
    Else
        Call MakeQuery("SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = " & ValueTest & ";")  ' Sin comillas

    End If
    
    'Revisamos si recibio un resultado
    If QueryData Is Nothing Then Exit Function

    'Obtenemos la variable
    GetDBValue = QueryData.Fields(0).Value

    Exit Function
    
ErrorHandler:
    Call LogDatabaseError("Error en GetDBValue: SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = '" & ValueTest & "';" & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetCuentaValue(CuentaEmail As String, Columna As String) As Variant
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que leer un unico valor de la cuenta
        
        On Error GoTo GetCuentaValue_Err
        
100     GetCuentaValue = GetDBValue("account", Columna, "email", LCase$(CuentaEmail))

        
        Exit Function

GetCuentaValue_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetCuentaValue", Erl)
        Resume Next
        
End Function

Public Function GetUserValue(CharName As String, Columna As String) As Variant
        ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que leer un unico valor del char
        
        On Error GoTo GetUserValue_Err
        
100     GetUserValue = GetDBValue("user", Columna, "name", CharName)

        
        Exit Function

GetUserValue_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetUserValue", Erl)
        Resume Next
        
End Function

Private Sub SetDBValue(Tabla As String, ColumnaSet As String, ByVal ValueSet As Variant, ColumnaTest As String, ByVal ValueTest As Variant)
    ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para escribir un unico valor de una unica fila

    On Error GoTo ErrorHandler

    'Agregamos comillas a strings
    If VarType(ValueSet) = vbString Then
        ValueSet = "'" & ValueSet & "'"

    End If
    
    If VarType(ValueTest) = vbString Then
        ValueTest = "'" & ValueTest & "'"

    End If
    
    'Hacemos la query
    Call MakeQuery("UPDATE " & Tabla & " SET " & ColumnaSet & " = " & ValueSet & " WHERE " & ColumnaTest & " = " & ValueTest & ";", True)

    Exit Sub
    
ErrorHandler:
    Call LogDatabaseError("Error en SetDBValue: UPDATE " & Tabla & " SET " & ColumnaSet & " = " & ValueSet & " WHERE " & ColumnaTest & " = " & ValueTest & ";" & ". " & Err.Number & " - " & Err.description)

End Sub

Private Sub SetCuentaValue(CuentaEmail As String, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor de la cuenta
        
        On Error GoTo SetCuentaValue_Err
        
100     Call SetDBValue("account", Columna, Value, "email", LCase$(CuentaEmail))

        
        Exit Sub

SetCuentaValue_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SetCuentaValue", Erl)
        Resume Next
        
End Sub

Private Sub SetUserValue(CharName As String, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor del char
        
        On Error GoTo SetUserValue_Err
        
100     Call SetDBValue("user", Columna, Value, "name", CharName)

        
        Exit Sub

SetUserValue_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SetUserValue", Erl)
        Resume Next
        
End Sub

Private Sub SetCuentaValueByID(AccountID As Long, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor de la cuenta
        ' Por ID
        
        On Error GoTo SetCuentaValueByID_Err
        
100     Call SetDBValue("account", Columna, Value, "id", AccountID)

        
        Exit Sub

SetCuentaValueByID_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SetCuentaValueByID", Erl)
        Resume Next
        
End Sub

Private Sub SetUserValueByID(Id As Long, Columna As String, Value As Variant)
        ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
        ' Para cuando hay que escribir un unico valor del char
        ' Por ID
        
        On Error GoTo SetUserValueByID_Err
        
100     Call SetDBValue("user", Columna, Value, "id", Id)

        
        Exit Sub

SetUserValueByID_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SetUserValueByID", Erl)
        Resume Next
        
End Sub

Public Function CheckUserDonatorDatabase(CuentaEmail As String) As Boolean
        
        On Error GoTo CheckUserDonatorDatabase_Err
        
100     CheckUserDonatorDatabase = GetCuentaValue(CuentaEmail, "is_donor")

        
        Exit Function

CheckUserDonatorDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.CheckUserDonatorDatabase", Erl)
        Resume Next
        
End Function

Public Function GetUserCreditosDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetUserCreditosDatabase_Err
        
100     GetUserCreditosDatabase = GetCuentaValue(CuentaEmail, "credits")

        
        Exit Function

GetUserCreditosDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetUserCreditosDatabase", Erl)
        Resume Next
        
End Function

Public Function GetUserCreditosCanjeadosDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetUserCreditosCanjeadosDatabase_Err
        
100     GetUserCreditosCanjeadosDatabase = GetCuentaValue(CuentaEmail, "credits_used")

        
        Exit Function

GetUserCreditosCanjeadosDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetUserCreditosCanjeadosDatabase", Erl)
        Resume Next
        
End Function

Public Function GetUserDiasDonadorDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetUserDiasDonadorDatabase_Err
        

        Dim DonadorExpire As Variant

100     DonadorExpire = SanitizeNullValue(GetCuentaValue(CuentaEmail, "donor_expire"), False)
    
102     If Not DonadorExpire Then Exit Function
104     GetUserDiasDonadorDatabase = DateDiff("d", Date, DonadorExpire)

        
        Exit Function

GetUserDiasDonadorDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetUserDiasDonadorDatabase", Erl)
        Resume Next
        
End Function

Public Function GetUserComprasDonadorDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetUserComprasDonadorDatabase_Err
        
100     GetUserComprasDonadorDatabase = GetCuentaValue(CuentaEmail, "donor_purchases")

        
        Exit Function

GetUserComprasDonadorDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetUserComprasDonadorDatabase", Erl)
        Resume Next
        
End Function

Public Function CheckUserExists(name As String) As Boolean
        
        On Error GoTo CheckUserExists_Err
        
100     CheckUserExists = GetUserValue(name, "COUNT(*)") > 0

        
        Exit Function

CheckUserExists_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.CheckUserExists", Erl)
        Resume Next
        
End Function

Public Function CheckCuentaExiste(CuentaEmail As String) As Boolean
        
        On Error GoTo CheckCuentaExiste_Err
        
100     CheckCuentaExiste = GetCuentaValue(CuentaEmail, "COUNT(*)") > 0

        
        Exit Function

CheckCuentaExiste_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.CheckCuentaExiste", Erl)
        Resume Next
        
End Function

Public Function BANCheckDatabase(name As String) As Boolean
        
        On Error GoTo BANCheckDatabase_Err
        
100     BANCheckDatabase = CBool(GetUserValue(name, "is_banned"))

        
        Exit Function

BANCheckDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.BANCheckDatabase", Erl)
        Resume Next
        
End Function

Public Function GetCodigoActivacionDatabase(name As String) As String
        
        On Error GoTo GetCodigoActivacionDatabase_Err
        
100     GetCodigoActivacionDatabase = GetCuentaValue(name, "validate_code")

        
        Exit Function

GetCodigoActivacionDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetCodigoActivacionDatabase", Erl)
        Resume Next
        
End Function

Public Function CheckCuentaActivadaDatabase(name As String) As Boolean
        
        On Error GoTo CheckCuentaActivadaDatabase_Err
        
100     CheckCuentaActivadaDatabase = GetCuentaValue(name, "validated")

        
        Exit Function

CheckCuentaActivadaDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.CheckCuentaActivadaDatabase", Erl)
        Resume Next
        
End Function

Public Function GetEmailDatabase(name As String) As String
        
        On Error GoTo GetEmailDatabase_Err
        
100     GetEmailDatabase = GetCuentaValue(name, "email")

        
        Exit Function

GetEmailDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetEmailDatabase", Erl)
        Resume Next
        
End Function

Public Function GetMacAddressDatabase(CuentaEmail As String) As String
        
        On Error GoTo GetMacAddressDatabase_Err
        
100     GetMacAddressDatabase = GetCuentaValue(CuentaEmail, "mac_address")

        
        Exit Function

GetMacAddressDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetMacAddressDatabase", Erl)
        Resume Next
        
End Function

Public Function GetHDSerialDatabase(CuentaEmail As String) As Long
        
        On Error GoTo GetHDSerialDatabase_Err
        
100     GetHDSerialDatabase = GetCuentaValue(CuentaEmail, "hd_serial")

        
        Exit Function

GetHDSerialDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetHDSerialDatabase", Erl)
        Resume Next
        
End Function

Public Function CheckBanCuentaDatabase(CuentaEmail As String) As Boolean
        
        On Error GoTo CheckBanCuentaDatabase_Err
        
100     CheckBanCuentaDatabase = CBool(GetCuentaValue(CuentaEmail, "is_banned"))

        
        Exit Function

CheckBanCuentaDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.CheckBanCuentaDatabase", Erl)
        Resume Next
        
End Function

Public Function GetMotivoBanCuentaDatabase(CuentaEmail As String) As String
        
        On Error GoTo GetMotivoBanCuentaDatabase_Err
        
100     GetMotivoBanCuentaDatabase = GetCuentaValue(CuentaEmail, "ban_reason")

        
        Exit Function

GetMotivoBanCuentaDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetMotivoBanCuentaDatabase", Erl)
        Resume Next
        
End Function

Public Function GetQuienBanCuentaDatabase(CuentaEmail As String) As String
        
        On Error GoTo GetQuienBanCuentaDatabase_Err
        
100     GetQuienBanCuentaDatabase = GetCuentaValue(CuentaEmail, "banned_by")

        
        Exit Function

GetQuienBanCuentaDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetQuienBanCuentaDatabase", Erl)
        Resume Next
        
End Function

Public Function GetCuentaLogeadaDatabase(CuentaEmail As String) As Boolean
        
        On Error GoTo GetCuentaLogeadaDatabase_Err
        
100     GetCuentaLogeadaDatabase = GetCuentaValue(CuentaEmail, "is_logged")

        
        Exit Function

GetCuentaLogeadaDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetCuentaLogeadaDatabase", Erl)
        Resume Next
        
End Function

Public Function GetUserStatusDatabase(name As String) As Integer
        
        On Error GoTo GetUserStatusDatabase_Err
        
100     GetUserStatusDatabase = GetUserValue(name, "status")

        
        Exit Function

GetUserStatusDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetUserStatusDatabase", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetAccountIDDatabase", Erl)
        Resume Next
        
End Function

Public Sub GetPasswordAndSaltDatabase(CuentaEmail As String, PasswordHash As String, Salt As String)

    On Error GoTo ErrorHandler
    
    Call MakeQuery("SELECT password, salt FROM account WHERE deleted = FALSE AND email = '" & LCase$(CuentaEmail) & "';")

    If QueryData Is Nothing Then Exit Sub
    
    PasswordHash = QueryData!Password
    Salt = QueryData!Salt
    
    Exit Sub
    
ErrorHandler:
    Call LogDatabaseError("Error in GetPasswordAndSaltDatabase. Cuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.description)
    
End Sub

Public Function GetPersonajesCountDatabase(CuentaEmail As String) As Byte

    On Error GoTo ErrorHandler

    Dim Id As Integer

    Id = GetDBValue("account", "id", "email", LCase$(CuentaEmail))
    
    GetPersonajesCountDatabase = GetPersonajesCountByIDDatabase(Id)
    
    Exit Function
    
ErrorHandler:
    Call LogDatabaseError("Error in GetPersonajesCountDatabase. name: " & CuentaEmail & ". " & Err.Number & " - " & Err.description)
    
End Function

Public Function GetPersonajesCountByIDDatabase(ByVal AccountID As Long) As Byte

    On Error GoTo ErrorHandler
    
    Call MakeQuery("SELECT COUNT(*) FROM user WHERE deleted = FALSE AND account_id = " & AccountID & ";")
    
    If QueryData Is Nothing Then Exit Function
    
    GetPersonajesCountByIDDatabase = QueryData.Fields(0).Value
    
    Exit Function
    
ErrorHandler:
    Call LogDatabaseError("Error in GetPersonajesCountByIDDatabase. AccountID: " & AccountID & ". " & Err.Number & " - " & Err.description)
    
End Function

Public Function GetPersonajesCuentaDatabase(ByVal AccountID As Long, Personaje() As PersonajeCuenta) As Byte
        
        On Error GoTo GetPersonajesCuentaDatabase_Err
        

100     Call MakeQuery("SELECT name, head_id, class_id, body_id, pos_map, level, status, helmet_id, shield_id, weapon_id, guild_index, is_dead FROM user WHERE deleted = FALSE AND account_id = " & AccountID & ";")

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

152         If val(QueryData!is_dead) = 1 Then
154             Personaje(i).Cabeza = iCabezaMuerto

            End If
        
156         QueryData.MoveNext
        Next

        
        Exit Function

GetPersonajesCuentaDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.GetPersonajesCuentaDatabase", Erl)
        Resume Next
        
End Function

Public Sub SetUserLoggedDatabase(ByVal Id As Long, ByVal AccountID As Long)
        
        On Error GoTo SetUserLoggedDatabase_Err
        
100     Call SetDBValue("user", "is_logged", 1, "id", Id)
102     Call MakeQuery("UPDATE account SET logged = logged + 1 WHERE id = " & AccountID & ";", True)

        
        Exit Sub

SetUserLoggedDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SetUserLoggedDatabase", Erl)
        Resume Next
        
End Sub

Public Sub ResetLoggedDatabase(ByVal AccountID As Long)
        
        On Error GoTo ResetLoggedDatabase_Err
        
100     Call MakeQuery("UPDATE account SET logged = 0 WHERE id = " & AccountID & ";", True)

        
        Exit Sub

ResetLoggedDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.ResetLoggedDatabase", Erl)
        Resume Next
        
End Sub

Public Sub SetUsersLoggedDatabase(ByVal NumUsers As Long)
        
        On Error GoTo SetUsersLoggedDatabase_Err
        
100     Call MakeQuery("UPDATE statistics SET value = '" & NumUsers & "' WHERE name = 'online';", True)
        
        Exit Sub

SetUsersLoggedDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SetUsersLoggedDatabase", Erl)
        Resume Next
        
End Sub

Public Function LeerRecordUsuariosDatabase() As Long
        
        On Error GoTo LeerRecordUsuariosDatabase_Err
        
100     Call MakeQuery("SELECT value FROM statistics WHERE name = 'record';")

        If QueryData Is Nothing Then Exit Function

        LeerRecordUsuariosDatabase = val(QueryData!Value)

        Exit Function

LeerRecordUsuariosDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.LeerRecordUsuariosDatabase", Erl)
        Resume Next
        
End Function

Public Sub SetRecordUsersDatabase(ByVal Record As Long)
        
        On Error GoTo SetRecordUsersDatabase_Err
        
100     Call MakeQuery("UPDATE statistics SET value = '" & Record & "' WHERE name = 'record';", True)
        
        Exit Sub

SetRecordUsersDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SetRecordUsersDatabase", Erl)
        Resume Next
        
End Sub

Public Sub LogoutAllUsersAndAccounts()
        
        On Error GoTo LogoutAllUsersAndAccounts_Err
        

        Dim Query As String

100     Query = "UPDATE user SET is_logged = FALSE; "
102     Query = Query & "UPDATE account SET logged = 0;"

104     Call MakeQuery(Query, True)

        
        Exit Sub

LogoutAllUsersAndAccounts_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.LogoutAllUsersAndAccounts", Erl)
        Resume Next
        
End Sub

Public Sub SaveBattlePointsDatabase(ByVal Id As Long, ByVal BattlePoints As Long)
        
        On Error GoTo SaveBattlePointsDatabase_Err
        
100     Call SetDBValue("user", "battle_points", BattlePoints, "id", Id)

        
        Exit Sub

SaveBattlePointsDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SaveBattlePointsDatabase", Erl)
        Resume Next
        
End Sub

Public Sub SaveVotoDatabase(ByVal Id As Long, ByVal Encuestas As Integer)
        
        On Error GoTo SaveVotoDatabase_Err
        
100     Call SetUserValueByID(Id, "votes_amount", Encuestas)

        
        Exit Sub

SaveVotoDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SaveVotoDatabase", Erl)
        Resume Next
        
End Sub

Public Sub SaveUserBodyDatabase(UserName As String, ByVal Body As Integer)
        
        On Error GoTo SaveUserBodyDatabase_Err
        
100     Call SetUserValue(UserName, "body_id", Body)

        
        Exit Sub

SaveUserBodyDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SaveUserBodyDatabase", Erl)
        Resume Next
        
End Sub

Public Sub SaveUserHeadDatabase(UserName As String, ByVal Head As Integer)
        
        On Error GoTo SaveUserHeadDatabase_Err
        
100     Call SetUserValue(UserName, "head_id", Head)

        
        Exit Sub

SaveUserHeadDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SaveUserHeadDatabase", Erl)
        Resume Next
        
End Sub

Public Sub SaveUserSkillDatabase(UserName As String, ByVal Skill As Integer, ByVal Value As Integer)
        
        On Error GoTo SaveUserSkillDatabase_Err
        
100     Call MakeQuery("UPDATE skillpoints SET value = " & Value & " WHERE number = " & Skill & " AND user_id = (SELECT id FROM user WHERE name = '" & UserName & "');", True)

        
        Exit Sub

SaveUserSkillDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SaveUserSkillDatabase", Erl)
        Resume Next
        
End Sub

Public Sub SaveUserSkillsLibres(UserName As String, ByVal SkillsLibres As Integer)
        
        On Error GoTo SaveUserHeadDatabase_Err
        
100     Call SetUserValue(UserName, "SaveUserSkillsLibres", SkillsLibres)

        
        Exit Sub

SaveUserHeadDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SaveUserHeadDatabase", Erl)
        Resume Next
        
End Sub

Public Sub SaveNewAccountDatabase(CuentaEmail As String, PasswordHash As String, Salt As String, Codigo As String)

    On Error GoTo ErrorHandler

    Dim q As String

    'Basic user data
    q = "INSERT INTO account SET "
    q = q & "email = '" & LCase$(CuentaEmail) & "', "
    q = q & "password = '" & PasswordHash & "', "
    q = q & "salt = '" & Salt & "', "
    q = q & "validate_code = '" & Codigo & "', "
    q = q & "date_created = NOW();"

    Call MakeQuery(q, True)
    
    Exit Sub
        
ErrorHandler:
    Call LogDatabaseError("Error en SaveNewAccountDatabase. Cuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub ValidarCuentaDatabase(UserCuenta As String)
        
        On Error GoTo ValidarCuentaDatabase_Err
        
100     Call SetCuentaValue(UserCuenta, "validated", 1)

        
        Exit Sub

ValidarCuentaDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.ValidarCuentaDatabase", Erl)
        Resume Next
        
End Sub

Public Sub BorrarUsuarioDatabase(name As String)

    On Error GoTo ErrorHandler
    
    Call MakeQuery("UPDATE user SET name = CONCAT('DELETED_', name), deleted = TRUE WHERE name = '" & name & "';", True)

    Exit Sub
    
ErrorHandler:
    Call LogDatabaseError("Error en BorrarUsuarioDatabase borrando user de la Mysql Database: " & name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub BorrarCuentaDatabase(CuentaEmail As String)

    On Error GoTo ErrorHandler

    Dim Id As Integer

    Id = GetDBValue("account", "id", "email", LCase$(CuentaEmail))

    Dim Query As String
    
    Query = "UPDATE account SET email = CONCAT('DELETED_', email), deleted = TRUE WHERE email = '" & LCase$(CuentaEmail) & "'; "
    
    Query = Query & "UPDATE user SET name = CONCAT('DELETED_', name), deleted = TRUE WHERE account_id = '" & Id & "';"
    
    Call MakeQuery(Query, True)

    Exit Sub
    
ErrorHandler:
    Call LogDatabaseError("Error en BorrarCuentaDatabase borrando user de la Mysql Database: " & CuentaEmail & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveBanDatabase(ByVal UserName As String, ByVal Reason As String, ByVal BannedBy As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim Query As String

    Query = "UPDATE user SET is_banned = TRUE WHERE name = '" & UserName & "'; "

    Query = Query & "INSERT INTO punishment SET "
    Query = Query & "user_id = (SELECT id from user WHERE name = '" & UserName & "'), "
    Query = Query & "number = number + 1, "
    Query = Query & "reason = '" & BannedBy & ": " & LCase$(Reason) & " " & Date & " " & Time & "';"

    Call MakeQuery(Query, True)

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in SaveBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SavePenaDatabase(UserName As String, Reason As String)

    On Error GoTo ErrorHandler

    Dim Query As String

    Query = Query & "INSERT INTO punishment SET "
    Query = Query & "user_id = (SELECT id from user WHERE name = '" & UserName & "'), "
    Query = Query & "number = number + 1, "
    Query = Query & "reason = '" & Reason & "';"

    Call MakeQuery(Query, True)

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in SavePenaDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub UnBanDatabase(UserName As String)

    On Error GoTo ErrorHandler

    Call MakeQuery("UPDATE user SET is_banned = FALSE WHERE name = '" & UserName & "';", True)

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in UnBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub EcharConsejoDatabase(UserName As String)
        
        On Error GoTo EcharConsejoDatabase_Err
        
100     Call MakeQuery("UPDATE user SET pertenece_consejo_real = FALSE, pertenece_consejo_caos = FALSE WHERE name = '" & UserName & "';", True)

        
        Exit Sub

EcharConsejoDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.EcharConsejoDatabase", Erl)
        Resume Next
        
End Sub

Public Sub EcharLegionDatabase(UserName As String)
        
        On Error GoTo EcharLegionDatabase_Err
        
100     Call MakeQuery("UPDATE user SET pertenece_caos = FALSE, reenlistadas = 200 WHERE name = '" & UserName & "';", True)

        
        Exit Sub

EcharLegionDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.EcharLegionDatabase", Erl)
        Resume Next
        
End Sub

Public Sub EcharArmadaDatabase(UserName As String)
        
        On Error GoTo EcharArmadaDatabase_Err
        
100     Call MakeQuery("UPDATE user SET pertenece_real = FALSE, reenlistadas = 200 WHERE name = '" & UserName & "';", True)

        
        Exit Sub

EcharArmadaDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.EcharArmadaDatabase", Erl)
        Resume Next
        
End Sub

Public Sub CambiarPenaDatabase(UserName As String, ByVal Numero As Integer, Pena As String)
        
        On Error GoTo CambiarPenaDatabase_Err
        
100     Call MakeQuery("UPDATE punishment SET reason = '" & Pena & "' WHERE number = " & Numero & " AND user_id = (SELECT id from user WHERE name = '" & UserName & "');", True)

        
        Exit Sub

CambiarPenaDatabase_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.CambiarPenaDatabase", Erl)
        Resume Next
        
End Sub

Public Function GetUserAmountOfPunishmentsDatabase(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Call MakeQuery("SELECT COUNT(*) as punishments FROM punishment WHERE user_id = (SELECT id from user WHERE name = '" & UserName & "');")

    If QueryData Is Nothing Then Exit Function

    GetUserAmountOfPunishmentsDatabase = QueryData!punishments

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserAmountOfPunishmentsDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetNombreCuentaDatabase(name As String) As String

    On Error GoTo ErrorHandler

    'Hacemos la query.
    Call MakeQuery("SELECT email FROM account WHERE id = (SELECT account_id FROM user WHERE name = '" & name & "');")
    
    'Verificamos que la query no devuelva un resultado vacio.
    If QueryData Is Nothing Then Exit Function
    
    'Obtenemos el nombre de la cuenta
    GetNombreCuentaDatabase = QueryData!name

    Exit Function
    
ErrorHandler:
    Call LogDatabaseError("Error en GetNombreCuentaDatabase leyendo user de la Mysql Database: " & name & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildIndexDatabase(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 09/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Call MakeQuery("SELECT guild_index FROM user WHERE UPPER(name) = '" & UCase$(UserName) & "';")

    If QueryData Is Nothing Then Exit Function

    GetUserGuildIndexDatabase = SanitizeNullValue(QueryData!Guild_Index, 0)

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildMemberDatabase(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Call MakeQuery("SELECT guild_member_history FROM user WHERE name = '" & UserName & "';")

    If QueryData Is Nothing Then Exit Function

    GetUserGuildMemberDatabase = SanitizeNullValue(QueryData!guild_member_history, vbNullString)

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildAspirantDatabase(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Call MakeQuery("SELECT guild_aspirant_index FROM user WHERE name = '" & UserName & "';")

    If QueryData Is Nothing Then Exit Function

    GetUserGuildAspirantDatabase = SanitizeNullValue(QueryData!guild_aspirant_index, 0)

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildRejectionReasonDatabase(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler
    
    Call MakeQuery("SELECT guild_rejected_because FROM user WHERE name = '" & UserName & "';")

    If QueryData Is Nothing Then Exit Function

    GetUserGuildRejectionReasonDatabase = SanitizeNullValue(QueryData!guild_rejected_because, vbNullString)

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildPedidosDatabase(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Call MakeQuery("SELECT guild_requests_history FROM user WHERE name = '" & UserName & "';")

    If QueryData Is Nothing Then Exit Function

    GetUserGuildPedidosDatabase = SanitizeNullValue(QueryData!guild_requests_history, vbNullString)

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserGuildRejectionReasonDatabase(ByVal UserName As String, ByVal Reason As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim Query As String

    Query = "UPDATE user SET "
    Query = Query & "guild_rejected_because = '" & Reason & "' "
    Query = Query & "WHERE name = '" & UserName & "';"

    Call MakeQuery(Query, True)

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildIndexDatabase(ByVal UserName As String, ByVal GuildIndex As Integer)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim Query As String

    Query = "UPDATE user SET "
    Query = Query & "guild_index = " & GuildIndex & " "
    Query = Query & "WHERE name = '" & UserName & "';"
    
    Call MakeQuery(Query, True)

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildAspirantDatabase(ByVal UserName As String, ByVal AspirantIndex As Integer)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim Query As String

    Query = "UPDATE user SET "
    Query = Query & "guild_aspirant_index = " & AspirantIndex & " "
    Query = Query & "WHERE name = '" & UserName & "';"

    Call MakeQuery(Query, True)

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildMemberDatabase(ByVal UserName As String, ByVal guilds As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim Query As String

    Query = "UPDATE user SET "
    Query = Query & "guild_member_history = '" & guilds & "' "
    Query = Query & "WHERE name = '" & UserName & "';"

    Call MakeQuery(Query, True)

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildPedidosDatabase(ByVal UserName As String, ByVal Pedidos As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim Query As String

    Query = "UPDATE user SET "
    Query = Query & "guild_requests_history = '" & Pedidos & "' "
    Query = Query & "WHERE name = '" & UserName & "';"

    Call MakeQuery(Query, True)

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

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

    Call MakeQuery("SELECT race_id, class_id, genre_id, level, gold, bank_gold, guild_requests_history, guild_index, guild_member_history, pertenece_real, pertenece_caos, ciudadanos_matados, criminales_matados FROM user WHERE name = '" & UserName & "';")

    If QueryData Is Nothing Then
        Call WriteConsoleMsg(UserIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    ' Get the character's current guild
    GuildActual = SanitizeNullValue(QueryData!Guild_Index, 0)

    If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
        gName = "<" & GuildName(GuildActual) & ">"
    Else
        gName = "Ninguno"

    End If

    'Get previous guilds
    Miembro = SanitizeNullValue(QueryData!guild_member_history, vbNullString)

    If Len(Miembro) > 400 Then
        Miembro = ".." & Right$(Miembro, 400)

    End If

    Call Protocol.WriteCharacterInfo(UserIndex, UserName, QueryData!race_id, QueryData!class_id, QueryData!genre_id, QueryData!level, QueryData!gold, QueryData!bank_gold, SanitizeNullValue(QueryData!guild_requests_history, vbNullString), gName, Miembro, QueryData!pertenece_real, QueryData!pertenece_caos, QueryData!ciudadanos_matados, QueryData!criminales_matados)

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendCharacterInfoDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function EnterAccountDatabase(ByVal UserIndex As Integer, CuentaEmail As String, Password As String, MacAddress As String, ByVal HDserial As Long, ip As String) As Boolean

    On Error GoTo ErrorHandler
    
    Call MakeQuery("SELECT id, password, salt, validated, is_banned, ban_reason, banned_by FROM account WHERE email = '" & LCase$(CuentaEmail) & "';")
    
    If Database_Connection.State = adStateClosed Then
        Call WriteShowMessageBox(UserIndex, "Ha ocurrido un error interno en el servidor. ¡Estamos tratando de resolverlo!")
        Exit Function
    End If
    
    If QueryData Is Nothing Then
        Call WriteShowMessageBox(UserIndex, "La cuenta no existe.")
        Exit Function
    End If
    
    If val(QueryData!is_banned) > 0 Then
        Call WriteShowMessageBox(UserIndex, "La cuenta se encuentra baneada debido a: " & QueryData!ban_reason & ". Esta decisión fue tomada por: " & QueryData!banned_by & ".")
        Exit Function
    End If
    
    If Not PasswordValida(Password, QueryData!Password, QueryData!Salt) Then
        Call WriteShowMessageBox(UserIndex, "Contraseña inválida.")
        Exit Function
    End If
    
    If val(QueryData!validated) = 0 Then
        Call WriteShowMessageBox(UserIndex, "¡La cuenta no ha sido validada aún!")
        Exit Function
    End If
    
    UserList(UserIndex).AccountID = QueryData!Id
    UserList(UserIndex).Cuenta = CuentaEmail
    
    Call MakeQuery("UPDATE account SET mac_address = '" & MacAddress & "', hd_serial = " & HDserial & ", last_ip = '" & ip & "', last_access = NOW() WHERE id = " & QueryData!Id & ";", True)
    
    EnterAccountDatabase = True
    
    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in EnterAccountDatabase. UserCuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub ChangePasswordDatabase(ByVal UserIndex As Integer, OldPassword As String, NewPassword As String)

    On Error GoTo ErrorHandler

    If LenB(NewPassword) = 0 Then
        Call WriteConsoleMsg(UserIndex, "Debe especificar una contraseña nueva, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    Call MakeQuery("SELECT password, salt FROM account WHERE id = " & UserList(UserIndex).AccountID & ";")
    
    If QueryData Is Nothing Then
        Call WriteConsoleMsg(UserIndex, "No se ha podido cambiar la contraseña por un error interno. Avise a un administrador.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    If Not PasswordValida(OldPassword, QueryData!Password, QueryData!Salt) Then
        Call WriteConsoleMsg(UserIndex, "La contraseña actual proporcionada no es correcta. La contraseña no ha sido cambiada, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    Dim Salt As String * 10

    Salt = RandomString(10) ' Alfanumerico
    
    Dim oSHA256 As CSHA256

    Set oSHA256 = New CSHA256

    Dim PasswordHash As String * 64

    PasswordHash = oSHA256.SHA256(NewPassword & Salt)
    
    Set oSHA256 = Nothing
    
    Call MakeQuery("UPDATE account SET password = '" & PasswordHash & "', salt = '" & Salt & "' WHERE id = " & UserList(UserIndex).AccountID & ";", True)
    
    Call WriteConsoleMsg(UserIndex, "La contraseña de su cuenta fue cambiada con éxito.", FontTypeNames.FONTTYPE_INFO)
    
    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in ChangePasswordDatabase. Username: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUsersLoggedAccountDatabase(ByVal AccountID As Integer) As Byte

    On Error GoTo ErrorHandler

    Call GetDBValue("account", "logged", "id", AccountID)
    
    If QueryData Is Nothing Then Exit Function
    
    GetUsersLoggedAccountDatabase = val(QueryData!logged)

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUsersLoggedAccountDatabase. AccountID: " & AccountID & ". " & Err.Number & " - " & Err.description)

End Function

Public Function SetPositionDatabase(UserName As String, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    On Error GoTo ErrorHandler

    Call MakeQuery("UPDATE user SET pos_map = " & Map & ", pos_x = " & X & ", pos_y = " & X & " WHERE UPPER(name) = '" & UCase$(UserName) & "';", True)
    
    SetPositionDatabase = True

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in SetPositionDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function AddOroBancoDatabase(UserName As String, ByVal OroGanado As Long) As Boolean
    On Error GoTo ErrorHandler

    Call MakeQuery("UPDATE user SET bank_gold = bank_gold + " & OroGanado & " WHERE UPPER(name) = '" & UCase$(UserName) & "';", True)
    
    AddOroBancoDatabase = True

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in AddOroBancoDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function DarLlaveAUsuarioDatabase(UserName As String, ByVal LlaveObj As Integer) As Boolean
    On Error GoTo ErrorHandler

    Call MakeQuery("INSERT INTO house_key SET key_obj = " & LlaveObj & ", account_id = (SELECT account_id FROM user WHERE UPPER(name) = '" & UCase$(UserName) & "');", True)
    
    DarLlaveAUsuarioDatabase = True

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in DarLlaveAUsuarioDatabase. UserName: " & UserName & ", LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.description)

End Function

Public Function DarLlaveACuentaDatabase(email As String, ByVal LlaveObj As Integer) As Boolean
    On Error GoTo ErrorHandler

    Call MakeQuery("INSERT INTO house_key SET key_obj = " & LlaveObj & ", account_id = (SELECT id FROM account WHERE UPPER(email) = '" & UCase$(email) & "');", True)
    
    DarLlaveACuentaDatabase = True

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in DarLlaveACuentaDatabase. Email: " & email & ", LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.description)

End Function

Public Function SacarLlaveDatabase(ByVal LlaveObj As Integer) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim UserCount As Integer
    Dim Users() As String

    ' Obtengo los usuarios logueados en la cuenta del dueño de la llave
    Call MakeQuery("SELECT name FROM user WHERE is_logged = TRUE AND account_id = (SELECT account_id FROM house_key WHERE key_obj = " & LlaveObj & ");")
    
    If QueryData Is Nothing Then Exit Function

    ' Los almaceno en un array
    UserCount = QueryData.RecordCount
    
    ReDim Users(1 To UserCount) As String
    
    QueryData.MoveFirst

    i = 1

    While Not QueryData.EOF
    
        Users(i) = QueryData!name
        i = i + 1

        QueryData.MoveNext
    Wend
    
    ' Intento borrar la llave de la db
    Call MakeQuery("DELETE FROM house_key WHERE key_obj = " & LlaveObj & ";", True)
    
    ' Si pudimos borrar, actualizamos los usuarios logueados
    Dim UserIndex As Integer
    
    For i = 1 To UserCount
        UserIndex = NameIndex(Users(i))
        
        If UserIndex <> 0 Then
            Call SacarLlaveDeLLavero(UserIndex, LlaveObj)
        End If
    Next
    
    SacarLlaveDatabase = True

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in SacarLlaveDatabase. LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub VerLlavesDatabase(ByVal UserIndex As Integer)
    On Error GoTo ErrorHandler

    Call MakeQuery("SELECT (SELECT email FROM account WHERE id = K.account_id) as email, key_obj FROM house_key AS K;")

    If QueryData Is Nothing Then
        Call WriteConsoleMsg(UserIndex, "No hay llaves otorgadas por el momento.", FontTypeNames.FONTTYPE_INFO)

    ElseIf QueryData.RecordCount = 0 Then
        Call WriteConsoleMsg(UserIndex, "No hay llaves otorgadas por el momento.", FontTypeNames.FONTTYPE_INFO)
    
    Else
        Dim message As String
        
        message = "Llaves usadas: " & QueryData.RecordCount & vbNewLine
    
        QueryData.MoveFirst

        While Not QueryData.EOF
        
            message = message & "Llave: " & QueryData!key_obj & " - Cuenta: " & QueryData!email & vbNewLine

            QueryData.MoveNext
        Wend
        
        message = Left$(message, Len(message) - 2)
        
        Call WriteConsoleMsg(UserIndex, message, FontTypeNames.FONTTYPE_INFO)
    End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in VerLlavesDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function SanitizeNullValue(ByVal Value As Variant, ByVal defaultValue As Variant) As Variant
        
        On Error GoTo SanitizeNullValue_Err
        
100     SanitizeNullValue = IIf(IsNull(Value), defaultValue, Value)

        
        Exit Function

SanitizeNullValue_Err:
        Call RegistrarError(Err.Number, Err.description, "modDatabase.SanitizeNullValue", Erl)
        Resume Next
        
End Function

