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
    
        Database_Connection.ConnectionString = "DRIVER={MySQL ODBC 8.0 ANSI Driver};" & "SERVER=" & Database_Host & ";" & "DATABASE=" & Database_Name & ";" & "USER=" & Database_Username & ";" & "PASSWORD=" & Database_Password & ";" & "OPTION=3;MULTI_STATEMENTS=1"

    End If
    
    Debug.Print Database_Connection.ConnectionString
    
    Database_Connection.CursorLocation = adUseClient
    Database_Connection.Open

    Exit Sub
    
ErrorHandler:
    Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.description)

End Sub

Public Sub Database_Close()

    '***************************************************
    'Author: Juan Andres Dalmasso
    'Last Modification: 18/09/2018
    'Nota by WyroX: Cerrar la conexion tambien libera
    'los recursos y cierra los RecordSet generados
    '***************************************************
    On Error GoTo ErrorHandler
     
    Database_Connection.Close
    Set Database_Connection = Nothing
     
    Exit Sub
     
ErrorHandler:
    Call LogDatabaseError("Unable to close Mysql Database: " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveNewUserDatabase(ByVal UserIndex As Integer)

    On Error GoTo ErrorHandler
    
    Dim q As String
    
    With UserList(UserIndex)
        'Basic user data
        q = "INSERT INTO user SET "
        q = q & "name = '" & .name & "', "
        q = q & "account_id = " & .AccountID & ", "
        q = q & "level = " & .Stats.ELV & ", "
        q = q & "exp = " & .Stats.Exp & ", "
        q = q & "elu = " & .Stats.ELU & ", "
        q = q & "genre_id = " & .genero & ", "
        q = q & "race_id = " & .raza & ", "
        q = q & "class_id = " & .clase & ", "
        q = q & "home_id = " & .Hogar & ", "
        q = q & "description = '" & .Desc & "', "
        q = q & "gold = " & .Stats.GLD & ", "
        q = q & "free_skillpoints = " & .Stats.SkillPts & ", "
        'Q = Q & "assigned_skillpoints = " & .Counters.AsignedSkills & ", "
        q = q & "pos_map = " & .Pos.Map & ", "
        q = q & "pos_x = " & .Pos.x & ", "
        q = q & "pos_y = " & .Pos.Y & ", "
        q = q & "body_id = " & .Char.Body & ", "
        q = q & "head_id = " & .Char.Head & ", "
        q = q & "weapon_id = " & .Char.WeaponAnim & ", "
        q = q & "helmet_id = " & .Char.CascoAnim & ", "
        q = q & "shield_id = " & .Char.ShieldAnim & ", "
        q = q & "items_Amount = " & .Invent.NroItems & ", "
        q = q & "slot_armour = " & .Invent.ArmourEqpSlot & ", "
        q = q & "slot_weapon = " & .Invent.WeaponEqpSlot & ", "
        q = q & "slot_shield = " & .Invent.EscudoEqpSlot & ", "
        q = q & "slot_helmet = " & .Invent.CascoEqpSlot & ", "
        q = q & "slot_ammo = " & .Invent.MunicionEqpSlot & ", "
        q = q & "slot_ring = " & .Invent.AnilloEqpSlot & ", "
        q = q & "slot_tool = " & .Invent.HerramientaEqpSlot & ", "
        q = q & "slot_magic = " & .Invent.MagicoSlot & ", "
        q = q & "slot_knuckles = " & .Invent.NudilloSlot & ", "
        q = q & "slot_ship = " & .Invent.BarcoSlot & ", "
        q = q & "slot_mount = " & .Invent.MonturaSlot & ", "
        q = q & "min_hp = " & .Stats.MinHp & ", "
        q = q & "max_hp = " & .Stats.MaxHp & ", "
        q = q & "min_man = " & .Stats.MinMAN & ", "
        q = q & "max_man = " & .Stats.MaxMAN & ", "
        q = q & "min_sta = " & .Stats.MinSta & ", "
        q = q & "max_sta = " & .Stats.MaxSta & ", "
        q = q & "min_ham = " & .Stats.MinHam & ", "
        q = q & "max_ham = " & .Stats.MaxHam & ", "
        q = q & "min_sed = " & .Stats.MinAGU & ", "
        q = q & "max_sed = " & .Stats.MaxAGU & ", "
        q = q & "min_hit = " & .Stats.MinHIT & ", "
        q = q & "max_hit = " & .Stats.MaxHit & ", "
        'Q = Q & "rep_noble = " & .NobleRep & ", "
        'Q = Q & "rep_plebe = " & .Reputacion.PlebeRep & ", "
        'Q = Q & "rep_average = " & .Reputacion.Promedio & ", "
        q = q & "is_naked = " & .flags.Desnudo & ", "
        q = q & "status = " & .Faccion.Status & ", "
        q = q & "is_logged = TRUE; "
        
        Call MakeQuery(q, True)

        ' Para recibir el ID del user
        Call MakeQuery("SELECT LAST_INSERT_ID();")

        If QueryData Is Nothing Then
            .Id = 1
        Else
            .Id = val(QueryData.Fields(0).Value)

        End If
        
        ' Comenzamos una cadena de queries (para enviar todo de una)
        Dim LoopC As Integer

        'User attributes
        q = "INSERT INTO attribute (user_id, number, value) VALUES "

        For LoopC = 1 To NUMATRIBUTOS
            q = q & "("
            q = q & .Id & ", "
            q = q & LoopC & ", "
            q = q & .Stats.UserAtributos(LoopC) & ")"

            If LoopC < NUMATRIBUTOS Then
                q = q & ", "
            Else
                q = q & "; "

            End If

        Next LoopC

        'User spells
        q = q & "INSERT INTO spell (user_id, number, spell_id) VALUES "

        For LoopC = 1 To MAXUSERHECHIZOS
            q = q & "("
            q = q & .Id & ", "
            q = q & LoopC & ", "
            q = q & .Stats.UserHechizos(LoopC) & ")"

            If LoopC < MAXUSERHECHIZOS Then
                q = q & ", "
            Else
                q = q & "; "

            End If

        Next LoopC

        'User inventory
        q = q & "INSERT INTO inventory_item (user_id, number, item_id, Amount, is_equipped) VALUES "

        For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
            q = q & "("
            q = q & .Id & ", "
            q = q & LoopC & ", "
            q = q & .Invent.Object(LoopC).ObjIndex & ", "
            q = q & .Invent.Object(LoopC).Amount & ", "
            q = q & .Invent.Object(LoopC).Equipped & ")"

            If LoopC < UserList(UserIndex).CurrentInventorySlots Then
                q = q & ", "
            Else
                q = q & "; "

            End If

        Next LoopC

        'User skills
        'Q = Q & "INSERT INTO skillpoint (user_id, number, value, exp, elu) VALUES "
        q = q & "INSERT INTO skillpoint (user_id, number, value) VALUES "

        For LoopC = 1 To NUMSKILLS
            q = q & "("
            q = q & .Id & ", "
            q = q & LoopC & ", "
            q = q & .Stats.UserSkills(LoopC) & ")"
            'Q = Q & .Stats.UserSkills(LoopC) & ", "
            'Q = Q & .Stats.ExpSkills(LoopC) & ", "
            'Q = Q & .Stats.EluSkills(LoopC) & ")"

            If LoopC < NUMSKILLS Then
                q = q & ", "
            Else
                q = q & "; "

            End If

        Next LoopC
        
        'User friends
        'Q = Q & "INSERT INTO friend (user_id, number) VALUES "

        'For LoopC = 1 To MAXAMIGOS
        '    Q = Q & "("
        '    Q = Q & .ID & ", "
        '    Q = Q & LoopC & ")"

        '    If LoopC < MAXAMIGOS Then
        '        Q = Q & ", "
        '    Else
        '        Q = Q & "; "

        '    End If
        'Next LoopC
        
        'User quests
        q = q & "INSERT INTO quest (user_id, number) VALUES "

        For LoopC = 1 To MAXUSERQUESTS
            q = q & "("
            q = q & .Id & ", "
            q = q & LoopC & ")"

            If LoopC < MAXUSERQUESTS Then
                q = q & ", "
            Else
                q = q & "; "

            End If

        Next LoopC

        'Enviamos todas las queries
        Call MakeQuery(q, True)

    End With

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error en SaveNewUserDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserDatabase(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean = False)

    On Error GoTo ErrorHandler
    
    Dim q As String

    'Basic user data
    With UserList(UserIndex)
        q = "UPDATE user SET "
        q = q & "name = '" & .name & "', "
        q = q & "level = " & .Stats.ELV & ", "
        q = q & "exp = " & CLng(.Stats.Exp) & ", "
        q = q & "elu = " & .Stats.ELU & ", "
        q = q & "genre_id = " & .genero & ", "
        q = q & "race_id = " & .raza & ", "
        q = q & "class_id = " & .clase & ", "
        q = q & "home_id = " & .Hogar & ", "
        q = q & "description = '" & .Desc & "', "
        q = q & "gold = " & .Stats.GLD & ", "
        q = q & "bank_gold = " & .Stats.Banco & ", "
        q = q & "free_skillpoints = " & .Stats.SkillPts & ", "
        'Q = Q & "assigned_skillpoints = " & .Counters.AsignedSkills & ", "
        'Q = Q & "pet_Amount = " & .NroMascotas & ", "
        q = q & "pos_map = " & .Pos.Map & ", "
        q = q & "pos_x = " & .Pos.x & ", "
        q = q & "pos_y = " & .Pos.Y & ", "
        q = q & "message_info = '" & .MENSAJEINFORMACION & "', "
        q = q & "body_id = " & .Char.Body & ", "
        q = q & "head_id = " & .OrigChar.Head & ", "
        q = q & "weapon_id = " & .Char.WeaponAnim & ", "
        q = q & "helmet_id = " & .Char.CascoAnim & ", "
        q = q & "shield_id = " & .Char.ShieldAnim & ", "
        q = q & "heading = " & .Char.heading & ", "
        q = q & "items_Amount = " & .Invent.NroItems & ", "
        q = q & "slot_armour = " & .Invent.ArmourEqpSlot & ", "
        q = q & "slot_weapon = " & .Invent.WeaponEqpSlot & ", "
        q = q & "slot_shield = " & .Invent.EscudoEqpSlot & ", "
        q = q & "slot_helmet = " & .Invent.CascoEqpSlot & ", "
        q = q & "slot_ammo = " & .Invent.MunicionEqpSlot & ", "
        q = q & "slot_ring = " & .Invent.AnilloEqpSlot & ", "
        q = q & "slot_tool = " & .Invent.HerramientaEqpSlot & ", "
        q = q & "slot_magic = " & .Invent.MagicoSlot & ", "
        q = q & "slot_knuckles = " & .Invent.NudilloSlot & ", "
        q = q & "slot_ship = " & .Invent.BarcoSlot & ", "
        q = q & "slot_mount = " & .Invent.MonturaSlot & ", "
        q = q & "min_hp = " & .Stats.MinHp & ", "
        q = q & "max_hp = " & .Stats.MaxHp & ", "
        q = q & "min_man = " & .Stats.MinMAN & ", "
        q = q & "max_man = " & .Stats.MaxMAN & ", "
        q = q & "min_sta = " & .Stats.MinSta & ", "
        q = q & "max_sta = " & .Stats.MaxSta & ", "
        q = q & "min_ham = " & .Stats.MinHam & ", "
        q = q & "max_ham = " & .Stats.MaxHam & ", "
        q = q & "min_sed = " & .Stats.MinAGU & ", "
        q = q & "max_sed = " & .Stats.MaxAGU & ", "
        q = q & "min_hit = " & .Stats.MinHIT & ", "
        q = q & "max_hit = " & .Stats.MaxHit & ", "
        q = q & "killed_npcs = " & .Stats.NPCsMuertos & ", "
        q = q & "killed_users = " & .Stats.UsuariosMatados & ", "
        q = q & "invent_level = " & .Stats.InventLevel & ", "
        'Q = Q & "rep_asesino = " & .Reputacion.AsesinoRep & ", "
        'Q = Q & "rep_bandido = " & .Reputacion.BandidoRep & ", "
        'Q = Q & "rep_burgues = " & .Reputacion.BurguesRep & ", "
        'Q = Q & "rep_ladron = " & .Reputacion.LadronesRep & ", "
        'Q = Q & "rep_noble = " & .Reputacion.NobleRep & ", "
        'Q = Q & "rep_plebe = " & .Reputacion.PlebeRep & ", "
        'Q = Q & "rep_average = " & .Reputacion.Promedio & ", "
        q = q & "is_naked = " & .flags.Desnudo & ", "
        q = q & "is_poisoned = " & .flags.Envenenado & ", "
        q = q & "is_hidden = " & .flags.Escondido & ", "
        q = q & "is_hungry = " & .flags.Hambre & ", "
        q = q & "is_thirsty = " & .flags.Sed & ", "
        'Q = Q & "is_banned = " & .flags.Ban & ", " Esto es innecesario porque se setea cuando lo baneas (creo)
        q = q & "is_dead = " & .flags.Muerto & ", "
        q = q & "is_sailing = " & .flags.Navegando & ", "
        q = q & "is_paralyzed = " & .flags.Paralizado & ", "
        q = q & "is_mounted = " & .flags.Montado & ", "
        q = q & "is_silenced = " & .flags.Silenciado & ", "
        q = q & "silence_minutes_left = " & .flags.MinutosRestantes & ", "
        q = q & "silence_elapsed_seconds = " & .flags.SegundosPasados & ", "
        q = q & "spouse = '" & .flags.Pareja & "', "
        q = q & "counter_pena = " & .Counters.Pena & ", "
        q = q & "deaths = " & .flags.VecesQueMoriste & ", "
        q = q & "pertenece_consejo_real = " & (.flags.Privilegios And PlayerType.RoyalCouncil) & ", "
        q = q & "pertenece_consejo_caos = " & (.flags.Privilegios And PlayerType.ChaosCouncil) & ", "
        q = q & "pertenece_real = " & .Faccion.ArmadaReal & ", "
        q = q & "pertenece_caos = " & .Faccion.FuerzasCaos & ", "
        q = q & "ciudadanos_matados = " & .Faccion.CiudadanosMatados & ", "
        q = q & "criminales_matados = " & .Faccion.CriminalesMatados & ", "
        q = q & "recibio_armadura_real = " & .Faccion.RecibioArmaduraReal & ", "
        q = q & "recibio_armadura_caos = " & .Faccion.RecibioArmaduraCaos & ", "
        q = q & "recibio_exp_real = " & .Faccion.RecibioExpInicialReal & ", "
        q = q & "recibio_exp_caos = " & .Faccion.RecibioExpInicialCaos & ", "
        q = q & "recompensas_real = " & .Faccion.RecompensasReal & ", "
        q = q & "recompensas_caos = " & .Faccion.RecompensasCaos & ", "
        q = q & "reenlistadas = " & .Faccion.Reenlistadas & ", "
        q = q & "fecha_ingreso = " & IIf(.Faccion.FechaIngreso <> vbNullString, "'" & .Faccion.FechaIngreso & "'", "NULL") & ", "
        q = q & "nivel_ingreso = " & .Faccion.NivelIngreso & ", "
        q = q & "matados_ingreso = " & .Faccion.MatadosIngreso & ", "
        q = q & "siguiente_recompensa = " & .Faccion.NextRecompensa & ", "
        q = q & "status = " & .Faccion.Status & ", "
        q = q & "battle_points = " & .flags.BattlePuntos & ", "
        q = q & "guild_index = " & .GuildIndex & ", "
        q = q & "chat_combate = " & .ChatCombate & ", "
        q = q & "chat_global = " & .ChatGlobal & ", "
        q = q & "is_logged = " & IIf(Logout, "FALSE", "TRUE")
        q = q & " WHERE id = " & .Id & "; "
        
        Dim LoopC As Integer

        'User attributes
        q = q & "INSERT INTO attribute (user_id, number, value) VALUES "

        For LoopC = 1 To NUMATRIBUTOS
            q = q & "("
            q = q & .Id & ", "
            q = q & LoopC & ", "
            q = q & .Stats.UserAtributosBackUP(LoopC) & ")"

            If LoopC < NUMATRIBUTOS Then
                q = q & ", "

            End If

        Next LoopC
        
        q = q & " ON DUPLICATE KEY UPDATE value=VALUES(value); "

        'User spells
        q = q & "INSERT INTO spell (user_id, number, spell_id) VALUES "

        For LoopC = 1 To MAXUSERHECHIZOS
            q = q & "("
            q = q & .Id & ", "
            q = q & LoopC & ", "
            q = q & .Stats.UserHechizos(LoopC) & ")"

            If LoopC < MAXUSERHECHIZOS Then
                q = q & ", "

            End If

        Next LoopC
        
        q = q & " ON DUPLICATE KEY UPDATE spell_id=VALUES(spell_id); "

        'User inventory
        q = q & "INSERT INTO inventory_item (user_id, number, item_id, Amount, is_equipped) VALUES "

        For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
            q = q & "("
            q = q & .Id & ", "
            q = q & LoopC & ", "
            q = q & .Invent.Object(LoopC).ObjIndex & ", "
            q = q & .Invent.Object(LoopC).Amount & ", "
            q = q & .Invent.Object(LoopC).Equipped & ")"

            If LoopC < UserList(UserIndex).CurrentInventorySlots Then
                q = q & ", "

            End If

        Next LoopC
        
        q = q & " ON DUPLICATE KEY UPDATE item_id=VALUES(item_id), amount=VALUES(Amount), is_equipped=VALUES(is_equipped); "

        'User bank inventory
        q = q & "INSERT INTO bank_item (user_id, number, item_id, Amount) VALUES "

        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            q = q & "("
            q = q & .Id & ", "
            q = q & LoopC & ", "
            q = q & .BancoInvent.Object(LoopC).ObjIndex & ", "
            q = q & .BancoInvent.Object(LoopC).Amount & ")"

            If LoopC < MAX_BANCOINVENTORY_SLOTS Then
                q = q & ", "

            End If

        Next LoopC
        
        q = q & " ON DUPLICATE KEY UPDATE item_id=VALUES(item_id), amount=VALUES(Amount); "

        'User skills
        'Q = Q & "INSERT INTO skillpoint (user_id, number, value, exp, elu) VALUES "
        q = q & "INSERT INTO skillpoint (user_id, number, value) VALUES "

        For LoopC = 1 To NUMSKILLS
            q = q & "("
            q = q & .Id & ", "
            q = q & LoopC & ", "
            q = q & .Stats.UserSkills(LoopC) & ")"
            'Q = Q & .Stats.UserSkills(LoopC) & ", "
            'Q  = Q & .Stats.ExpSkills(LoopC) & ", "
            'Q = Q & .Stats.EluSkills(LoopC) & ")"

            If LoopC < NUMSKILLS Then
                q = q & ", "

            End If

        Next LoopC
        
        'Q = Q & " ON DUPLICATE KEY UPDATE value=VALUES(value), exp=VALUES(exp), elu=VALUES(elu); "
        q = q & " ON DUPLICATE KEY UPDATE value=VALUES(value); "

        'User pets
        'Dim petType As Integer
        
        'Q = Q & "INSERT INTO pet (user_id, number, pet_id) VALUES "

        'For LoopC = 1 To MAXMASCOTAS
        '    Q = Q & "("
        '    Q = Q & .ID & ", "
        '    Q = Q & LoopC & ", "

        'CHOTS | I got this logic from SaveUserToCharfile
        '    If .MascotasIndex(LoopC) > 0 Then
        '        If Npclist(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
        '            petType = .MascotasType(LoopC)
        '        Else
        '            petType = 0

        '        End If

        '    Else
        '        petType = .MascotasType(LoopC)

        '    End If

        '    Q = Q & petType & ")"

        '    If LoopC < MAXMASCOTAS Then
        '        Q = Q & ", "
        '    End If

        'Next LoopC

        'Q = Q & " ON DUPLICATE KEY UPDATE pet_id=VALUES(pet_id); "
        
        'User friends
        'Q = "INSERT INTO friend (user_id, number, friend, ignored) VALUES "

        'For LoopC = 1 To MAXAMIGOS
        '    Q = Q & "("
        '    Q = Q & .ID & ", "
        '    Q = Q & LoopC & ", "
        '    Q = Q & "'" & .Amigos(LoopC).Nombre & "', "
        '    Q = Q & .Amigos(LoopC).Ignorado & ")"

        '    If LoopC < MAXAMIGOS Then
        '        Q = Q & ", "
        '    Else
        '        Q = Q & ";"

        '    End If
        'Next LoopC
       
        'Agrego ip del user
        q = q & "INSERT INTO connection (user_id, ip, date_last_login) VALUES ("
        q = q & .Id & ", "
        q = q & "'" & .ip & "', "
        q = q & "NOW()) "
        q = q & "ON DUPLICATE KEY UPDATE "
        q = q & "date_last_login = VALUES(date_last_login); "
        
        'Borro la mas vieja si hay mas de 5 (WyroX: si alguien sabe una forma mejor de hacerlo me avisa)
        q = q & "DELETE FROM connection WHERE"
        q = q & " user_id = " & .Id
        q = q & " AND date_last_login < (SELECT min(date_last_login) FROM (SELECT date_last_login FROM connection WHERE"
        q = q & " user_id = " & .Id
        q = q & " ORDER BY date_last_login DESC LIMIT 5) AS d); "
        
        'User quests
        q = q & "INSERT INTO quest (user_id, number, quest_id, npcs) VALUES "
        
        Dim Tmp As Integer, LoopK As Integer

        For LoopC = 1 To MAXUSERQUESTS
            q = q & "("
            q = q & .Id & ", "
            q = q & LoopC & ", "
            q = q & .QuestStats.Quests(LoopC).QuestIndex & ", '"
            
            If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
                Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredNPCs

                If Tmp Then

                    For LoopK = 1 To Tmp
                        q = q & .QuestStats.Quests(LoopC).NPCsKilled(LoopK)
                        
                        If LoopK < Tmp Then
                            q = q & "-"

                        End If

                    Next LoopK

                End If

            End If
            
            q = q & "')"

            If LoopC < MAXUSERQUESTS Then
                q = q & ", "

            End If

        Next LoopC
        
        q = q & " ON DUPLICATE KEY UPDATE quest_id=VALUES(quest_id), npcs=VALUES(npcs); "
        
        'User completed quests
        If .QuestStats.NumQuestsDone > 0 Then
            q = q & "INSERT INTO quest_done (user_id, quest_id) VALUES "
    
            For LoopC = 1 To .QuestStats.NumQuestsDone
                q = q & "("
                q = q & .Id & ", "
                q = q & .QuestStats.QuestsDone(LoopC) & ")"
    
                If LoopC < .QuestStats.NumQuestsDone Then
                    q = q & ", "

                End If
    
            Next LoopC
            
            q = q & " ON DUPLICATE KEY UPDATE quest_id=VALUES(quest_id); "

        End If
        
        'User mail
        'TODO:
        
        ' Si deslogueó, actualizo la cuenta
        If Logout Then
            q = q & "UPDATE account SET logged = logged - 1 WHERE id = " & .AccountID & ";"

        End If

        Call MakeQuery(q, True)

    End With

    Exit Sub

ErrorHandler:
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
        .Pos.x = QueryData!pos_x
        .Pos.Y = QueryData!pos_y
        .MENSAJEINFORMACION = QueryData!message_info
        .OrigChar.Body = QueryData!body_id
        .OrigChar.Head = QueryData!head_id
        .OrigChar.WeaponAnim = QueryData!weapon_id
        .OrigChar.CascoAnim = QueryData!helmet_id
        .OrigChar.ShieldAnim = QueryData!shield_id
        .OrigChar.heading = QueryData!heading
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
        'Call MakeQuery("SELECT * FROM pet WHERE user_id = " & .ID & ";")

        'If Not QueryData Is Nothing Then
        '    QueryData.MoveFirst

        '    While Not QueryData.EOF

        '        .MascotasType(QueryData!Number) = QueryData!pet_id

        '        QueryData.MoveNext
        '    Wend
        'End If

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
    GetCuentaValue = GetDBValue("account", Columna, "email", LCase$(CuentaEmail))

End Function

Public Function GetUserValue(CharName As String, Columna As String) As Variant
    ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para cuando hay que leer un unico valor del char
    GetUserValue = GetDBValue("user", Columna, "name", CharName)

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
    Call SetDBValue("account", Columna, Value, "email", LCase$(CuentaEmail))

End Sub

Private Sub SetUserValue(CharName As String, Columna As String, Value As Variant)
    ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para cuando hay que escribir un unico valor del char
    Call SetDBValue("user", Columna, Value, "name", CharName)

End Sub

Private Sub SetCuentaValueByID(AccountID As Long, Columna As String, Value As Variant)
    ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para cuando hay que escribir un unico valor de la cuenta
    ' Por ID
    Call SetDBValue("account", Columna, Value, "id", AccountID)

End Sub

Private Sub SetUserValueByID(Id As Long, Columna As String, Value As Variant)
    ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para cuando hay que escribir un unico valor del char
    ' Por ID
    Call SetDBValue("user", Columna, Value, "id", Id)

End Sub

Public Function CheckUserDonatorDatabase(CuentaEmail As String) As Boolean
    CheckUserDonatorDatabase = GetCuentaValue(CuentaEmail, "is_donor")

End Function

Public Function GetUserCreditosDatabase(CuentaEmail As String) As Long
    GetUserCreditosDatabase = GetCuentaValue(CuentaEmail, "credits")

End Function

Public Function GetUserCreditosCanjeadosDatabase(CuentaEmail As String) As Long
    GetUserCreditosCanjeadosDatabase = GetCuentaValue(CuentaEmail, "credits_used")

End Function

Public Function GetUserDiasDonadorDatabase(CuentaEmail As String) As Long

    Dim DonadorExpire As Variant

    DonadorExpire = SanitizeNullValue(GetCuentaValue(CuentaEmail, "donor_expire"), False)
    
    If Not DonadorExpire Then Exit Function
    GetUserDiasDonadorDatabase = DateDiff("d", Date, DonadorExpire)

End Function

Public Function GetUserComprasDonadorDatabase(CuentaEmail As String) As Long
    GetUserComprasDonadorDatabase = GetCuentaValue(CuentaEmail, "donor_purchases")

End Function

Public Function CheckUserExists(name As String) As Boolean
    CheckUserExists = GetUserValue(name, "COUNT(*)") > 0

End Function

Public Function CheckCuentaExiste(CuentaEmail As String) As Boolean
    CheckCuentaExiste = GetCuentaValue(CuentaEmail, "COUNT(*)") > 0

End Function

Public Function BANCheckDatabase(name As String) As Boolean
    BANCheckDatabase = CBool(GetUserValue(name, "is_banned"))

End Function

Public Function GetCodigoActivacionDatabase(name As String) As String
    GetCodigoActivacionDatabase = GetCuentaValue(name, "validate_code")

End Function

Public Function CheckCuentaActivadaDatabase(name As String) As Boolean
    CheckCuentaActivadaDatabase = GetCuentaValue(name, "validated")

End Function

Public Function GetEmailDatabase(name As String) As String
    GetEmailDatabase = GetCuentaValue(name, "email")

End Function

Public Function GetMacAddressDatabase(CuentaEmail As String) As String
    GetMacAddressDatabase = GetCuentaValue(CuentaEmail, "mac_address")

End Function

Public Function GetHDSerialDatabase(CuentaEmail As String) As Long
    GetHDSerialDatabase = GetCuentaValue(CuentaEmail, "hd_serial")

End Function

Public Function CheckBanCuentaDatabase(CuentaEmail As String) As Boolean
    CheckBanCuentaDatabase = CBool(GetCuentaValue(CuentaEmail, "is_banned"))

End Function

Public Function GetMotivoBanCuentaDatabase(CuentaEmail As String) As String
    GetMotivoBanCuentaDatabase = GetCuentaValue(CuentaEmail, "ban_reason")

End Function

Public Function GetQuienBanCuentaDatabase(CuentaEmail As String) As String
    GetQuienBanCuentaDatabase = GetCuentaValue(CuentaEmail, "banned_by")

End Function

Public Function GetCuentaLogeadaDatabase(CuentaEmail As String) As Boolean
    GetCuentaLogeadaDatabase = GetCuentaValue(CuentaEmail, "is_logged")

End Function

Public Function GetUserStatusDatabase(name As String) As Integer
    GetUserStatusDatabase = GetUserValue(name, "status")

End Function

Public Function GetAccountIDDatabase(name As String) As Long

    Dim temp As Variant

    temp = GetUserValue(name, "account_id")
    
    If VBA.IsEmpty(temp) Then
        GetAccountIDDatabase = -1
    Else
        GetAccountIDDatabase = val(temp)

    End If

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

    Call MakeQuery("SELECT name, head_id, class_id, body_id, pos_map, level, status, helmet_id, shield_id, weapon_id, guild_index, is_dead FROM user WHERE deleted = FALSE AND account_id = " & AccountID & ";")

    If QueryData Is Nothing Then Exit Function
    
    GetPersonajesCuentaDatabase = QueryData.RecordCount
        
    QueryData.MoveFirst
    
    Dim i As Integer

    For i = 1 To GetPersonajesCuentaDatabase
        Personaje(i).nombre = QueryData!name
        Personaje(i).Cabeza = QueryData!head_id
        Personaje(i).clase = QueryData!class_id
        Personaje(i).cuerpo = QueryData!body_id
        Personaje(i).Mapa = QueryData!pos_map
        Personaje(i).nivel = QueryData!level
        Personaje(i).Status = QueryData!Status
        Personaje(i).Casco = QueryData!helmet_id
        Personaje(i).Escudo = QueryData!shield_id
        Personaje(i).Arma = QueryData!weapon_id
        Personaje(i).ClanIndex = QueryData!Guild_Index
        
        If EsRolesMaster(Personaje(i).nombre) Then
            Personaje(i).Status = 3
        ElseIf EsConsejero(Personaje(i).nombre) Then
            Personaje(i).Status = 4
        ElseIf EsSemiDios(Personaje(i).nombre) Then
            Personaje(i).Status = 5
        ElseIf EsDios(Personaje(i).nombre) Then
            Personaje(i).Status = 6
        ElseIf EsAdmin(Personaje(i).nombre) Then
            Personaje(i).Status = 7

        End If

        If val(QueryData!is_dead) = 1 Then
            Personaje(i).Cabeza = iCabezaMuerto

        End If
        
        QueryData.MoveNext
    Next

End Function

Public Sub SetUserLoggedDatabase(ByVal Id As Long, ByVal AccountID As Long)
    Call SetDBValue("user", "is_logged", 1, "id", Id)
    Call MakeQuery("UPDATE account SET logged = logged + 1 WHERE id = " & AccountID & ";", True)

End Sub

Public Sub ResetLoggedDatabase(ByVal AccountID As Long)
    Call MakeQuery("UPDATE account SET logged = 0 WHERE id = " & AccountID & ";", True)

End Sub

Public Sub SetUsersLoggedDatabase(ByVal NumUsers As Long)
    Call MakeQuery("UPDATE statistics SET value = '" & NumUsers & "' WHERE name = 'online';", True)

End Sub

Public Sub LogoutAllUsersAndAccounts()

    Dim Query As String

    Query = "UPDATE user SET is_logged = FALSE; "
    Query = Query & "UPDATE account SET logged = 0;"

    Call MakeQuery(Query, True)

End Sub

Public Sub SaveBattlePointsDatabase(ByVal Id As Long, ByVal BattlePoints As Long)
    Call SetDBValue("user", "battle_points", BattlePoints, "id", Id)

End Sub

Public Sub SaveVotoDatabase(ByVal Id As Long, ByVal Encuestas As Integer)
    Call SetUserValueByID(Id, "votes_amount", Encuestas)

End Sub

Public Sub SaveUserBodyDatabase(UserName As String, ByVal Body As Integer)
    Call SetUserValue(UserName, "body_id", Body)

End Sub

Public Sub SaveUserHeadDatabase(UserName As String, ByVal Head As Integer)
    Call SetUserValue(UserName, "head_id", Head)

End Sub

Public Sub SaveUserSkillDatabase(UserName As String, ByVal Skill As Integer, ByVal Value As Integer)
    Call MakeQuery("UPDATE skillpoints SET value = " & Value & " WHERE number = " & Skill & " AND user_id = (SELECT id FROM user WHERE name = '" & UserName & "');", True)

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
    Call SetCuentaValue(UserCuenta, "validated", 1)

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
    Call MakeQuery("UPDATE user SET pertenece_consejo_real = FALSE, pertenece_consejo_caos = FALSE WHERE name = '" & UserName & "';", True)

End Sub

Public Sub EcharLegionDatabase(UserName As String)
    Call MakeQuery("UPDATE user SET pertenece_caos = FALSE, reenlistadas = 200 WHERE name = '" & UserName & "';", True)

End Sub

Public Sub EcharArmadaDatabase(UserName As String)
    Call MakeQuery("UPDATE user SET pertenece_real = FALSE, reenlistadas = 200 WHERE name = '" & UserName & "';", True)

End Sub

Public Sub CambiarPenaDatabase(UserName As String, ByVal Numero As Integer, Pena As String)
    Call MakeQuery("UPDATE punishment SET reason = '" & Pena & "' WHERE number = " & Numero & " AND user_id = (SELECT id from user WHERE name = '" & UserName & "');", True)

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

Public Function SanitizeNullValue(ByVal Value As Variant, ByVal defaultValue As Variant) As Variant
    SanitizeNullValue = IIf(IsNull(Value), defaultValue, Value)

End Function

