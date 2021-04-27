Attribute VB_Name = "Database"
'Argentum Online Libre
'Database connection module
'Obtained from GS-Zone
'Adapted and modified by Juan Andres Dalmasso (CHOTS)
'September 2018
'Rewrited for Argentum20 by Alexis Caraballo (WyroX)
'October 2020

Option Explicit

Public Database_Enabled     As Boolean
Public Database_DataSource  As String
Public Database_Host        As String
Public Database_Name        As String
Public Database_Username    As String
Public Database_Password    As String

Private Database_Connection As ADODB.Connection
Private Command             As ADODB.Command
Private QueryData           As ADODB.Recordset
Private RecordsAffected     As Long

Private QueryBuilder        As cStringBuilder
Private ConnectedOnce       As Boolean

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
    Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description)
    
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
    
    Set Command = Nothing
     
    Call Database_Connection.Close
    
    Set Database_Connection = Nothing
     
    Exit Sub
     
ErrorHandler:
    Call LogDatabaseError("Unable to close Mysql Database: " & Err.Number & " - " & Err.Description)

End Sub


Public Sub SaveNewUserDatabase(ByVal UserIndex As Integer)

    On Error GoTo ErrorHandler
    
    Dim LoopC As Long
    Dim ParamC As Long
    Dim Params() As Variant
    
    'Constructor de queries.
    'Me permite concatenar strings MUCHO MAS rapido
    Set QueryBuilder = New cStringBuilder
    
    With UserList(UserIndex)
        
        ReDim Params(45)
        
        '  ************ Basic user data *******************
        Params(0) = .name
        Params(1) = .AccountId
        Params(2) = .Stats.ELV
        Params(3) = .Stats.Exp
        Params(4) = .genero
        Params(5) = .raza
        Params(6) = .clase
        Params(7) = .Hogar
        Params(8) = .Desc
        Params(9) = .Stats.GLD
        Params(10) = .Stats.SkillPts
        Params(11) = .Pos.Map
        Params(12) = .Pos.X
        Params(13) = .Pos.Y
        Params(14) = .Char.Body
        Params(15) = .Char.Head
        Params(16) = .Char.WeaponAnim
        Params(17) = .Char.CascoAnim
        Params(18) = .Char.ShieldAnim
        Params(19) = .Invent.NroItems
        Params(20) = .Invent.ArmourEqpSlot
        Params(21) = .Invent.WeaponEqpSlot
        Params(22) = .Invent.EscudoEqpSlot
        Params(23) = .Invent.CascoEqpSlot
        Params(24) = .Invent.MunicionEqpSlot
        Params(25) = .Invent.DañoMagicoEqpSlot
        Params(26) = .Invent.ResistenciaEqpSlot
        Params(27) = .Invent.HerramientaEqpSlot
        Params(28) = .Invent.MagicoSlot
        Params(29) = .Invent.NudilloSlot
        Params(30) = .Invent.BarcoSlot
        Params(31) = .Invent.MonturaSlot
        Params(32) = .Stats.MinHp
        Params(33) = .Stats.MaxHp
        Params(34) = .Stats.MinMAN
        Params(35) = .Stats.MaxMAN
        Params(36) = .Stats.MinSta
        Params(37) = .Stats.MaxSta
        Params(38) = .Stats.MinHam
        Params(39) = .Stats.MaxHam
        Params(40) = .Stats.MinAGU
        Params(41) = .Stats.MaxAGU
        Params(42) = .Stats.MinHIT
        Params(43) = .Stats.MaxHit
        Params(44) = .flags.Desnudo
        Params(45) = .Faccion.Status
        
        Call MakeQuery(QUERY_SAVE_MAINPJ, True, Params)

        ' Para recibir el ID del user
        Call MakeQuery("SELECT LAST_INSERT_ID();", False)

        If QueryData Is Nothing Then
            .Id = 1
        Else
            .Id = val(QueryData.Fields(0).Value)
        End If
        
        ' ******************* ATRIBUTOS *******************
        ReDim Params(NUMATRIBUTOS * 3 - 1)
        ParamC = 0
        
        For LoopC = 1 To NUMATRIBUTOS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = .Stats.UserAtributos(LoopC)
            
            ParamC = ParamC + 3
        Next LoopC
        
        Call MakeQuery(QUERY_SAVE_ATTRIBUTES, True, Params)
        
        ' ******************* SPELLS **********************
        ReDim Params(MAXUSERHECHIZOS * 3 - 1)
        ParamC = 0
        
        For LoopC = 1 To MAXUSERHECHIZOS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = .Stats.UserHechizos(LoopC)
            
            ParamC = ParamC + 3
        Next LoopC

        Call MakeQuery(QUERY_SAVE_SPELLS, True, Params)
        
        ' ******************* INVENTORY *******************
        ReDim Params(MAX_INVENTORY_SLOTS * 5 - 1)
        ParamC = 0
        
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = .Invent.Object(LoopC).ObjIndex
            Params(ParamC + 3) = .Invent.Object(LoopC).amount
            Params(ParamC + 4) = .Invent.Object(LoopC).Equipped
            
            ParamC = ParamC + 5
        Next LoopC
        
        Call MakeQuery(QUERY_SAVE_INVENTORY, True, Params)
        
        ' ******************* SKILLS *******************
        ReDim Params(NUMSKILLS * 3 - 1)
        ParamC = 0
        
        For LoopC = 1 To NUMSKILLS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = .Stats.UserSkills(LoopC)
            
            ParamC = ParamC + 3
        Next LoopC
        
        Call MakeQuery(QUERY_SAVE_SKILLS, True, Params)
        
        ' ******************* QUESTS *******************
        ReDim Params(MAXUSERQUESTS * 2 - 1)
        ParamC = 0
        
        For LoopC = 1 To MAXUSERQUESTS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            
            ParamC = ParamC + 2
        Next LoopC
        
        Call MakeQuery(QUERY_SAVE_QUESTS, True, Params)
        
        ' ******************* PETS ********************
        ReDim Params(MAXMASCOTAS * 3 - 1)
        ParamC = 0
        
        For LoopC = 1 To MAXMASCOTAS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = 0
            
            ParamC = ParamC + 3
        Next LoopC
    
        Call MakeQuery(QUERY_SAVE_PETS, True, Params)
        
        Set QueryBuilder = Nothing
    
    End With

    Exit Sub

ErrorHandler:
    
    Set QueryBuilder = Nothing
    
    Call LogDatabaseError("Error en SaveNewUserDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserDatabase(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean = False)

    On Error GoTo ErrorHandler
    
    Dim Params() As Variant
    Dim LoopC As Long
    Dim ParamC As Long
    
    'Constructor de queries.
    'Me permite concatenar strings MUCHO MAS rapido
    Set QueryBuilder = New cStringBuilder

    'Basic user data
    With UserList(UserIndex)
        
        ReDim Params(89)
        
        Params(0) = .name
        Params(1) = .Stats.ELV
        Params(2) = CLng(.Stats.Exp)
        Params(3) = .genero
        Params(4) = .raza
        Params(5) = .clase
        Params(6) = .Hogar
        Params(7) = .Desc
        Params(8) = .Stats.GLD
        Params(9) = .Stats.Banco
        Params(10) = .Stats.SkillPts
        Params(11) = .flags.MascotasGuardadas
        Params(12) = .Pos.Map
        Params(13) = .Pos.X
        Params(14) = .Pos.Y
        Params(15) = .flags.lastMap
        Params(16) = .MENSAJEINFORMACION
        Params(17) = .Char.Body
        Params(18) = .OrigChar.Head
        Params(19) = .Char.WeaponAnim
        Params(20) = .Char.CascoAnim
        Params(21) = .Char.ShieldAnim
        Params(22) = .Char.Heading
        Params(23) = .Invent.NroItems
        Params(24) = .Invent.ArmourEqpSlot
        Params(25) = .Invent.WeaponEqpSlot
        Params(26) = .Invent.EscudoEqpSlot
        Params(27) = .Invent.CascoEqpSlot
        Params(28) = .Invent.MunicionEqpSlot
        Params(29) = .Invent.DañoMagicoEqpSlot
        Params(30) = .Invent.ResistenciaEqpSlot
        Params(31) = .Invent.HerramientaEqpSlot
        Params(32) = .Invent.MagicoSlot
        Params(33) = .Invent.NudilloSlot
        Params(34) = .Invent.BarcoSlot
        Params(35) = .Invent.MonturaSlot
        Params(36) = .Stats.MinHp
        Params(37) = .Stats.MaxHp
        Params(38) = .Stats.MinMAN
        Params(39) = .Stats.MaxMAN
        Params(40) = .Stats.MinSta
        Params(41) = .Stats.MaxSta
        Params(42) = .Stats.MinHam
        Params(43) = .Stats.MaxHam
        Params(44) = .Stats.MinAGU
        Params(45) = .Stats.MaxAGU
        Params(46) = .Stats.MinHIT
        Params(47) = .Stats.MaxHit
        Params(48) = .Stats.NPCsMuertos
        Params(49) = .Stats.UsuariosMatados
        Params(50) = .Stats.InventLevel
        Params(51) = .flags.Desnudo
        Params(52) = .flags.Envenenado
        Params(53) = .flags.Escondido
        Params(54) = .flags.Hambre
        Params(55) = .flags.Sed
        Params(56) = .flags.Muerto
        Params(57) = .flags.Navegando
        Params(58) = .flags.Paralizado
        Params(59) = .flags.Montado
        Params(60) = .flags.Silenciado
        Params(61) = .flags.MinutosRestantes
        Params(62) = .flags.SegundosPasados
        Params(63) = .flags.Pareja
        Params(64) = .Counters.Pena
        Params(65) = .flags.VecesQueMoriste
        Params(66) = (.flags.Privilegios And PlayerType.RoyalCouncil)
        Params(67) = (.flags.Privilegios And PlayerType.ChaosCouncil)
        Params(68) = .Faccion.ArmadaReal
        Params(69) = .Faccion.FuerzasCaos
        Params(70) = .Faccion.ciudadanosMatados
        Params(71) = .Faccion.CriminalesMatados
        Params(72) = .Faccion.RecibioArmaduraReal
        Params(73) = .Faccion.RecibioArmaduraCaos
        Params(74) = .Faccion.RecibioExpInicialReal
        Params(75) = .Faccion.RecibioExpInicialCaos
        Params(76) = .Faccion.RecompensasReal
        Params(77) = .Faccion.RecompensasCaos
        Params(78) = .Faccion.Reenlistadas
        Params(79) = IIf(Len(.Faccion.FechaIngreso) <> 0, .Faccion.FechaIngreso, "NULL")
        Params(80) = .Faccion.NivelIngreso
        Params(81) = .Faccion.MatadosIngreso
        Params(82) = .Faccion.NextRecompensa
        Params(83) = .Faccion.Status
        Params(84) = .GuildIndex
        Params(85) = .ChatCombate
        Params(86) = .ChatGlobal
        Params(87) = IIf(Logout, "FALSE", "TRUE")
        Params(88) = .Stats.Advertencias
        
        ' WHERE block
        Params(89) = .Id
        
        Call MakeQuery(QUERY_UPDATE_MAINPJ, True, Params)
        
        Call QueryBuilder.Clear

        ' ************************** User attributes ****************************
        ReDim Params(NUMATRIBUTOS * 3 - 1)
        ParamC = 0
        
        For LoopC = 1 To NUMATRIBUTOS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = .Stats.UserAtributosBackUP(LoopC)
            
            ParamC = ParamC + 3
        Next LoopC
        
        Call MakeQuery(QUERY_UPSERT_ATTRIBUTES, True, Params)

        ' ************************** User spells *********************************
        ReDim Params(MAXUSERHECHIZOS * 3 - 1)
        ParamC = 0
        
        For LoopC = 1 To MAXUSERHECHIZOS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = .Stats.UserHechizos(LoopC)
            
            ParamC = ParamC + 3
        Next LoopC
        
        Call MakeQuery(QUERY_UPSERT_SPELLS, True, Params)

        ' ************************** User inventory *********************************
        ReDim Params(MAX_INVENTORY_SLOTS * 5 - 1)
        ParamC = 0
        
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = .Invent.Object(LoopC).ObjIndex
            Params(ParamC + 3) = .Invent.Object(LoopC).amount
            Params(ParamC + 4) = .Invent.Object(LoopC).Equipped
            
            ParamC = ParamC + 5
        Next LoopC
        
        Call MakeQuery(QUERY_UPSERT_INVENTORY, True, Params)

        ' ************************** User bank inventory *********************************
        ReDim Params(MAX_BANCOINVENTORY_SLOTS * 4 - 1)
        ParamC = 0
        
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = .BancoInvent.Object(LoopC).ObjIndex
            Params(ParamC + 3) = .BancoInvent.Object(LoopC).amount
            
            ParamC = ParamC + 4
        Next LoopC
        
        Call MakeQuery(QUERY_SAVE_BANCOINV, True, Params)

        ' ************************** User skills *********************************
        ReDim Params(NUMSKILLS * 3 - 1)
        ParamC = 0
        
        For LoopC = 1 To NUMSKILLS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC
            Params(ParamC + 2) = .Stats.UserSkills(LoopC)
            
            ParamC = ParamC + 3
        Next LoopC
        
        Call MakeQuery(QUERY_UPSERT_SKILLS, True, Params)

        ' ************************** User pets *********************************
        ReDim Params(MAXMASCOTAS * 3 - 1)
        ParamC = 0
        Dim petType As Integer

        For LoopC = 1 To MAXMASCOTAS
            Params(ParamC) = .Id
            Params(ParamC + 1) = LoopC

            'CHOTS | I got this logic from SaveUserToCharfile
            If .MascotasIndex(LoopC) > 0 Then
            
                If NpcList(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
                    petType = .MascotasType(LoopC)
                Else
                    petType = 0
                End If

            Else
                petType = .MascotasType(LoopC)

            End If

            Params(ParamC + 2) = petType
            
            ParamC = ParamC + 3
        Next LoopC
        
        Call MakeQuery(QUERY_UPSERT_PETS, True, Params)
        
        ' ************************** User connection logs *********************************
        
        'Agrego ip del user
        Call MakeQuery(QUERY_SAVE_CONNECTION, True, .Id, .ip)
        
        'Borro la mas vieja si hay mas de 5 (WyroX: si alguien sabe una forma mejor de hacerlo me avisa)
        Call MakeQuery(QUERY_DELETE_LAST_CONNECTIONS, True, .Id, .Id)
        
        ' ************************** User quests *********************************
        QueryBuilder.Append "INSERT INTO quest (user_id, number, quest_id, npcs, npcstarget) VALUES "
        
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
            
            QueryBuilder.Append "', '"
            
            If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
                
                Tmp = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredTargetNPCs
                    
                For LoopK = 1 To Tmp

                    QueryBuilder.Append CStr(.QuestStats.Quests(LoopC).NPCsTarget(LoopK))
                    
                    If LoopK < Tmp Then
                        QueryBuilder.Append "-"
                    End If
                
                Next LoopK
            
            End If
            
            QueryBuilder.Append "')"

            If LoopC < MAXUSERQUESTS Then
                QueryBuilder.Append ", "
            End If

        Next LoopC
        
        QueryBuilder.Append " ON DUPLICATE KEY UPDATE quest_id=VALUES(quest_id), npcs=VALUES(npcs); "
        
        Call MakeQuery(QueryBuilder.ToString, True)
        
        Call QueryBuilder.Clear
        
        ' ************************** User completed quests *********************************
        If .QuestStats.NumQuestsDone > 0 Then
            
            ' Armamos la query con los placeholders
            QueryBuilder.Append "INSERT INTO quest_done (user_id, quest_id) VALUES "
            
            For LoopC = 1 To .QuestStats.NumQuestsDone
                QueryBuilder.Append "(?, ?)"
            
                If LoopC < .QuestStats.NumQuestsDone Then
                    QueryBuilder.Append ", "
                End If
            
            Next LoopC
                    
            QueryBuilder.Append " ON DUPLICATE KEY UPDATE quest_id=VALUES(quest_id); "
            
            ' Metemos los parametros
            ReDim Params(.QuestStats.NumQuestsDone * 2 - 1)
            ParamC = 0
            
            For LoopC = 1 To .QuestStats.NumQuestsDone
                Params(ParamC) = .Id
                Params(ParamC + 1) = .QuestStats.QuestsDone(LoopC)
                
                ParamC = ParamC + 2
            Next LoopC
            
            ' Mandamos la query
            Call MakeQuery(QueryBuilder.ToString, True, Params)
        
        End If

        ' ************************** User logout *********************************
        ' Si deslogueó, actualizo la cuenta
        If Logout Then
            Call MakeQuery("UPDATE account SET logged = logged - 1 WHERE id = ?", True, .AccountId)
        End If

    End With
    
    Set QueryBuilder = Nothing
    
    Exit Sub

ErrorHandler:

    Set QueryBuilder = Nothing
    
    Call LogDatabaseError("Error en SaveUserDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description)

End Sub

Sub LoadUserDatabase(ByVal UserIndex As Integer)

    On Error GoTo ErrorHandler

    'Basic user data
    With UserList(UserIndex)

        Call MakeQuery("SELECT *, DATE_FORMAT(fecha_ingreso, '%Y-%m-%d') as 'fecha_ingreso_format' FROM user WHERE name = ?;", False, .name)

        If QueryData Is Nothing Then Exit Sub

        'Start setting data
        .Id = QueryData!Id
        .name = QueryData!name
        .Stats.ELV = QueryData!level
        .Stats.Exp = QueryData!Exp
        .genero = QueryData!genre_id
        .raza = QueryData!race_id
        .clase = QueryData!class_id
        .Hogar = QueryData!home_id
        .Desc = QueryData!Description
        .Stats.GLD = QueryData!gold
        .Stats.Banco = QueryData!bank_gold
        .Stats.SkillPts = QueryData!free_skillpoints
        '.Counters.AsignedSkills = QueryData!assigned_skillpoints
        .Pos.Map = QueryData!pos_map
        .Pos.X = QueryData!pos_x
        .Pos.Y = QueryData!pos_y
        .flags.lastMap = QueryData!last_map
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
        .Invent.DañoMagicoEqpSlot = SanitizeNullValue(QueryData!slot_dm, 0)
        .Invent.ResistenciaEqpSlot = SanitizeNullValue(QueryData!slot_rm, 0)
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
        .flags.Montado = QueryData!is_mounted
        .flags.Pareja = QueryData!spouse
        .flags.Casado = IIf(Len(.flags.Pareja) > 0, 1, 0)
        .flags.Silenciado = QueryData!is_silenced
        .flags.MinutosRestantes = QueryData!silence_minutes_left
        .flags.SegundosPasados = QueryData!silence_elapsed_seconds
        .flags.MascotasGuardadas = QueryData!pets_saved
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
        .Faccion.ciudadanosMatados = QueryData!ciudadanos_matados
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
        
        .Stats.Advertencias = QueryData!warnings
        
        'User attributes
        Call MakeQuery("SELECT * FROM attribute WHERE user_id = ?;", False, .Id)
    
        If Not QueryData Is Nothing Then
            QueryData.MoveFirst

            While Not QueryData.EOF

                .Stats.UserAtributos(QueryData!Number) = QueryData!Value
                .Stats.UserAtributosBackUP(QueryData!Number) = .Stats.UserAtributos(QueryData!Number)

                QueryData.MoveNext
            Wend

        End If

        'User spells
        Call MakeQuery("SELECT * FROM spell WHERE user_id = ?;", False, .Id)

        If Not QueryData Is Nothing Then
            QueryData.MoveFirst

            While Not QueryData.EOF

                .Stats.UserHechizos(QueryData!Number) = QueryData!spell_id

                QueryData.MoveNext
            Wend

        End If

        'User pets
        Call MakeQuery("SELECT * FROM pet WHERE user_id = ?;", False, .Id)

        If Not QueryData Is Nothing Then
            QueryData.MoveFirst

            While Not QueryData.EOF

                .MascotasType(QueryData!Number) = QueryData!pet_id
                
                If val(QueryData!pet_id) <> 0 Then
                    .NroMascotas = .NroMascotas + 1

                End If

                QueryData.MoveNext
            Wend

        End If

        'User inventory
        Call MakeQuery("SELECT * FROM inventory_item WHERE user_id = ?;", False, .Id)

        If Not QueryData Is Nothing Then
            QueryData.MoveFirst

            While Not QueryData.EOF

                With .Invent.Object(QueryData!Number)
                    .ObjIndex = QueryData!item_id
                
                    If .ObjIndex <> 0 Then
                        If LenB(ObjData(.ObjIndex).name) Then
                            .amount = QueryData!amount
                            .Equipped = QueryData!is_equipped
                        Else
                            .ObjIndex = 0

                        End If

                    End If

                End With

                QueryData.MoveNext
            Wend

        End If

        'User bank inventory
        Call MakeQuery("SELECT * FROM bank_item WHERE user_id = ?;", False, .Id)

        If Not QueryData Is Nothing Then
            QueryData.MoveFirst

            While Not QueryData.EOF

                With .BancoInvent.Object(QueryData!Number)
                    .ObjIndex = QueryData!item_id
                
                    If .ObjIndex <> 0 Then
                        If LenB(ObjData(.ObjIndex).name) Then
                            .amount = QueryData!amount
                        Else
                            .ObjIndex = 0

                        End If

                    End If

                End With

                QueryData.MoveNext
            Wend

        End If

        'User skills
        Call MakeQuery("SELECT * FROM skillpoint WHERE user_id = ?;", False, .Id)

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
        'Call MakeQuery("SELECT * FROM friend WHERE user_id = ?;", False, .Id)

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
        Call MakeQuery("SELECT * FROM quest WHERE user_id = ?;", False, .Id)

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
                    
                    If QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredTargetNPCs Then

                        Dim NPCsTarget() As String

                        NPCsTarget = Split(QueryData!NPCsTarget, "-")
                        ReDim .QuestStats.Quests(QueryData!Number).NPCsTarget(1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredTargetNPCs)

                        For LoopC = 1 To QuestList(.QuestStats.Quests(QueryData!Number).QuestIndex).RequiredTargetNPCs
                            .QuestStats.Quests(QueryData!Number).NPCsTarget(LoopC) = val(NPCsTarget(LoopC - 1))
                        Next LoopC

                    End If

                End If

                QueryData.MoveNext
            Wend

        End If
        
        'User quests done
        Call MakeQuery("SELECT * FROM quest_done WHERE user_id = ?;", False, .Id)

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
        Call MakeQuery("SELECT key_obj FROM house_key WHERE account_id = ?", False, .AccountId)

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
    Call LogDatabaseError("Error en LoadUserDatabase: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description & ". Línea: " & Erl)

    Resume Next

End Sub

Public Function MakeQuery(query As String, ByVal NoResult As Boolean, ParamArray Query_Parameters() As Variant) As Boolean
    ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Hace una unica query a la db. Asume una conexion.
    ' Si NoResult = False, el metodo lee el resultado de la query
    ' Guarda el resultado en QueryData
    
    On Error GoTo ErrorHandler
    
    Dim Params As Variant

    Set Command = New ADODB.Command
    
    With Command

        .ActiveConnection = Database_Connection
        .CommandType = adCmdText
        .NamedParameters = False
        .CommandText = query
        
        If UBound(Query_Parameters) < 0 Then
            Params = Null
            
        Else
            Params = Query_Parameters
            
            If IsArray(Query_Parameters(0)) Then
                .Prepared = True
                Params = Query_Parameters(0)
            End If

        End If

        If NoResult Then
            Call .Execute(RecordsAffected, Params, adExecuteNoRecords)
    
        Else
            Set QueryData = .Execute(RecordsAffected, Params)
    
            If QueryData.BOF Or QueryData.EOF Then
                Set QueryData = Nothing
            End If
    
        End If
        
    End With
    
    Exit Function
    
ErrorHandler:

    If Not adoIsConnected(Database_Connection) Then
        Call LogDatabaseError("Alerta en MakeQuery: Se perdió la conexión con la DB. Reconectando.")
        Call Database_Connect
        Resume
        
    Else
        Call LogDatabaseError("Error en MakeQuery: query = '" & query & "'. " & Err.Number & " - " & Err.Description)
        
        On Error GoTo 0

        Err.raise Err.Number

    End If

End Function

Private Function GetDBValue(Tabla As String, ColumnaGet As String, ColumnaTest As String, ValueTest As Variant) As Variant
    ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para leer un unico valor de una unica fila

    On Error GoTo ErrorHandler
    
    Call MakeQuery("SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = ?;", False, ValueTest)
    
    'Revisamos si recibio un resultado
    If QueryData Is Nothing Then Exit Function

    'Obtenemos la variable
    GetDBValue = QueryData.Fields(0).Value

    Exit Function
    
ErrorHandler:
    Call LogDatabaseError("Error en GetDBValue: SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = '" & ValueTest & "';" & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetCuentaValue(CuentaEmail As String, Columna As String) As Variant
    ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para cuando hay que leer un unico valor de la cuenta
        
    On Error GoTo GetCuentaValue_Err
        
    GetCuentaValue = GetDBValue("account", Columna, "email", LCase$(CuentaEmail))

        
    Exit Function

GetCuentaValue_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetCuentaValue", Erl)
    Resume Next
        
End Function

Public Function GetUserValue(CharName As String, Columna As String) As Variant
    ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para cuando hay que leer un unico valor del char
        
    On Error GoTo GetUserValue_Err
        
    GetUserValue = GetDBValue("user", Columna, "name", CharName)

        
    Exit Function

GetUserValue_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetUserValue", Erl)
    Resume Next
        
End Function

Private Sub SetDBValue(Tabla As String, ColumnaSet As String, ByVal ValueSet As Variant, ColumnaTest As String, ByVal ValueTest As Variant)
    ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para escribir un unico valor de una unica fila

    On Error GoTo ErrorHandler
    
    'Hacemos la query
    Call MakeQuery("UPDATE " & Tabla & " SET " & ColumnaSet & " = ? WHERE " & ColumnaTest & " = ?;", True, ValueSet, ValueTest)

    Exit Sub
    
ErrorHandler:
    Call LogDatabaseError("Error en SetDBValue: UPDATE " & Tabla & " SET " & ColumnaSet & " = " & ValueSet & " WHERE " & ColumnaTest & " = " & ValueTest & ";" & ". " & Err.Number & " - " & Err.Description)

End Sub

Private Sub SetCuentaValue(CuentaEmail As String, Columna As String, Value As Variant)
    ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para cuando hay que escribir un unico valor de la cuenta
        
    On Error GoTo SetCuentaValue_Err
        
    Call SetDBValue("account", Columna, Value, "email", LCase$(CuentaEmail))

        
    Exit Sub

SetCuentaValue_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetCuentaValue", Erl)
    Resume Next
        
End Sub

Private Sub SetUserValue(CharName As String, Columna As String, Value As Variant)
    ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para cuando hay que escribir un unico valor del char
        
    On Error GoTo SetUserValue_Err
        
    Call SetDBValue("user", Columna, Value, "name", CharName)

        
    Exit Sub

SetUserValue_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetUserValue", Erl)
    Resume Next
        
End Sub

Private Sub SetCuentaValueByID(AccountId As Long, Columna As String, Value As Variant)
    ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para cuando hay que escribir un unico valor de la cuenta
    ' Por ID
        
    On Error GoTo SetCuentaValueByID_Err
        
    Call SetDBValue("account", Columna, Value, "id", AccountId)

        
    Exit Sub

SetCuentaValueByID_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetCuentaValueByID", Erl)
    Resume Next
        
End Sub

Private Sub SetUserValueByID(Id As Long, Columna As String, Value As Variant)
    ' 18/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Para cuando hay que escribir un unico valor del char
    ' Por ID
        
    On Error GoTo SetUserValueByID_Err
        
    Call SetDBValue("user", Columna, Value, "id", Id)

        
    Exit Sub

SetUserValueByID_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetUserValueByID", Erl)
    Resume Next
        
End Sub

Public Function CheckUserDonatorDatabase(CuentaEmail As String) As Boolean
        
    On Error GoTo CheckUserDonatorDatabase_Err
        
    CheckUserDonatorDatabase = GetCuentaValue(CuentaEmail, "is_donor")

        
    Exit Function

CheckUserDonatorDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.CheckUserDonatorDatabase", Erl)
    Resume Next
        
End Function

Public Function GetUserCreditosDatabase(CuentaEmail As String) As Long
        
    On Error GoTo GetUserCreditosDatabase_Err
        
    GetUserCreditosDatabase = GetCuentaValue(CuentaEmail, "credits")

        
    Exit Function

GetUserCreditosDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetUserCreditosDatabase", Erl)
    Resume Next
        
End Function

Public Function GetUserCreditosCanjeadosDatabase(CuentaEmail As String) As Long
        
    On Error GoTo GetUserCreditosCanjeadosDatabase_Err
        
    GetUserCreditosCanjeadosDatabase = GetCuentaValue(CuentaEmail, "credits_used")

        
    Exit Function

GetUserCreditosCanjeadosDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetUserCreditosCanjeadosDatabase", Erl)
    Resume Next
        
End Function

Public Function GetUserDiasDonadorDatabase(CuentaEmail As String) As Long
        
    On Error GoTo GetUserDiasDonadorDatabase_Err
        

    Dim DonadorExpire As Variant

    DonadorExpire = SanitizeNullValue(GetCuentaValue(CuentaEmail, "donor_expire"), False)
    
    If Not DonadorExpire Then Exit Function
    GetUserDiasDonadorDatabase = DateDiff("d", Date, DonadorExpire)

        
    Exit Function

GetUserDiasDonadorDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetUserDiasDonadorDatabase", Erl)
    Resume Next
        
End Function

Public Function GetUserComprasDonadorDatabase(CuentaEmail As String) As Long
        
    On Error GoTo GetUserComprasDonadorDatabase_Err
        
    GetUserComprasDonadorDatabase = GetCuentaValue(CuentaEmail, "donor_purchases")

        
    Exit Function

GetUserComprasDonadorDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetUserComprasDonadorDatabase", Erl)
    Resume Next
        
End Function

Public Function CheckUserExists(name As String) As Boolean
        
    On Error GoTo CheckUserExists_Err
        
    CheckUserExists = GetUserValue(name, "COUNT(*)") > 0

        
    Exit Function

CheckUserExists_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.CheckUserExists", Erl)
    Resume Next
        
End Function

Public Function CheckCuentaExiste(CuentaEmail As String) As Boolean
        
    On Error GoTo CheckCuentaExiste_Err
        
    CheckCuentaExiste = GetCuentaValue(CuentaEmail, "COUNT(*)") > 0

        
    Exit Function

CheckCuentaExiste_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.CheckCuentaExiste", Erl)
    Resume Next
        
End Function

Public Function BANCheckDatabase(name As String) As Boolean
        
    On Error GoTo BANCheckDatabase_Err
        
    BANCheckDatabase = CBool(GetUserValue(name, "is_banned"))

        
    Exit Function

BANCheckDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.BANCheckDatabase", Erl)
    Resume Next
        
End Function

Public Function GetCodigoActivacionDatabase(name As String) As String
        
    On Error GoTo GetCodigoActivacionDatabase_Err
        
    GetCodigoActivacionDatabase = GetCuentaValue(name, "validate_code")

        
    Exit Function

GetCodigoActivacionDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetCodigoActivacionDatabase", Erl)
    Resume Next
        
End Function

Public Function CheckCuentaActivadaDatabase(name As String) As Boolean
        
    On Error GoTo CheckCuentaActivadaDatabase_Err
        
    CheckCuentaActivadaDatabase = GetCuentaValue(name, "validated")

        
    Exit Function

CheckCuentaActivadaDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.CheckCuentaActivadaDatabase", Erl)
    Resume Next
        
End Function

Public Function GetEmailDatabase(name As String) As String
        
    On Error GoTo GetEmailDatabase_Err
        
    GetEmailDatabase = GetCuentaValue(name, "email")

        
    Exit Function

GetEmailDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetEmailDatabase", Erl)
    Resume Next
        
End Function

Public Function GetMacAddressDatabase(CuentaEmail As String) As String
        
    On Error GoTo GetMacAddressDatabase_Err
        
    GetMacAddressDatabase = GetCuentaValue(CuentaEmail, "mac_address")

        
    Exit Function

GetMacAddressDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetMacAddressDatabase", Erl)
    Resume Next
        
End Function

Public Function GetHDSerialDatabase(CuentaEmail As String) As Long
        
    On Error GoTo GetHDSerialDatabase_Err
        
    GetHDSerialDatabase = GetCuentaValue(CuentaEmail, "hd_serial")

        
    Exit Function

GetHDSerialDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetHDSerialDatabase", Erl)
    Resume Next
        
End Function

Public Function CheckBanCuentaDatabase(CuentaEmail As String) As Boolean
        
    On Error GoTo CheckBanCuentaDatabase_Err
        
    CheckBanCuentaDatabase = CBool(GetCuentaValue(CuentaEmail, "is_banned"))

        
    Exit Function

CheckBanCuentaDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.CheckBanCuentaDatabase", Erl)
    Resume Next
        
End Function

Public Function GetMotivoBanCuentaDatabase(CuentaEmail As String) As String
        
    On Error GoTo GetMotivoBanCuentaDatabase_Err
        
    GetMotivoBanCuentaDatabase = GetCuentaValue(CuentaEmail, "ban_reason")

        
    Exit Function

GetMotivoBanCuentaDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetMotivoBanCuentaDatabase", Erl)
    Resume Next
        
End Function

Public Function GetQuienBanCuentaDatabase(CuentaEmail As String) As String
        
    On Error GoTo GetQuienBanCuentaDatabase_Err
        
    GetQuienBanCuentaDatabase = GetCuentaValue(CuentaEmail, "banned_by")

        
    Exit Function

GetQuienBanCuentaDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetQuienBanCuentaDatabase", Erl)
    Resume Next
        
End Function

Public Function GetCuentaLogeadaDatabase(CuentaEmail As String) As Boolean
        
    On Error GoTo GetCuentaLogeadaDatabase_Err
        
    GetCuentaLogeadaDatabase = GetCuentaValue(CuentaEmail, "is_logged")

        
    Exit Function

GetCuentaLogeadaDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetCuentaLogeadaDatabase", Erl)
    Resume Next
        
End Function

Public Function GetUserStatusDatabase(name As String) As Integer
        
    On Error GoTo GetUserStatusDatabase_Err
        
    GetUserStatusDatabase = GetUserValue(name, "status")

        
    Exit Function

GetUserStatusDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetUserStatusDatabase", Erl)
    Resume Next
        
End Function

Public Function GetAccountIDDatabase(name As String) As Long
        
    On Error GoTo GetAccountIDDatabase_Err
        

    Dim temp As Variant

    temp = GetUserValue(name, "account_id")
    
    If VBA.IsEmpty(temp) Then
        GetAccountIDDatabase = -1
    Else
        GetAccountIDDatabase = val(temp)

    End If

        
    Exit Function

GetAccountIDDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetAccountIDDatabase", Erl)
    Resume Next
        
End Function

Public Sub GetPasswordAndSaltDatabase(CuentaEmail As String, PasswordHash As String, Salt As String)

    On Error GoTo ErrorHandler
    
    Call MakeQuery("SELECT password, salt FROM account WHERE deleted = FALSE AND email = ?;", False, LCase$(CuentaEmail))

    If QueryData Is Nothing Then Exit Sub
    
    PasswordHash = QueryData!Password
    Salt = QueryData!Salt
    
    Exit Sub
    
ErrorHandler:
    Call LogDatabaseError("Error in GetPasswordAndSaltDatabase. Cuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)
    
End Sub

Public Function GetPersonajesCountDatabase(CuentaEmail As String) As Byte

    On Error GoTo ErrorHandler

    Dim Id As Integer

    Id = GetDBValue("account", "id", "email", LCase$(CuentaEmail))
    
    GetPersonajesCountDatabase = GetPersonajesCountByIDDatabase(Id)
    
    Exit Function
    
ErrorHandler:
    Call LogDatabaseError("Error in GetPersonajesCountDatabase. name: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)
    
End Function

Public Function GetPersonajesCountByIDDatabase(ByVal AccountId As Long) As Byte

    On Error GoTo ErrorHandler
    
    Call MakeQuery("SELECT COUNT(*) FROM user WHERE deleted = FALSE AND account_id = ?;", False, AccountId)
    
    If QueryData Is Nothing Then Exit Function
    
    GetPersonajesCountByIDDatabase = QueryData.Fields(0).Value
    
    Exit Function
    
ErrorHandler:
    Call LogDatabaseError("Error in GetPersonajesCountByIDDatabase. AccountID: " & AccountId & ". " & Err.Number & " - " & Err.Description)
    
End Function

Public Function GetPersonajesCuentaDatabase(ByVal AccountId As Long, Personaje() As PersonajeCuenta) As Byte
        
    On Error GoTo GetPersonajesCuentaDatabase_Err
        

    Call MakeQuery("SELECT name, head_id, class_id, body_id, pos_map, level, status, helmet_id, shield_id, weapon_id, guild_index, is_dead, is_sailing FROM user WHERE deleted = FALSE AND account_id = ?;", False, AccountId)

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

        If val(QueryData!is_dead) = 1 Or val(QueryData!is_sailing) = 1 Then
            Personaje(i).Cabeza = 0
        End If
        
        QueryData.MoveNext
    Next

        
    Exit Function

GetPersonajesCuentaDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.GetPersonajesCuentaDatabase", Erl)
    Resume Next
        
End Function

Public Sub SetUserLoggedDatabase(ByVal Id As Long, ByVal AccountId As Long)
        
    On Error GoTo SetUserLoggedDatabase_Err
        
    Call SetDBValue("user", "is_logged", 1, "id", Id)
    Call MakeQuery("UPDATE account SET logged = logged + 1 WHERE id = ?;", True, AccountId)

        
    Exit Sub

SetUserLoggedDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetUserLoggedDatabase", Erl)
    Resume Next
        
End Sub

Public Sub ResetLoggedDatabase(ByVal AccountId As Long)
        
    On Error GoTo ResetLoggedDatabase_Err
        
    Call MakeQuery("UPDATE account SET logged = 0 WHERE id = ?;", True, AccountId)

        
    Exit Sub

ResetLoggedDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.ResetLoggedDatabase", Erl)
    Resume Next
        
End Sub

Public Sub SetUsersLoggedDatabase(ByVal NumUsers As Long)
        
    On Error GoTo SetUsersLoggedDatabase_Err
        
    Call MakeQuery("UPDATE statistics SET value = ? WHERE name = 'online';", True, NumUsers)
        
    Exit Sub

SetUsersLoggedDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetUsersLoggedDatabase", Erl)
    Resume Next
        
End Sub

Public Function LeerRecordUsuariosDatabase() As Long
        
    On Error GoTo LeerRecordUsuariosDatabase_Err
        
    Call MakeQuery("SELECT value FROM statistics WHERE name = 'record';", False)

    If QueryData Is Nothing Then Exit Function

    LeerRecordUsuariosDatabase = val(QueryData!Value)

    Exit Function

LeerRecordUsuariosDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.LeerRecordUsuariosDatabase", Erl)
    Resume Next
        
End Function

Public Sub SetRecordUsersDatabase(ByVal Record As Long)
        
    On Error GoTo SetRecordUsersDatabase_Err
        
    Call MakeQuery("UPDATE statistics SET value = ? WHERE name = 'record';", True, CStr(Record))
        
    Exit Sub

SetRecordUsersDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.SetRecordUsersDatabase", Erl)
    Resume Next
        
End Sub

Public Sub LogoutAllUsersAndAccounts()
        
    On Error GoTo LogoutAllUsersAndAccounts_Err

    Call MakeQuery("UPDATE user SET is_logged = FALSE; UPDATE account SET logged = 0;", True)
        
    Exit Sub

LogoutAllUsersAndAccounts_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.LogoutAllUsersAndAccounts", Erl)
    Resume Next
        
End Sub

Public Sub SaveVotoDatabase(ByVal Id As Long, ByVal Encuestas As Integer)
        
    On Error GoTo SaveVotoDatabase_Err
        
    Call SetUserValueByID(Id, "votes_amount", Encuestas)

        
    Exit Sub

SaveVotoDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.SaveVotoDatabase", Erl)
    Resume Next
        
End Sub

Public Sub SaveUserBodyDatabase(UserName As String, ByVal Body As Integer)
        
    On Error GoTo SaveUserBodyDatabase_Err
        
    Call SetUserValue(UserName, "body_id", Body)

        
    Exit Sub

SaveUserBodyDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.SaveUserBodyDatabase", Erl)
    Resume Next
        
End Sub

Public Sub SaveUserHeadDatabase(UserName As String, ByVal Head As Integer)
        
    On Error GoTo SaveUserHeadDatabase_Err
        
    Call SetUserValue(UserName, "head_id", Head)

        
    Exit Sub

SaveUserHeadDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.SaveUserHeadDatabase", Erl)
    Resume Next
        
End Sub

Public Sub SaveUserSkillDatabase(UserName As String, ByVal Skill As Integer, ByVal Value As Integer)
        
    On Error GoTo SaveUserSkillDatabase_Err
        
    Call MakeQuery("UPDATE skillpoints SET value = ? WHERE number = ? AND user_id = (SELECT id FROM user WHERE UPPER(name) = ?);", True, Value, Skill, UCase$(UserName))
        
    Exit Sub

SaveUserSkillDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.SaveUserSkillDatabase", Erl)
    Resume Next
        
End Sub

Public Sub SaveUserSkillsLibres(UserName As String, ByVal SkillsLibres As Integer)
        
    On Error GoTo SaveUserHeadDatabase_Err
        
    Call SetUserValue(UserName, "free_skillpoints", SkillsLibres)
        
    Exit Sub

SaveUserHeadDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.SaveUserHeadDatabase", Erl)
    Resume Next
        
End Sub

Public Sub SaveNewAccountDatabase(CuentaEmail As String, PasswordHash As String, Salt As String, Codigo As String)

    On Error GoTo ErrorHandler
    
    Call MakeQuery("INSERT INTO account SET email = ?, password = ?, salt = ?, validate_code = ?, date_created = NOW();", True, LCase$(CuentaEmail), PasswordHash, Salt, Codigo)
    
    Exit Sub
        
ErrorHandler:
    Call LogDatabaseError("Error en SaveNewAccountDatabase. Cuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub ValidarCuentaDatabase(UserCuenta As String)
        
    On Error GoTo ValidarCuentaDatabase_Err
        
    Call SetCuentaValue(UserCuenta, "validated", 1)
        
    Exit Sub

ValidarCuentaDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.ValidarCuentaDatabase", Erl)
    Resume Next
        
End Sub

Public Function CheckUserAccount(name As String, ByVal AccountId As Integer) As Boolean

    CheckUserAccount = (val(GetUserValue(name, "account_id")) = AccountId)

End Function

Public Sub BorrarUsuarioDatabase(name As String)

    On Error GoTo ErrorHandler
    
    Call MakeQuery("UPDATE user SET name = CONCAT('DELETED_', name), deleted = TRUE WHERE UPPER(name) = ?;", True, UCase$(name))

    Exit Sub
    
ErrorHandler:
    Call LogDatabaseError("Error en BorrarUsuarioDatabase borrando user de la Mysql Database: " & name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub BorrarCuentaDatabase(CuentaEmail As String)

    On Error GoTo ErrorHandler

    Dim Id As Integer

    Id = GetDBValue("account", "id", "email", LCase$(CuentaEmail))

    Call MakeQuery("UPDATE account SET email = CONCAT('DELETED_', email), deleted = TRUE WHERE email = ?;", True, LCase$(CuentaEmail))

    Call MakeQuery("UPDATE user SET name = CONCAT('DELETED_', name), deleted = TRUE WHERE account_id = ?;", True, Id)

    Exit Sub
    
ErrorHandler:
    Call LogDatabaseError("Error en BorrarCuentaDatabase borrando user de la Mysql Database: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveBanDatabase(UserName As String, Reason As String, BannedBy As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call MakeQuery("UPDATE user SET is_banned = TRUE WHERE UPPER(name) = ?;", True, UCase$(UserName))

    query = "INSERT INTO punishment SET "
    query = query & "user_id = (SELECT id from user WHERE UPPER(name) = ?), "
    query = query & "number = number + 1, "
    query = query & "reason = ?;"

    Call MakeQuery(query, True, UCase$(UserName), BannedBy & ": " & Reason & " " & Date & " " & Time)

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in SaveBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveWarnDatabase(UserName As String, Reason As String, WarnedBy As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call MakeQuery("UPDATE user SET warnings = warnings + 1 WHERE UPPER(name) = ?;", True, UCase$(UserName))

    query = "INSERT INTO punishment SET "
    query = query & "user_id = (SELECT id from user WHERE UPPER(name) = ?), "
    query = query & "number = number + 1, "
    query = query & "reason = ?;"

    Call MakeQuery(query, True, UCase$(UserName), WarnedBy & ": " & Reason & " " & Date & " " & Time)

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in SaveWarnDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SavePenaDatabase(UserName As String, Reason As String)

    On Error GoTo ErrorHandler

    Dim query As String

    query = query & "INSERT INTO punishment SET "
    query = query & "user_id = (SELECT id from user WHERE UPPER(name) = ?), "
    query = query & "number = number + 1, "
    query = query & "reason = ?;"

    Call MakeQuery(query, True, UCase$(UserName), Reason)

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in SavePenaDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub UnBanDatabase(UserName As String)

    On Error GoTo ErrorHandler

    Call MakeQuery("UPDATE user SET is_banned = FALSE WHERE UPPER(name) = ?;", True, UCase$(UserName))

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in UnBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveBanCuentaDatabase(ByVal AccountId As Integer, Reason As String, BannedBy As String)

    On Error GoTo ErrorHandler

    Call MakeQuery("UPDATE account SET is_banned = TRUE, banned_by = ?, ban_reason = ? WHERE id = ?;", True, BannedBy, Reason, AccountId)

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in SaveBanCuentaDatabase: AccountId=" & AccountId & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub EcharConsejoDatabase(UserName As String)
        
    On Error GoTo EcharConsejoDatabase_Err
        
    Call MakeQuery("UPDATE user SET pertenece_consejo_real = FALSE, pertenece_consejo_caos = FALSE WHERE UPPER(name) = ?;", True, UCase$(UserName))

        
    Exit Sub

EcharConsejoDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.EcharConsejoDatabase", Erl)
    Resume Next
        
End Sub

Public Sub EcharLegionDatabase(UserName As String)
        
    On Error GoTo EcharLegionDatabase_Err
        
    Call MakeQuery("UPDATE user SET pertenece_caos = FALSE, reenlistadas = 200 WHERE UPPER(name) = ?;", True, UCase$(UserName))

        
    Exit Sub

EcharLegionDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.EcharLegionDatabase", Erl)
    Resume Next
        
End Sub

Public Sub EcharArmadaDatabase(UserName As String)
        
    On Error GoTo EcharArmadaDatabase_Err
        
    Call MakeQuery("UPDATE user SET pertenece_real = FALSE, reenlistadas = 200 WHERE UPPER(name) = ?;", True, UCase$(UserName))

        
    Exit Sub

EcharArmadaDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.EcharArmadaDatabase", Erl)
    Resume Next
        
End Sub

Public Sub CambiarPenaDatabase(UserName As String, ByVal Numero As Integer, Pena As String)
        
    On Error GoTo CambiarPenaDatabase_Err
        
    Call MakeQuery("UPDATE punishment SET reason = ? WHERE number = ? AND user_id = (SELECT id from user WHERE UPPER(name) = ?);", True, Pena, Numero, UCase$(UserName))

        
    Exit Sub

CambiarPenaDatabase_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.CambiarPenaDatabase", Erl)
    Resume Next
        
End Sub

Public Function GetUserAmountOfPunishmentsDatabase(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Call MakeQuery("SELECT COUNT(*) as punishments FROM `punishment` INNER JOIN `user` ON punishment.user_id = user.id WHERE UPPER(user.name) = ?;", False, UCase$(UserName))

    If QueryData Is Nothing Then Exit Function

    GetUserAmountOfPunishmentsDatabase = QueryData!punishments

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserAmountOfPunishmentsDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Sub SendUserPunishmentsDatabase(ByVal UserIndex As Integer, ByVal UserName As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Call MakeQuery("SELECT user_id, number, reason FROM `punishment` INNER JOIN `user` ON punishment.user_id = user.id WHERE UPPER(user.name) = ?;", False, UCase$(UserName))
    
    If QueryData Is Nothing Then Exit Sub

    If Not QueryData.RecordCount = 0 Then
        QueryData.MoveFirst

        While Not QueryData.EOF
            Call WriteConsoleMsg(UserIndex, QueryData!Number & " - " & QueryData!Reason, FontTypeNames.FONTTYPE_INFO)
            
            QueryData.MoveNext
        Wend

    End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserPunishmentsDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub


Public Function GetNombreCuentaDatabase(name As String) As String

    On Error GoTo ErrorHandler

    'Hacemos la query.
    Call MakeQuery("SELECT email FROM `account` INNER JOIN `user` ON user.account_id = account.id WHERE UPPER(user.name) = ?;", False, UCase$(name))
    
    'Verificamos que la query no devuelva un resultado vacio.
    If QueryData Is Nothing Then Exit Function
    
    'Obtenemos el nombre de la cuenta
    GetNombreCuentaDatabase = QueryData!name

    Exit Function
    
ErrorHandler:
    Call LogDatabaseError("Error en GetNombreCuentaDatabase leyendo user de la Mysql Database: " & name & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildIndexDatabase(UserName As String) As Integer

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 09/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    GetUserGuildIndexDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_index"), 0)

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildMemberDatabase(UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    GetUserGuildMemberDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_member_history"), vbNullString)

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildAspirantDatabase(UserName As String) As Integer

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    GetUserGuildAspirantDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_aspirant_index"), 0)

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildRejectionReasonDatabase(UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    GetUserGuildRejectionReasonDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_rejected_because"), vbNullString)

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildPedidosDatabase(UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    GetUserGuildPedidosDatabase = SanitizeNullValue(GetUserValue(UserName, "guild_requests_history"), vbNullString)

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Sub SaveUserGuildRejectionReasonDatabase(UserName As String, Reason As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Call SetUserValue(UserName, "guild_rejected_because", Reason)

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildIndexDatabase(ByVal UserName As String, ByVal GuildIndex As Integer)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Call SetUserValue(UserName, "guild_index", GuildIndex)

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildAspirantDatabase(ByVal UserName As String, ByVal AspirantIndex As Integer)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Call SetUserValue(UserName, "guild_aspirant_index", AspirantIndex)

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildMemberDatabase(ByVal UserName As String, ByVal guilds As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Call SetUserValue(UserName, "guild_member_history", guilds)

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildPedidosDatabase(ByVal UserName As String, ByVal Pedidos As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Call SetUserValue(UserName, "guild_requests_history", Pedidos)

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

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

    Call MakeQuery("SELECT race_id, class_id, genre_id, level, gold, bank_gold, guild_requests_history, guild_index, guild_member_history, pertenece_real, pertenece_caos, ciudadanos_matados, criminales_matados FROM user WHERE UPPER(name) = ?;", False, UCase$(UserName))

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
    Call LogDatabaseError("Error in SendCharacterInfoDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function EnterAccountDatabase(ByVal UserIndex As Integer, CuentaEmail As String, Password As String, MacAddress As String, ByVal HDserial As Long, ip As String) As Boolean

    On Error GoTo ErrorHandler
    
    Call MakeQuery("SELECT id, password, salt, validated, is_banned, ban_reason, banned_by FROM account WHERE email = ?;", False, LCase$(CuentaEmail))
    
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
    
    UserList(UserIndex).AccountId = QueryData!Id
    UserList(UserIndex).Cuenta = CuentaEmail
    
    Call MakeQuery("UPDATE account SET mac_address = ?, hd_serial = ?, last_ip = ?, last_access = NOW() WHERE id = ?;", True, MacAddress, HDserial, ip, QueryData!Id)
    
    EnterAccountDatabase = True
    
    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in EnterAccountDatabase. UserCuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function PersonajePerteneceEmail(ByVal UserName As String, ByVal AccountEmail As String) As Boolean
    
    Call MakeQuery("SELECT id FROM user INNER JOIN account ON user.account_id = account.id WHERE user.name = ? AND account.email = ?;", False, UserName, AccountEmail)
    
    If QueryData Is Nothing Then
        PersonajePerteneceEmail = False
        Exit Function
    End If
    
    PersonajePerteneceEmail = True
    
End Function

Public Function PersonajePerteneceID(ByVal UserName As String, ByVal AccountId As Integer) As Boolean
    
    Call MakeQuery("SELECT id FROM user WHERE name = ? AND account_id = ?;", False, UserName, AccountId)
    
    If QueryData Is Nothing Then
        PersonajePerteneceID = False
        Exit Function
    End If
    
    PersonajePerteneceID = True
    
End Function

Public Sub ChangePasswordDatabase(ByVal UserIndex As Integer, OldPassword As String, NewPassword As String)

    On Error GoTo ErrorHandler

    If LenB(NewPassword) = 0 Then
        Call WriteConsoleMsg(UserIndex, "Debe especificar una contraseña nueva, inténtelo de nuevo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    Call MakeQuery("SELECT password, salt FROM account WHERE id = ?;", False, UserList(UserIndex).AccountId)
    
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
    
    Call MakeQuery("UPDATE account SET password = ?, salt = ? WHERE id = ?;", True, PasswordHash, Salt, UserList(UserIndex).AccountId)
    
    Call WriteConsoleMsg(UserIndex, "La contraseña de su cuenta fue cambiada con éxito.", FontTypeNames.FONTTYPE_INFO)
    
    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in ChangePasswordDatabase. Username: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function GetUsersLoggedAccountDatabase(ByVal AccountId As Integer) As Byte

    On Error GoTo ErrorHandler

    Call GetDBValue("account", "logged", "id", AccountId)
    
    If QueryData Is Nothing Then Exit Function
    
    GetUsersLoggedAccountDatabase = val(QueryData!logged)

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUsersLoggedAccountDatabase. AccountID: " & AccountId & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function SetPositionDatabase(UserName As String, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    On Error GoTo ErrorHandler

    Call MakeQuery("UPDATE user SET pos_map = ?, pos_x = ?, pos_y = ? WHERE UPPER(name) = ?;", True, Map, X, Y, UCase$(UserName))
    
    SetPositionDatabase = RecordsAffected > 0

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in SetPositionDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetMapDatabase(UserName As String) As Integer
    On Error GoTo ErrorHandler

    GetMapDatabase = val(GetUserValue(UserName, "pos_map"))

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in SetPositionDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function AddOroBancoDatabase(UserName As String, ByVal OroGanado As Long) As Boolean
    On Error GoTo ErrorHandler

    Call MakeQuery("UPDATE user SET bank_gold = bank_gold + ? WHERE UPPER(name) = ?;", True, OroGanado, UCase$(UserName))
    
    AddOroBancoDatabase = RecordsAffected > 0

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in AddOroBancoDatabase. UserName: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function DarLlaveAUsuarioDatabase(UserName As String, ByVal LlaveObj As Integer) As Boolean
    On Error GoTo ErrorHandler

    Call MakeQuery("INSERT INTO house_key SET key_obj = ?, account_id = (SELECT account_id FROM user WHERE UPPER(name) = ?);", True, LlaveObj, UCase$(UserName))
    
    DarLlaveAUsuarioDatabase = RecordsAffected > 0

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in DarLlaveAUsuarioDatabase. UserName: " & UserName & ", LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function DarLlaveACuentaDatabase(email As String, ByVal LlaveObj As Integer) As Boolean
    On Error GoTo ErrorHandler

    Call MakeQuery("INSERT INTO house_key SET key_obj = ?, account_id = (SELECT id FROM account WHERE UPPER(email) = ?);", True, LlaveObj, UCase$(email))
    
    DarLlaveACuentaDatabase = RecordsAffected > 0
    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in DarLlaveACuentaDatabase. Email: " & email & ", LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function SacarLlaveDatabase(ByVal LlaveObj As Integer) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim UserCount As Integer
    Dim Users() As String

    ' Obtengo los usuarios logueados en la cuenta del dueño de la llave
    Call MakeQuery("SELECT name FROM `user` INNER JOIN `account` ON `user`.account_id = account.id INNER JOIN `house_key` ON `house_key`.account_id = account.id WHERE `user`.is_logged = TRUE AND `house_key`.key_obj = ?;", False, LlaveObj)
    
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
    Call MakeQuery("DELETE FROM house_key WHERE key_obj = ?;", True, LlaveObj)
    
    ' Si pudimos borrar, actualizamos los usuarios logueados
    Dim UserIndex As Integer
    
    For i = 1 To UserCount
        UserIndex = NameIndex(Users(i))
        
        If UserIndex <> 0 Then
            Call SacarLlaveDeLLavero(UserIndex, LlaveObj)
        End If
    Next
    
    SacarLlaveDatabase = RecordsAffected > 0

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in SacarLlaveDatabase. LlaveObj: " & LlaveObj & ". " & Err.Number & " - " & Err.Description)

End Function

Public Sub VerLlavesDatabase(ByVal UserIndex As Integer)
    On Error GoTo ErrorHandler

    Call MakeQuery("SELECT email, key_obj FROM `house_key` INNER JOIN `account` ON `house_key`.account_id = `account`.id;", False)

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
    Call LogDatabaseError("Error in VerLlavesDatabase. UserName: " & UserList(UserIndex).name & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function SanitizeNullValue(ByVal Value As Variant, ByVal defaultValue As Variant) As Variant
        
    On Error GoTo SanitizeNullValue_Err
        
    SanitizeNullValue = IIf(IsNull(Value), defaultValue, Value)

        
    Exit Function

SanitizeNullValue_Err:
    Call RegistrarError(Err.Number, Err.Description, "modDatabase.SanitizeNullValue", Erl)
    Resume Next
        
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
    Cmd.CommandText = "SELECT 1"
    Cmd.ActiveConnection = adoCn

    'Run a simple query, to test the connection
        
    i = Cmd.Execute.Fields(0)
    On Error GoTo 0

    'Tidy up
    Set Cmd = Nothing

    'If i is 1, connection is open
    If i = 1 Then
        adoIsConnected = True
    Else
        adoIsConnected = False
    End If

End Function
