VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Packets
    LOGIN_ACCOUNT
    LOGIN_CHAR
    CREATE_CHAR
    SAVE_CHAR
    GUARD_NOTICE_RESPONSE
End Enum

Private Enum ManagerPackets
    LOGIN_ACCOUNT_RESULT
    LOGIN_CHAR_RESULT
    CREATE_CHAR_RESULT
    SAVE_CHAR_RESULT
    GUARD_NOTICE_RESPONSE_RESULT
End Enum

Private Closing As Boolean

Private WithEvents Socket As Winsock
Attribute Socket.VB_VarHelpID = -1
Private Port As Integer

Private Buffer As cStringBuilder

Public Sub Initialize(pSocket As Winsock, ByVal pPort As Integer)
    Set Socket = pSocket
    Port = pPort
    Set Buffer = New cStringBuilder
    
    Call Connect
End Sub

Private Sub Connect()
    Call Socket.Connect("localhost", Port)
End Sub

Public Sub Send(ByVal Packet As Packets, Data As JS_Object, Optional Instance As JS_Object)
    Dim Msg As New JS_Object

    Msg.Item("id") = Packet
    Msg.Item("data") = Data
    
    If Not Instance Is Nothing Then
        Msg.Item("instance") = Instance
    End If

    Call Socket.SendData(Chr(1) & Msg.ToString() & Chr(0))
End Sub

Private Sub Class_Terminate()
    Closing = True
    Socket.Close
    Set Socket = Nothing
End Sub

Private Sub Socket_Close()
    If Not Closing Then
        Call Connect
    End If
End Sub

Private Sub Socket_Connect()
    Debug.Print "Conectado con el Manager"
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim DataRecv As String
    
    Call Socket.getData(DataRecv, vbString, bytesTotal)
    Call Buffer.Append(DataRecv)

    Dim FirstSeparator As Long
    FirstSeparator = Buffer.Find(Chr(0))

    Dim message As String
    If FirstSeparator > 0 Then
        message = Buffer.SubStr(1, FirstSeparator - 2)

        Dim Data As Object
        Set Data = JSON.parse(message)

          Call HandlePacket(Data.Item("id"), Data.Item("data"), Data.Item("instance"))

        Call Buffer.Remove(0, FirstSeparator)
    End If
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Not Closing Then
        Socket.Close
        DoEvents

        Call Connect
    End If
End Sub

Private Sub HandlePacket(ByVal ID As Long, Data As Object, Instance As Object)

    On Error GoTo Handler

    Select Case ID
    
        Case ManagerPackets.LOGIN_ACCOUNT_RESULT
            Call HandleLoginAccount(ID, Data, Instance)
            
        Case ManagerPackets.LOGIN_CHAR_RESULT
            Call HandleLoginChar(ID, Data, Instance)
        
        Case ManagerPackets.CREATE_CHAR_RESULT
            Call HandleCreateChar(ID, Data, Instance)
            
        Case ManagerPackets.SAVE_CHAR_RESULT
            Call HandleSaveChar(ID, Data, Instance)
            
        Case ManagerPackets.GUARD_NOTICE_RESPONSE_RESULT
            Call HandleGuardNotice(ID, Data, Instance)
            
    End Select
    
    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.Description, "ServerManager.HandlePacket", Erl)
    
End Sub

Private Function CheckPlayer(ByVal Slot As Integer, UUID As String) As Boolean
    CheckPlayer = False

    If Slot <= 0 Or Slot > MaxUsers Then Exit Function

    With UserList(Slot)
        If .WaitingPacket < 0 Then Exit Function

        If .UUID <> UUID Then Exit Function
    End With
    
    CheckPlayer = True

End Function

Private Sub HandleLoginAccount(ByVal ID As Long, Data As Object, Instance As Object)
    Dim UserIndex As Integer
    UserIndex = Instance.Item("slot")

    ' Si no está online el jugador, ignoramos la respuesta
    If Not CheckPlayer(UserIndex, Instance.Item("uuid")) Then Exit Sub
    
    With UserList(UserIndex)
        .WaitingPacket = -1

        ' Si el login fue exitoso
        If Data.Item("success") Then
            ' Sólo si está logueando por primera vez recibimos estas variables
            If .AccountID < 0 Then
                ' Seteamos variables de la cuenta
                .AccountID = Data.Item("account_id")
                .Cuenta = .Email
                
                Dim i As Integer
                For i = 1 To Data.Item("keys").Count
                    .keys(i) = Data.Item("keys").Item(i)
                Next
            End If
        
            If Data.Item("verify") Then
                Call WriteGuardNotice(UserIndex)
                Call WriteShowMessageBox(UserIndex, Data.Item("message"))
            Else
                Call WritePersonajesDeCuenta(UserIndex, Data.Item("chars"))
            End If
    
        ' Si no pudo loguearse, le mostramos el mensaje de error y cerramos el socket
        Else
            Call WriteShowMessageBox(UserIndex, Data.Item("message"))
            Call CloseSocket(UserIndex)
        End If

    End With
End Sub

Private Sub HandleGuardNotice(ByVal ID As Long, Data As Object, Instance As Object)
    Dim UserIndex As Integer
    UserIndex = Instance.Item("slot")

    ' Si no está online el jugador, ignoramos la respuesta
    If Not CheckPlayer(UserIndex, Instance.Item("uuid")) Then Exit Sub
    
    UserList(UserIndex).WaitingPacket = -1

    ' Si el login fue exitoso
    If Data.Item("success") Then
        Call WritePersonajesDeCuenta(UserIndex, Data.Item("chars"))

    ' Si no pudo loguearse, le mostramos el mensaje de error y cerramos el socket
    Else
        Call WriteShowMessageBox(UserIndex, Data.Item("message"))
        Call CloseSocket(UserIndex)
    End If
End Sub

Private Sub HandleLoginChar(ByVal ID As Long, Data As Object, Instance As Object)
    Dim UserIndex As Integer
    UserIndex = Instance.Item("slot")

    ' Si no está online el jugador, ignoramos la respuesta
    If Not CheckPlayer(UserIndex, Instance.Item("uuid")) Then Exit Sub

    With UserList(UserIndex)
        .WaitingPacket = -1

        ' Si el login fue exitoso
        If Data.Item("success") Then
            Dim i As Long
            Dim LoopC As Long
        
            ' Llenar el slot con los datos
            .ID = Data.Item("id")
            .AccountID = Data.Item("account_id")
            .Name = Data.Item("name")
            .genero = Data.Item("genre_id")
            .clase = Data.Item("class_id")
            .raza = Data.Item("race_id")
            .Hogar = Data.Item("home_id")
            .Desc = Data.Item("description")
            .MENSAJEINFORMACION = Data.Item("message_info")
            
            ' Chat
            .ChatGlobal = Data.Item("chat_global")
            .ChatCombate = Data.Item("chat_combate")
            
            ' Character
            .OrigChar.Body = Data.Item("body_id")
            .OrigChar.Head = Data.Item("head_id")
            .OrigChar.WeaponAnim = Data.Item("weapon_id")
            .OrigChar.CascoAnim = Data.Item("helmet_id")
            .OrigChar.ShieldAnim = Data.Item("shield_id")
            .OrigChar.Heading = Data.Item("heading")
            
            ' Position
            .Pos.Map = Data.Item("pos_map")
            .Pos.X = Data.Item("pos_x")
            .Pos.Y = Data.Item("pos_y")
            
            ' Stats
            .Stats.ELV = Data.Item("level")
            .Stats.Exp = Data.Item("exp")
            .Stats.MinHp = Data.Item("min_hp")
            .Stats.MaxHp = Data.Item("max_hp")
            .Stats.MinMAN = Data.Item("min_man")
            .Stats.MaxMAN = Data.Item("max_man")
            .Stats.MinSta = Data.Item("min_sta")
            .Stats.MaxSta = Data.Item("max_sta")
            .Stats.MinHIT = Data.Item("min_hit")
            .Stats.MaxHit = Data.Item("max_hit")
            .Stats.MinHam = Data.Item("min_ham")
            .Stats.MaxHam = Data.Item("max_ham")
            .Stats.MinAGU = Data.Item("min_sed")
            .Stats.MaxAGU = Data.Item("max_sed")
            .Stats.GLD = Data.Item("gold")
            .Stats.Banco = Data.Item("bank_gold")
            .Stats.SkillPts = Data.Item("free_skillpoints")
            .Stats.NPCsMuertos = Data.Item("killed_npcs")
            .Stats.UsuariosMatados = Data.Item("killed_users")
            .Stats.CriminalesMatados = Data.Item("criminales_matados")
            .Stats.InventLevel = Data.Item("invent_level")
            .Stats.ELO = Data.Item("elo")
            .Stats.Advertencias = Data.Item("warnings")
            
            ' Inventario
            .Invent.ArmourEqpSlot = Data.Item("slot_armour")
            .Invent.WeaponEqpSlot = Data.Item("slot_weapon")
            .Invent.CascoEqpSlot = Data.Item("slot_helmet")
            .Invent.EscudoEqpSlot = Data.Item("slot_shield")
            .Invent.MunicionEqpSlot = Data.Item("slot_ammo")
            .Invent.BarcoSlot = Data.Item("slot_ship")
            .Invent.MonturaSlot = Data.Item("slot_mount")
            .Invent.DañoMagicoEqpSlot = Data.Item("slot_dm")
            .Invent.ResistenciaEqpSlot = Data.Item("slot_rm")
            .Invent.MagicoSlot = Data.Item("slot_magic")
            .Invent.NudilloSlot = Data.Item("slot_knuckles")
            .Invent.HerramientaEqpSlot = Data.Item("slot_tool")
            
            ' Flags
            .flags.Desnudo = Data.Item("is_naked")
            .flags.Envenenado = Data.Item("is_poisoned")
            .flags.Incinerado = Data.Item("is_incinerated")
            .flags.Escondido = False
            .flags.Hambre = (.Stats.MinHam = 0)
            .flags.Sed = (.Stats.MinAGU = 0)
            .flags.Ban = Data.Item("is_banned")
            .flags.BanMotivo = Data.Item("ban_reason")
            .flags.Muerto = (.Stats.MinHp = 0)
            .flags.Navegando = Data.Item("is_sailing")
            .flags.Paralizado = Data.Item("is_paralyzed")
            .flags.VecesQueMoriste = Data.Item("deaths")
            .flags.Montado = Data.Item("is_mounted")
            .flags.Pareja = Data.Item("spouse")
            .flags.Casado = IIf(Len(.flags.Pareja) > 0, 1, 0)
            .flags.Silenciado = Data.Item("is_silenced")
            .flags.MinutosRestantes = Data.Item("silence_minutes_left")
            .flags.SegundosPasados = Data.Item("silence_elapsed_seconds")
            .flags.MascotasGuardadas = Data.Item("pets_saved")
            .flags.ReturnPos.Map = Data.Item("return_map")
            .flags.ReturnPos.X = Data.Item("return_x")
            .flags.ReturnPos.Y = Data.Item("return_y")
            
            If Data.Item("pertenece_consejo_real") Then
                 .flags.Privilegios = .flags.Privilegios Or e_PlayerType.RoyalCouncil
            
            ElseIf Data.Item("pertenece_consejo_caos") Then
                .flags.Privilegios = .flags.Privilegios Or e_PlayerType.ChaosCouncil
                
            End If

            ' Contadores
            .Counters.Pena = Data.Item("counter_pena")
            
            ' Faccion
            .Faccion.ArmadaReal = Data.Item("pertenece_real")
            .Faccion.FuerzasCaos = Data.Item("pertenece_caos")
            .Faccion.ciudadanosMatados = Data.Item("ciudadanos_matados")
            .Faccion.CriminalesMatados = Data.Item("criminales_matados")
            .Faccion.RecibioArmaduraReal = Data.Item("recibio_armadura_real")
            .Faccion.RecibioArmaduraCaos = Data.Item("recibio_armadura_caos")
            .Faccion.RecibioExpInicialReal = Data.Item("recibio_exp_real")
            .Faccion.RecibioExpInicialCaos = Data.Item("recibio_exp_caos")
            .Faccion.RecompensasReal = Data.Item("recompensas_real")
            .Faccion.RecompensasCaos = Data.Item("recompensas_caos")
            .Faccion.Reenlistadas = Data.Item("Reenlistadas")
            .Faccion.NivelIngreso = Data.Item("nivel_ingreso")
            .Faccion.MatadosIngreso = Data.Item("matados_ingreso")
            .Faccion.NextRecompensa = Data.Item("siguiente_recompensa")
            .Faccion.Status = Data.Item("status")
            
            .GuildIndex = Data.Item("guild_Index")
            .LastGuildRejection = Data.Item("guild_rejected_because")
            
            ' Atributos
            .Stats.UserAtributos(e_Atributos.Fuerza) = Data.Item("Attribute").Item("strength")
            .Stats.UserAtributos(e_Atributos.Agilidad) = Data.Item("Attribute").Item("agility")
            .Stats.UserAtributos(e_Atributos.Constitucion) = Data.Item("Attribute").Item("constitution")
            .Stats.UserAtributos(e_Atributos.Inteligencia) = Data.Item("Attribute").Item("intelligence")
            .Stats.UserAtributos(e_Atributos.Carisma) = Data.Item("Attribute").Item("charisma")

            .Stats.UserAtributosBackUP(e_Atributos.Fuerza) = .Stats.UserAtributos(e_Atributos.Fuerza)
            .Stats.UserAtributosBackUP(e_Atributos.Agilidad) = .Stats.UserAtributos(e_Atributos.Agilidad)
            .Stats.UserAtributosBackUP(e_Atributos.Constitucion) = .Stats.UserAtributos(e_Atributos.Constitucion)
            .Stats.UserAtributosBackUP(e_Atributos.Inteligencia) = .Stats.UserAtributos(e_Atributos.Inteligencia)
            .Stats.UserAtributosBackUP(e_Atributos.Carisma) = .Stats.UserAtributos(e_Atributos.Carisma)
            
            ' Habilidades
            .Stats.UserSkills(e_Skill.Magia) = Data.Item("Skillpoint").Item("magia")
            .Stats.UserSkills(e_Skill.Robar) = Data.Item("Skillpoint").Item("robar")
            .Stats.UserSkills(e_Skill.Tacticas) = Data.Item("Skillpoint").Item("tacticas")
            .Stats.UserSkills(e_Skill.Armas) = Data.Item("Skillpoint").Item("armas")
            .Stats.UserSkills(e_Skill.Meditar) = Data.Item("Skillpoint").Item("meditar")
            .Stats.UserSkills(e_Skill.Apuñalar) = Data.Item("Skillpoint").Item("apunalar")
            .Stats.UserSkills(e_Skill.Ocultarse) = Data.Item("Skillpoint").Item("ocultarse")
            .Stats.UserSkills(e_Skill.Supervivencia) = Data.Item("Skillpoint").Item("supervivencia")
            .Stats.UserSkills(e_Skill.Comerciar) = Data.Item("Skillpoint").Item("comerciar")
            .Stats.UserSkills(e_Skill.Defensa) = Data.Item("Skillpoint").Item("defensa")
            .Stats.UserSkills(e_Skill.liderazgo) = Data.Item("Skillpoint").Item("liderazgo")
            .Stats.UserSkills(e_Skill.Proyectiles) = Data.Item("Skillpoint").Item("proyectiles")
            .Stats.UserSkills(e_Skill.Wrestling) = Data.Item("Skillpoint").Item("wrestling")
            .Stats.UserSkills(e_Skill.Navegacion) = Data.Item("Skillpoint").Item("navegacion")
            .Stats.UserSkills(e_Skill.equitacion) = Data.Item("Skillpoint").Item("equitacion")
            .Stats.UserSkills(e_Skill.Resistencia) = Data.Item("Skillpoint").Item("resistencia")
            .Stats.UserSkills(e_Skill.Talar) = Data.Item("Skillpoint").Item("talar")
            .Stats.UserSkills(e_Skill.Pescar) = Data.Item("Skillpoint").Item("pescar")
            .Stats.UserSkills(e_Skill.Mineria) = Data.Item("Skillpoint").Item("mineria")
            .Stats.UserSkills(e_Skill.Herreria) = Data.Item("Skillpoint").Item("herreria")
            .Stats.UserSkills(e_Skill.Carpinteria) = Data.Item("Skillpoint").Item("carpinteria")
            .Stats.UserSkills(e_Skill.Alquimia) = Data.Item("Skillpoint").Item("alquimia")
            .Stats.UserSkills(e_Skill.Sastreria) = Data.Item("Skillpoint").Item("sastreria")
            .Stats.UserSkills(e_Skill.Domar) = Data.Item("Skillpoint").Item("domar")
            
            ' Mascotas
            For i = 1 To Data.Item("Pets").Count
                .MascotasType(i) = Data.Item("Pets").Item(i).Item("pet_id")
                
                If .MascotasType(i) > 0 Then
                    .NroMascotas = .NroMascotas + 1
                End If
            Next
            
            ' Hechizos
            For i = 1 To Data.Item("Spells").Count
                .Stats.UserHechizos(i) = Data.Item("Spells").Item(i).Item("spell_id")
            Next
            
            ' Inventario
            LoopC = 0
            For i = 1 To Data.Item("InventoryItems").Count
            
                With .Invent.Object(i)
                
                    .ObjIndex = Data.Item("InventoryItems").Item(i).Item("item_id")
                    
                    If .ObjIndex <> 0 Then
                    
                        If LenB(ObjData(.ObjIndex).Name) Then
                            LoopC = LoopC + 1
                            
                            .amount = Data.Item("InventoryItems").Item(i).Item("amount")
                            .Equipped = Data.Item("InventoryItems").Item(i).Item("is_equipped")
                        Else
                            .ObjIndex = 0
                        End If
                        
                    End If
                    
                End With
                
            Next
            .Invent.NroItems = LoopC
            
            ' Inventario de la Boveda
            LoopC = 0
            For i = 1 To Data.Item("BankItems").Count
            
                With .BancoInvent.Object(i)
                
                    .ObjIndex = Data.Item("BankItems").Item(i).Item("item_id")
                    
                    If .ObjIndex <> 0 Then
                    
                        If LenB(ObjData(.ObjIndex).Name) Then
                            LoopC = LoopC + 1
                            
                            .amount = Data.Item("BankItems").Item(i).Item("amount")
                        Else
                            .ObjIndex = 0
                        End If
                        
                    End If
                    
                End With
                
            Next
            .BancoInvent.NroItems = LoopC
            
            ' Misiones
            LoopC = 0
            For i = 1 To Data.Item("Quests").Count
                With .QuestStats.Quests(i)
                
                    .QuestIndex = Data.Item("Quests").Item(i).Item("quest_id")
                    
                    If .QuestIndex > 0 Then
                    
                        If QuestList(.QuestIndex).RequiredNPCs Then
                            Dim NPCs() As String: NPCs = Split(Data.Item("Quests").Item(i).Item("NPCs"), "-")
                            ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
                            
                            For LoopC = 1 To QuestList(.QuestIndex).RequiredNPCs
                                .NPCsKilled(LoopC) = val(NPCs(LoopC - 1))
                            Next LoopC
                        End If
                        
                        If QuestList(.QuestIndex).RequiredTargetNPCs Then
                            Dim NPCsTarget() As String: NPCsTarget = Split(Data.Item("Quests").Item(i).Item("NPCsTarget"), "-")
                            ReDim .NPCsTarget(1 To QuestList(.QuestIndex).RequiredTargetNPCs)
                            
                            For LoopC = 1 To QuestList(.QuestIndex).RequiredTargetNPCs
                                .NPCsTarget(LoopC) = val(NPCsTarget(LoopC - 1))
                            Next LoopC
                        End If
                        
                    End If
                    
                End With
            Next i
            
            ' Misiones Completadas
            With .QuestStats
            
                .NumQuestsDone = Data.Item("QuestDones").Count
                
                If .NumQuestsDone > 0 Then
                    ReDim .QuestsDone(1 To .NumQuestsDone)
                    
                    For i = 1 To Data.Item("QuestDones").Count
                        .QuestsDone(i) = Data.Item("QuestDones").Item(i).Item("quest_id")
                    Next
                End If
                
            End With
            
            If ConnectUser_Check(UserIndex, .Name) Then
                Call ConnectUser_Complete(UserIndex, .Name)
            End If
        ' Si no pudo loguearse, le mostramos el mensaje de error y mantenemos el socket abierto
        Else
            Call WriteShowMessageBox(UserIndex, Data.Item("message"))
        End If
    
    End With
End Sub

Private Sub HandleCreateChar(ByVal ID As Long, Data As Object, Instance As Object)
    Dim UserIndex As Integer
    UserIndex = Instance.Item("slot")

    ' Si no está online el jugador, ignoramos la respuesta
    If Not CheckPlayer(UserIndex, Instance.Item("uuid")) Then Exit Sub
    
    With UserList(UserIndex)
        .WaitingPacket = -1

        ' Si el login fue exitoso
        If Data.Item("success") Then
            ' Llenar el slot con los datos
            .ID = Data.Item("id")
            
            UltimoChar = .Name
            
            If ConnectUser_Check(UserIndex, .Name) Then
                Call ConnectUser_Prepare(UserIndex, .Name)
                Call ConnectUser_Complete(UserIndex, .Name)
            End If
        ' Si no pudo loguearse, le mostramos el mensaje de error y mantenemos el socket abierto
        Else
            Call WriteShowMessageBox(UserIndex, Data.Item("message"))
        End If
        
    End With
End Sub

Private Sub HandleSaveChar(ByVal ID As Long, Data As Object, Instance As Object)
    Dim UserIndex As Integer
    UserIndex = Instance.Item("slot")

    ' Si no está online el jugador, ignoramos la respuesta
    If Not CheckPlayer(UserIndex, Instance.Item("uuid")) Then Exit Sub
    
    With UserList(UserIndex)
        .WaitingPacket = -1

        ' Si se guardó el personaje correctamente
        If Data.Item("success") Then
            ' Limpiamos el usuario
            Call UserClosed(UserIndex)

            Dim chars As Object
            Set chars = Data.Item("chars")

            ' Si recibimos chars
            If Not IsNull(chars) Then
                ' Se los enviamos
                Call WritePersonajesDeCuenta(UserIndex, chars)
            End If

        ' Si hubo un error interno, entonces el personaje no se pudo guardar (panic!)
        Else
            ' Reintentamos el guardado
            Call SaveUser(UserIndex, Instance.Item("logout"))
        End If

    End With
End Sub
