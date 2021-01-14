Attribute VB_Name = "JSON_User"
Option Explicit

Private Objeto As New JS_Object
Private Matriz As New JS_Array

Private i As Long

Function Principal(ByRef UserIndex As Integer, ByRef Logout As Boolean) As JS_Object
    
    Call Objeto.Clear
    
    With UserList(UserIndex)
    
        Objeto.Item("id") = .Id
        Objeto.Item("name") = .name
        Objeto.Item("level") = .Stats.ELV
        Objeto.Item("exp") = CLng(.Stats.Exp)
        Objeto.Item("elu") = .Stats.ELU
        Objeto.Item("genre_id") = .genero
        Objeto.Item("race_id") = .raza
        Objeto.Item("class_id") = .clase
        Objeto.Item("home_id") = .Hogar
        Objeto.Item("description") = .Desc
        Objeto.Item("gold") = .Stats.GLD
        Objeto.Item("bank_gold") = .Stats.Banco
        Objeto.Item("free_skillpoints") = .Stats.SkillPts
        Objeto.Item("pets_saved") = .flags.MascotasGuardadas
        Objeto.Item("pos_map") = .Pos.Map
        Objeto.Item("pos_x") = .Pos.X
        Objeto.Item("pos_y") = .Pos.Y
        Objeto.Item("last_map") = .flags.lastMap
        Objeto.Item("message_info") = .MENSAJEINFORMACION
        Objeto.Item("body_id") = .Char.Body
        Objeto.Item("head_id") = .OrigChar.Head
        Objeto.Item("weapon_id") = .Char.WeaponAnim
        Objeto.Item("helmet_id") = .Char.CascoAnim
        Objeto.Item("shield_id") = .Char.ShieldAnim
        Objeto.Item("heading") = .Char.Heading
        Objeto.Item("items_Amount") = .Invent.NroItems
        Objeto.Item("slot_armour") = .Invent.ArmourEqpSlot
        Objeto.Item("slot_weapon") = .Invent.WeaponEqpSlot
        Objeto.Item("slot_shield") = .Invent.EscudoEqpSlot
        Objeto.Item("slot_helmet") = .Invent.CascoEqpSlot
        Objeto.Item("slot_ammo") = .Invent.MunicionEqpSlot
        'Objeto.Item("slot_ring") = .Invent.AnilloEqpSlot
        Objeto.Item("slot_tool") = .Invent.HerramientaEqpSlot
        Objeto.Item("slot_magic") = .Invent.MagicoSlot
        Objeto.Item("slot_knuckles") = .Invent.NudilloSlot
        Objeto.Item("slot_ship") = .Invent.BarcoSlot
        Objeto.Item("slot_mount") = .Invent.MonturaSlot
        Objeto.Item("min_hp") = .Stats.MinHp
        Objeto.Item("max_hp") = .Stats.MaxHp
        Objeto.Item("min_man") = .Stats.MinMAN
        Objeto.Item("max_man") = .Stats.MaxMAN
        Objeto.Item("min_sta") = .Stats.MinSta
        Objeto.Item("max_sta") = .Stats.MaxSta
        Objeto.Item("min_ham") = .Stats.MinHam
        Objeto.Item("max_ham") = .Stats.MaxHam
        Objeto.Item("min_sed") = .Stats.MinAGU
        Objeto.Item("max_sed") = .Stats.MaxAGU
        Objeto.Item("min_hit") = .Stats.MinHIT
        Objeto.Item("max_hit") = .Stats.MaxHit
        Objeto.Item("killed_npcs") = .Stats.NPCsMuertos
        Objeto.Item("killed_users") = .Stats.UsuariosMatados
        Objeto.Item("invent_level") = .Stats.InventLevel
        'Objeto.Item( "rep_asesino") =.Reputacion.AsesinoRep
        'Objeto.Item( "rep_bandido") =.Reputacion.BandidoRep
        'Objeto.Item( "rep_burgues") =.Reputacion.BurguesRep
        'Objeto.Item( "rep_ladron") =.Reputacion.LadronesRep
        'Objeto.Item( "rep_noble") =.Reputacion.NobleRep
        'Objeto.Item( "rep_plebe") =.Reputacion.PlebeRep
        'Objeto.Item( "rep_average") =.Reputacion.Promedio
        Objeto.Item("is_naked") = .flags.Desnudo
        Objeto.Item("is_poisoned") = .flags.Envenenado
        Objeto.Item("is_hidden") = .flags.Escondido
        Objeto.Item("is_hungry") = .flags.Hambre
        Objeto.Item("is_thirsty") = .flags.Sed
        'Objeto.Item( "is_banned") =.flags.Ban & ") =" Esto es innecesario porque se setea cuando lo baneas
        Objeto.Item("is_dead") = .flags.Muerto
        Objeto.Item("is_sailing") = .flags.Navegando
        Objeto.Item("is_paralyzed") = .flags.Paralizado
        Objeto.Item("is_mounted") = .flags.Montado
        Objeto.Item("is_silenced") = .flags.Silenciado
        Objeto.Item("silence_minutes_left") = .flags.MinutosRestantes
        Objeto.Item("silence_elapsed_seconds") = .flags.SegundosPasados
        Objeto.Item("spouse") = .flags.Pareja
        Objeto.Item("counter_pena") = .Counters.Pena
        Objeto.Item("deaths") = .flags.VecesQueMoriste
        Objeto.Item("pertenece_consejo_real") = (.flags.Privilegios And PlayerType.RoyalCouncil)
        Objeto.Item("pertenece_consejo_caos") = (.flags.Privilegios And PlayerType.ChaosCouncil)
        Objeto.Item("pertenece_real") = .Faccion.ArmadaReal
        Objeto.Item("pertenece_caos") = .Faccion.FuerzasCaos
        Objeto.Item("ciudadanos_matados") = .Faccion.CiudadanosMatados
        Objeto.Item("criminales_matados") = .Faccion.CriminalesMatados
        Objeto.Item("recibio_armadura_real") = .Faccion.RecibioArmaduraReal
        Objeto.Item("recibio_armadura_caos") = .Faccion.RecibioArmaduraCaos
        Objeto.Item("recibio_exp_real") = .Faccion.RecibioExpInicialReal
        Objeto.Item("recibio_exp_caos") = .Faccion.RecibioExpInicialCaos
        Objeto.Item("recompensas_real") = .Faccion.RecompensasReal
        Objeto.Item("recompensas_caos") = .Faccion.RecompensasCaos
        Objeto.Item("reenlistadas") = .Faccion.Reenlistadas
        Objeto.Item("fecha_ingreso") = .Faccion.FechaIngreso
        Objeto.Item("nivel_ingreso") = .Faccion.NivelIngreso
        Objeto.Item("matados_ingreso") = .Faccion.MatadosIngreso
        Objeto.Item("siguiente_recompensa") = .Faccion.NextRecompensa
        Objeto.Item("status") = .Faccion.Status
        Objeto.Item("battle_points") = .flags.BattlePuntos
        Objeto.Item("guild_index") = .GuildIndex
        Objeto.Item("chat_combate") = .ChatCombate
        Objeto.Item("chat_global") = .ChatGlobal
        Objeto.Item("is_logged") = Not Logout
        Objeto.Item("warnings") = .Stats.Advertencias
    
    End With
    
    Set Principal = Objeto

End Function

Function Atributos(ByRef UserIndex As Integer) As JS_Array
    
    Call Objeto.Clear
    Call Matriz.Clear
    
    With UserList(UserIndex)
    
        For i = 1 To NUMATRIBUTOS
            Objeto.Item("user_id") = .Id
            Objeto.Item("number") = i
            Objeto.Item("spellId") = .Stats.UserAtributosBackUP(i)
    
            ' Lo meto en el array de items
            Matriz.Push Objeto
                    
            ' Limpio el objeto para la proxima iteracion
            Objeto.Clear
        Next i
    
    End With

    Set Atributos = Matriz

End Function

Function Hechizo(ByRef UserIndex As Integer) As JS_Array
    
    Call Objeto.Clear
    Call Matriz.Clear
    
    With UserList(UserIndex)
    
        For i = 1 To MAXUSERHECHIZOS
            Objeto.Item("user_id") = .Id
            Objeto.Item("number") = i
            Objeto.Item("spellId") = .Stats.UserHechizos(i)
                    
            ' Lo meto en el array de items
            Matriz.Push Objeto
                    
            ' Limpio el objeto para la proxima iteracion
            Objeto.Clear
        Next i
    
    End With

    Set Hechizo = Matriz

End Function

Function Inventario(ByRef UserIndex As Integer) As JS_Array
    
    Call Objeto.Clear
    Call Matriz.Clear
    
    With UserList(UserIndex)
    
        For i = 1 To .CurrentInventorySlots
            Objeto.Item("user_id") = .Id
            Objeto.Item("number") = i
            Objeto.Item("item_id") = .Invent.Object(i).ObjIndex
            Objeto.Item("amount") = .Invent.Object(i).Amount
            Objeto.Item("is_equipped") = .Invent.Object(i).Equipped
                
            ' Lo meto en el array de items
            Matriz.Push Objeto
                
            ' Limpio el objeto para la proxima iteracion
            Objeto.Clear
        Next i
    
    End With

    Set Inventario = Matriz

End Function

Function InventarioBanco(ByRef UserIndex As Integer) As JS_Array
    
    Call Objeto.Clear
    Call Matriz.Clear
    
    With UserList(UserIndex)
    
        For i = 1 To MAX_BANCOINVENTORY_SLOTS
            Objeto.Item("user_id") = .Id
            Objeto.Item("number") = i
            Objeto.Item("item_id") = .BancoInvent.Object(i).ObjIndex
            Objeto.Item("amount") = .BancoInvent.Object(i).Amount
                
            ' Lo meto en el array de items
            Matriz.Push Objeto
                
            ' Limpio el objeto para la proxima iteracion
            Objeto.Clear
        Next i
    
    End With

    Set InventarioBanco = Matriz

End Function

Function Habilidades(ByRef UserIndex As Integer) As JS_Array
    
    Call Objeto.Clear
    Call Matriz.Clear
    
    With UserList(UserIndex)
    
        For i = 1 To NUMSKILLS
            Objeto.Item("user_id") = .Id
            Objeto.Item("number") = i
            Objeto.Item("value") = .Stats.UserSkills(i)
                
            ' Lo meto en el array de items
            Matriz.Push Objeto
                
            ' Limpio el objeto para la proxima iteracion
            Objeto.Clear
        Next i
    
    End With

    Set Habilidades = Matriz

End Function

Function Mascotas(ByRef UserIndex As Integer) As JS_Array
    
    Call Objeto.Clear
    Call Matriz.Clear
    
    Dim petType As Integer
    
    With UserList(UserIndex)

        For i = 1 To MAXMASCOTAS
            Objeto.Item("user_id") = .Id
            Objeto.Item("number") = i
                
            'CHOTS | I got this logic from SaveUserToCharfile
            If .MascotasIndex(i) > 0 Then
            
                If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then
                    petType = .MascotasType(i)
                Else
                    petType = 0
                End If

            Else
                petType = .MascotasType(i)

            End If
                
            Objeto.Item("pet_id") = petType
                
            ' Lo meto en el array de items
            Matriz.Push Objeto
                
            ' Limpio el objeto para la proxima iteracion
            Objeto.Clear
        Next i
    
    End With

    Set Mascotas = Matriz
    
End Function

Function Conexiones(ByRef UserIndex As Integer) As JS_Object
    
    Call Objeto.Clear
    
    With UserList(UserIndex)

        Objeto.Item("user_id") = .Id
        Objeto.Item("ip") = .ip
        Objeto.Item("date_last_login") = CStr(Now())
    
    End With

    Set Conexiones = Objeto

End Function

Function Quest(ByRef UserIndex As Integer) As JS_Array
    
    Call Objeto.Clear
    Call Matriz.Clear
    
    Dim LoopK As Long, tempCount As Integer, Tmp As String, tempString As String
    
    With UserList(UserIndex)
    
        For i = 1 To MAXUSERQUESTS
            Objeto.Item("user_id") = .Id
            Objeto.Item("number") = i
            Objeto.Item("quest_id") = .QuestStats.Quests(i).QuestIndex
            
            If .QuestStats.Quests(i).QuestIndex > 0 Then
            
                Tmp = QuestList(.QuestStats.Quests(i).QuestIndex).RequiredNPCs
                tempString = vbNullString
                
                If Tmp Then

                    For LoopK = 1 To Tmp
                        tempString = tempString & CStr(.QuestStats.Quests(i).NPCsKilled(LoopK))
                        
                        If LoopK < Tmp Then
                            tempString = tempString & "-"
                        End If

                    Next LoopK
                    
                End If
                
                Objeto.Item("npcs") = tempString
                
                Tmp = QuestList(.QuestStats.Quests(i).QuestIndex).RequiredTargetNPCs
                tempString = vbNullString
                
                If Tmp Then
                
                    For LoopK = 1 To Tmp
                        tempString = tempString & CStr(.QuestStats.Quests(i).NPCsTarget(LoopK))
                        
                        If LoopK < Tmp Then
                            tempString = tempString & "-"
                        End If
                    
                    Next LoopK

                End If
                
                Objeto.Item("npcstarget") = .Stats.UserSkills(i)
                
            End If

            ' Lo meto en el array de items
            Matriz.Push Objeto
                
            ' Limpio el objeto para la proxima iteracion
            Objeto.Clear
        Next i
    
    End With

    Set Quest = Matriz

End Function

Function QuestTerminadas(ByRef UserIndex As Integer) As JS_Array

    With UserList(UserIndex)

        Call Objeto.Clear
        Call Matriz.Clear
                
        For i = 1 To .QuestStats.NumQuestsDone
            Objeto.Item("user_id") = .Id
            Objeto.Item("quest_id") = .QuestStats.QuestsDone(i)
                        
            ' Lo meto en el array de items
            Matriz.Push Objeto
                        
            ' Limpio el objeto para la proxima iteracion
            Objeto.Clear
        Next i
        
    End With

    Set QuestTerminadas = Matriz

End Function
