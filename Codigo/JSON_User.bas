Attribute VB_Name = "JSON_User"
Option Explicit

Private Objeto As New JS_Object
Private Matriz As New JS_Array

Private i As Long

Function Principal(ByRef UserIndex As Integer, Optional ByVal Logout As Boolean = False) As JS_Object
        
        On Error GoTo Principal_Err
    
100     Call Objeto.Clear
    
102     With UserList(UserIndex)
            If .ID >= 0 Then
104             Objeto.Item("id") = .ID
            End If

106         Objeto.Item("name") = .Name
108         Objeto.Item("level") = .Stats.ELV
110         Objeto.Item("exp") = .Stats.Exp
114         Objeto.Item("genre_id") = .genero
116         Objeto.Item("race_id") = .raza
118         Objeto.Item("class_id") = .clase
120         Objeto.Item("home_id") = .Hogar
122         Objeto.Item("description") = .Desc
124         Objeto.Item("gold") = .Stats.GLD
126         Objeto.Item("bank_gold") = .Stats.Banco
128         Objeto.Item("free_skillpoints") = .Stats.SkillPts
130         Objeto.Item("pets_saved") = .flags.MascotasGuardadas
132         Objeto.Item("pos_map") = .Pos.Map
134         Objeto.Item("pos_x") = .Pos.X
136         Objeto.Item("pos_y") = .Pos.Y
140         Objeto.Item("message_info") = .MENSAJEINFORMACION
142         Objeto.Item("body_id") = .Char.Body
144         Objeto.Item("head_id") = .Char.Head
146         Objeto.Item("weapon_id") = .Char.WeaponAnim
148         Objeto.Item("helmet_id") = .Char.CascoAnim
150         Objeto.Item("shield_id") = .Char.ShieldAnim
152         Objeto.Item("heading") = .Char.Heading
156         Objeto.Item("slot_armour") = .Invent.ArmourEqpSlot
158         Objeto.Item("slot_weapon") = .Invent.WeaponEqpSlot
160         Objeto.Item("slot_shield") = .Invent.EscudoEqpSlot
162         Objeto.Item("slot_helmet") = .Invent.CascoEqpSlot
164         Objeto.Item("slot_ammo") = .Invent.MunicionEqpSlot
166         Objeto.Item("slot_tool") = .Invent.HerramientaEqpSlot
168         Objeto.Item("slot_magic") = .Invent.MagicoSlot
170         Objeto.Item("slot_knuckles") = .Invent.NudilloSlot
172         Objeto.Item("slot_ship") = .Invent.BarcoSlot
174         Objeto.Item("slot_mount") = .Invent.MonturaSlot
176         Objeto.Item("min_hp") = .Stats.MinHp
178         Objeto.Item("max_hp") = .Stats.MaxHp
180         Objeto.Item("min_man") = .Stats.MinMAN
182         Objeto.Item("max_man") = .Stats.MaxMAN
184         Objeto.Item("min_sta") = .Stats.MinSta
186         Objeto.Item("max_sta") = .Stats.MaxSta
188         Objeto.Item("min_ham") = .Stats.MinHam
190         Objeto.Item("max_ham") = .Stats.MaxHam
192         Objeto.Item("min_sed") = .Stats.MinAGU
194         Objeto.Item("max_sed") = .Stats.MaxAGU
196         Objeto.Item("min_hit") = .Stats.MinHIT
198         Objeto.Item("max_hit") = .Stats.MaxHit
200         Objeto.Item("killed_npcs") = .Stats.NPCsMuertos
202         Objeto.Item("killed_users") = .Stats.UsuariosMatados
204         Objeto.Item("invent_level") = .Stats.InventLevel
206         Objeto.Item("is_naked") = (.flags.Desnudo = 1)
208         Objeto.Item("is_poisoned") = (.flags.Envenenado = 1)
            Objeto.Item("is_banned") = (.flags.Ban = 1)
218         Objeto.Item("is_sailing") = (.flags.Navegando = 1)
220         Objeto.Item("is_paralyzed") = (.flags.Paralizado = 1)
222         Objeto.Item("is_mounted") = (.flags.Montado = 1)
224         Objeto.Item("is_silenced") = (.flags.Silenciado = 1)
            Objeto.Item("is_logged") = Not Logout
226         Objeto.Item("silence_minutes_left") = .flags.MinutosRestantes
228         Objeto.Item("silence_elapsed_seconds") = .flags.SegundosPasados
232         Objeto.Item("counter_pena") = .Counters.Pena
234         Objeto.Item("deaths") = .flags.VecesQueMoriste
240         Objeto.Item("pertenece_real") = (.Faccion.ArmadaReal = 1)
242         Objeto.Item("pertenece_caos") = (.Faccion.FuerzasCaos = 1)
244         Objeto.Item("ciudadanos_matados") = .Faccion.ciudadanosMatados
246         Objeto.Item("criminales_matados") = .Faccion.CriminalesMatados
248         Objeto.Item("recibio_armadura_real") = (.Faccion.RecibioArmaduraReal = 1)
250         Objeto.Item("recibio_armadura_caos") = (.Faccion.RecibioArmaduraCaos = 1)
252         Objeto.Item("recibio_exp_real") = (.Faccion.RecibioExpInicialReal = 1)
254         Objeto.Item("recibio_exp_caos") = (.Faccion.RecibioExpInicialCaos = 1)
256         Objeto.Item("recompensas_real") = .Faccion.RecompensasReal
258         Objeto.Item("recompensas_caos") = .Faccion.RecompensasCaos
260         Objeto.Item("reenlistadas") = .Faccion.Reenlistadas
264         Objeto.Item("nivel_ingreso") = .Faccion.NivelIngreso
266         Objeto.Item("matados_ingreso") = .Faccion.MatadosIngreso
268         Objeto.Item("siguiente_recompensa") = .Faccion.NextRecompensa
270         Objeto.Item("status") = .Faccion.Status
274         Objeto.Item("guild_index") = .GuildIndex
276         Objeto.Item("chat_combate") = (.ChatCombate = 1)
278         Objeto.Item("chat_global") = (.ChatGlobal = 1)
282         Objeto.Item("warnings") = .Stats.Advertencias
            Objeto.Item("elo") = .Stats.ELO
            Objeto.Item("return_map") = .flags.ReturnPos.Map
            Objeto.Item("return_x") = .flags.ReturnPos.X
            Objeto.Item("return_y") = .flags.ReturnPos.Y
    
        End With
    
284     Set Principal = Objeto

        
        Exit Function

Principal_Err:
        Call RegistrarError(Err.Number, Err.Description, "API_User.Principal", Erl)
        Resume Next
        
End Function

Function Atributos(ByRef UserIndex As Integer) As JS_Object
        
        On Error GoTo Atributos_Err
        
100     Call Objeto.Clear
    
104     With UserList(UserIndex)
            
            If .ID >= 0 Then
                Objeto.Item("user_id") = .ID
            End If

            Objeto.Item("strength") = .Stats.UserAtributos(e_Atributos.Fuerza)
            Objeto.Item("agility") = .Stats.UserAtributos(e_Atributos.Agilidad)
            Objeto.Item("intelligence") = .Stats.UserAtributos(e_Atributos.Inteligencia)
            Objeto.Item("constitution") = .Stats.UserAtributos(e_Atributos.Constitucion)
            Objeto.Item("charisma") = .Stats.UserAtributos(e_Atributos.Carisma)
            
        End With

120     Set Atributos = Objeto

        
        Exit Function

Atributos_Err:
        Call RegistrarError(Err.Number, Err.Description, "API_User.Atributos", Erl)
        Resume Next
        
End Function

Function Hechizo(ByRef UserIndex As Integer) As JS_Array

        On Error GoTo Hechizo_Err
    
100     Call Objeto.Clear
102     Call Matriz.Clear
    
104     With UserList(UserIndex)
    
106         For i = 1 To MAXUSERHECHIZOS
                If .ID >= 0 Then
                    Objeto.Item("user_id") = .ID
                End If

110             Objeto.Item("number") = i
112             Objeto.Item("spell_id") = .Stats.UserHechizos(i)
                    
                ' Lo meto en el array de items
114             Matriz.Push Objeto
                    
                ' Limpio el objeto para la proxima iteracion
116             Objeto.Clear
118         Next i
    
        End With

120     Set Hechizo = Matriz
        
        Exit Function

Hechizo_Err:
122     Call RegistrarError(Err.Number, Err.Description, "API_User.Hechizo", Erl)

124     Resume Next

End Function

Function Inventario(ByRef UserIndex As Integer) As JS_Array
        
        On Error GoTo Inventario_Err
        
    
100     Call Objeto.Clear
102     Call Matriz.Clear
    
104     With UserList(UserIndex)
    
106         For i = 1 To .CurrentInventorySlots
                If .ID >= 0 Then
                    Objeto.Item("user_id") = .ID
                End If

110             Objeto.Item("number") = i
112             Objeto.Item("item_id") = .Invent.Object(i).ObjIndex
114             Objeto.Item("amount") = .Invent.Object(i).amount
116             Objeto.Item("is_equipped") = (.Invent.Object(i).Equipped = 1)
                
                ' Lo meto en el array de items
118             Matriz.Push Objeto
                
                ' Limpio el objeto para la proxima iteracion
120             Objeto.Clear
122         Next i
    
        End With

124     Set Inventario = Matriz

        
        Exit Function

Inventario_Err:
        Call RegistrarError(Err.Number, Err.Description, "API_User.Inventario", Erl)
        Resume Next
        
End Function

Function InventarioBanco(ByRef UserIndex As Integer) As JS_Array
        
        On Error GoTo InventarioBanco_Err
        
    
100     Call Objeto.Clear
102     Call Matriz.Clear
    
104     With UserList(UserIndex)
    
106         For i = 1 To MAX_BANCOINVENTORY_SLOTS
                If .ID >= 0 Then
                    Objeto.Item("user_id") = .ID
                End If

110             Objeto.Item("number") = i
112             Objeto.Item("item_id") = .BancoInvent.Object(i).ObjIndex
114             Objeto.Item("amount") = .BancoInvent.Object(i).amount
                
                ' Lo meto en el array de items
116             Matriz.Push Objeto
                
                ' Limpio el objeto para la proxima iteracion
118             Objeto.Clear
120         Next i
    
        End With

122     Set InventarioBanco = Matriz

        
        Exit Function

InventarioBanco_Err:
        Call RegistrarError(Err.Number, Err.Description, "API_User.InventarioBanco", Erl)
        Resume Next
        
End Function

Function Habilidades(ByRef UserIndex As Integer) As JS_Object
        
    On Error GoTo Habilidades_Err
    
    Call Objeto.Clear
    
    With UserList(UserIndex)
    
        If .ID >= 0 Then
            Objeto.Item("user_id") = .ID
        End If

        Objeto.Item("magia") = .Stats.UserSkills(e_Skill.Magia)
        Objeto.Item("robar") = .Stats.UserSkills(e_Skill.Robar)
        Objeto.Item("tacticas") = .Stats.UserSkills(e_Skill.Tacticas)
        Objeto.Item("armas") = .Stats.UserSkills(e_Skill.Armas)
        Objeto.Item("meditar") = .Stats.UserSkills(e_Skill.Meditar)
        Objeto.Item("apunalar") = .Stats.UserSkills(e_Skill.Apuñalar)
        Objeto.Item("ocultarse") = .Stats.UserSkills(e_Skill.Ocultarse)
        Objeto.Item("supervivencia") = .Stats.UserSkills(e_Skill.Supervivencia)
        Objeto.Item("comerciar") = .Stats.UserSkills(e_Skill.Comerciar)
        Objeto.Item("defensa") = .Stats.UserSkills(e_Skill.Defensa)
        Objeto.Item("liderazgo") = .Stats.UserSkills(e_Skill.liderazgo)
        Objeto.Item("proyectiles") = .Stats.UserSkills(e_Skill.Proyectiles)
        Objeto.Item("wrestling") = .Stats.UserSkills(e_Skill.Wrestling)
        Objeto.Item("navegacion") = .Stats.UserSkills(e_Skill.Navegacion)
        Objeto.Item("equitacion") = .Stats.UserSkills(e_Skill.equitacion)
        Objeto.Item("resistencia") = .Stats.UserSkills(e_Skill.Resistencia)
        Objeto.Item("talar") = .Stats.UserSkills(e_Skill.Talar)
        Objeto.Item("pescar") = .Stats.UserSkills(e_Skill.Pescar)
        Objeto.Item("mineria") = .Stats.UserSkills(e_Skill.Mineria)
        Objeto.Item("herreria") = .Stats.UserSkills(e_Skill.Herreria)
        Objeto.Item("carpinteria") = .Stats.UserSkills(e_Skill.Carpinteria)
        Objeto.Item("alquimia") = .Stats.UserSkills(e_Skill.Alquimia)
        Objeto.Item("sastreria") = .Stats.UserSkills(e_Skill.Sastreria)
        Objeto.Item("domar") = .Stats.UserSkills(e_Skill.Domar)
    
    End With

    Set Habilidades = Objeto
        
    Exit Function

Habilidades_Err:
    Call RegistrarError(Err.Number, Err.Description, "API_User.Habilidades", Erl)
    Resume Next
        
End Function

Function Mascotas(ByRef UserIndex As Integer) As JS_Array
        
        On Error GoTo Mascotas_Err
        
    
100     Call Objeto.Clear
102     Call Matriz.Clear
    
        Dim petType As Integer
    
104     With UserList(UserIndex)

106         For i = 1 To MAXMASCOTAS
                If .ID >= 0 Then
                    Objeto.Item("user_id") = .ID
                End If

110             Objeto.Item("number") = i
                
                'CHOTS | I got this logic from SaveUserToCharfile
112             If .MascotasIndex(i) > 0 Then
            
114                 If NpcList(.MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then
116                     petType = .MascotasType(i)
                    Else
118                     petType = 0
                    End If

                Else
120                 petType = .MascotasType(i)

                End If
                
122             Objeto.Item("pet_id") = petType
                
                ' Lo meto en el array de items
124             Matriz.Push Objeto
                
                ' Limpio el objeto para la proxima iteracion
126             Objeto.Clear
128         Next i
    
        End With

130     Set Mascotas = Matriz
    
        
        Exit Function

Mascotas_Err:
        Call RegistrarError(Err.Number, Err.Description, "API_User.Mascotas", Erl)
        Resume Next
        
End Function

Function Conexiones(ByRef UserIndex As Integer) As JS_Object
        
        On Error GoTo Conexiones_Err
        
    
100     Call Objeto.Clear
    
102     With UserList(UserIndex)
            If .ID >= 0 Then
                Objeto.Item("user_id") = .ID
            End If

106         Objeto.Item("ip") = .IP
108         Objeto.Item("date_last_login") = CStr(Now())
    
        End With

110     Set Conexiones = Objeto

        
        Exit Function

Conexiones_Err:
        Call RegistrarError(Err.Number, Err.Description, "API_User.Conexiones", Erl)
        Resume Next
        
End Function

Function Quest(ByRef UserIndex As Integer) As JS_Array
        
        On Error GoTo Quest_Err
    
100     Call Objeto.Clear
102     Call Matriz.Clear
    
        Dim LoopK As Long, tempCount As Integer, Tmp As String, tempString As String
    
104     With UserList(UserIndex)
    
106         For i = 1 To MAXUSERQUESTS
                
108             With .QuestStats.Quests(i)
                    If UserList(UserIndex).ID >= 0 Then
                        Objeto.Item("user_id") = UserList(UserIndex).ID
                    End If

112                 Objeto.Item("number") = i
114                 Objeto.Item("quest_id") = .QuestIndex
            
116                 If .QuestIndex > 0 Then
            
118                     Tmp = QuestList(.QuestIndex).RequiredNPCs
120                     tempString = vbNullString
                
122                     If Tmp Then

124                         For LoopK = 1 To Tmp
126                             tempString = tempString & CStr(.NPCsKilled(LoopK))
                        
128                             If LoopK < Tmp Then
130                                 tempString = tempString & "-"
                                End If

132                         Next LoopK
                    
                        End If
                
134                     Objeto.Item("npcs") = tempString
                
136                     Tmp = QuestList(.QuestIndex).RequiredTargetNPCs
138                     tempString = vbNullString
                
140                     If Tmp Then
                
142                         For LoopK = 1 To Tmp
144                             tempString = tempString & CStr(.NPCsTarget(LoopK))
                        
146                             If LoopK < Tmp Then
148                                 tempString = tempString & "-"
                                End If
                    
150                         Next LoopK

                        End If
                
152                     Objeto.Item("npcstarget") = UserList(UserIndex).Stats.UserSkills(i)
                
                    End If

                    ' Lo meto en el array de items
154                 Matriz.Push Objeto
                
                    ' Limpio el objeto para la proxima iteracion
156                 Objeto.Clear
                
                End With
            
158         Next i
    
        End With

160     Set Quest = Matriz
        
        Exit Function

Quest_Err:
162     Call RegistrarError(Err.Number, Err.Description, "API_User.Quest", Erl)
164     Resume Next
        
End Function

Function QuestTerminadas(ByRef UserIndex As Integer) As JS_Array
        
        On Error GoTo QuestTerminadas_Err
        

100     With UserList(UserIndex)

102         Call Objeto.Clear
104         Call Matriz.Clear
                
106         For i = 1 To .QuestStats.NumQuestsDone
                If .ID >= 0 Then
108                 Objeto.Item("user_id") = .ID
                End If
110             Objeto.Item("quest_id") = .QuestStats.QuestsDone(i)
                        
                ' Lo meto en el array de items
112             Matriz.Push Objeto
                        
                ' Limpio el objeto para la proxima iteracion
114             Objeto.Clear
116         Next i
        
        End With

118     Set QuestTerminadas = Matriz

        
        Exit Function

QuestTerminadas_Err:
        Call RegistrarError(Err.Number, Err.Description, "API_User.QuestTerminadas", Erl)
        Resume Next
        
End Function
