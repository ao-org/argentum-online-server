Attribute VB_Name = "API_User"
Option Explicit

Private Objeto As New JS_Object
Private Matriz As New JS_Array

Private i As Long

Function Principal(ByRef UserIndex As Integer, ByRef Logout As Boolean) As JS_Object
        
        On Error GoTo Principal_Err
        
    
100     Call Objeto.Clear
    
102     With UserList(UserIndex)
    
104         Objeto.Item("id") = .Id
106         Objeto.Item("name") = .name
108         Objeto.Item("level") = .Stats.ELV
110         Objeto.Item("exp") = CLng(.Stats.Exp)
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
138         Objeto.Item("last_map") = .flags.lastMap
140         Objeto.Item("message_info") = .MENSAJEINFORMACION
142         Objeto.Item("body_id") = .Char.Body
144         Objeto.Item("head_id") = .OrigChar.Head
146         Objeto.Item("weapon_id") = .Char.WeaponAnim
148         Objeto.Item("helmet_id") = .Char.CascoAnim
150         Objeto.Item("shield_id") = .Char.ShieldAnim
152         Objeto.Item("heading") = .Char.Heading
154         Objeto.Item("items_Amount") = .Invent.NroItems
156         Objeto.Item("slot_armour") = .Invent.ArmourEqpSlot
158         Objeto.Item("slot_weapon") = .Invent.WeaponEqpSlot
160         Objeto.Item("slot_shield") = .Invent.EscudoEqpSlot
162         Objeto.Item("slot_helmet") = .Invent.CascoEqpSlot
164         Objeto.Item("slot_ammo") = .Invent.MunicionEqpSlot
            'Objeto.Item("slot_ring") = .Invent.AnilloEqpSlot
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
            'Objeto.Item( "rep_asesino") =.Reputacion.AsesinoRep
            'Objeto.Item( "rep_bandido") =.Reputacion.BandidoRep
            'Objeto.Item( "rep_burgues") =.Reputacion.BurguesRep
            'Objeto.Item( "rep_ladron") =.Reputacion.LadronesRep
            'Objeto.Item( "rep_noble") =.Reputacion.NobleRep
            'Objeto.Item( "rep_plebe") =.Reputacion.PlebeRep
            'Objeto.Item( "rep_average") =.Reputacion.Promedio
206         Objeto.Item("is_naked") = .flags.Desnudo
208         Objeto.Item("is_poisoned") = .flags.Envenenado
210         Objeto.Item("is_hidden") = .flags.Escondido
212         Objeto.Item("is_hungry") = .flags.Hambre
214         Objeto.Item("is_thirsty") = .flags.Sed
            'Objeto.Item( "is_banned") =.flags.Ban & ") =" Esto es innecesario porque se setea cuando lo baneas
216         Objeto.Item("is_dead") = .flags.Muerto
218         Objeto.Item("is_sailing") = .flags.Navegando
220         Objeto.Item("is_paralyzed") = .flags.Paralizado
222         Objeto.Item("is_mounted") = .flags.Montado
224         Objeto.Item("is_silenced") = .flags.Silenciado
226         Objeto.Item("silence_minutes_left") = .flags.MinutosRestantes
228         Objeto.Item("silence_elapsed_seconds") = .flags.SegundosPasados
230         Objeto.Item("spouse") = .flags.Pareja
232         Objeto.Item("counter_pena") = .Counters.Pena
234         Objeto.Item("deaths") = .flags.VecesQueMoriste
236         Objeto.Item("pertenece_consejo_real") = (.flags.Privilegios And PlayerType.RoyalCouncil)
238         Objeto.Item("pertenece_consejo_caos") = (.flags.Privilegios And PlayerType.ChaosCouncil)
240         Objeto.Item("pertenece_real") = .Faccion.ArmadaReal
242         Objeto.Item("pertenece_caos") = .Faccion.FuerzasCaos
244         Objeto.Item("ciudadanos_matados") = .Faccion.ciudadanosMatados
246         Objeto.Item("criminales_matados") = .Faccion.CriminalesMatados
248         Objeto.Item("recibio_armadura_real") = .Faccion.RecibioArmaduraReal
250         Objeto.Item("recibio_armadura_caos") = .Faccion.RecibioArmaduraCaos
252         Objeto.Item("recibio_exp_real") = .Faccion.RecibioExpInicialReal
254         Objeto.Item("recibio_exp_caos") = .Faccion.RecibioExpInicialCaos
256         Objeto.Item("recompensas_real") = .Faccion.RecompensasReal
258         Objeto.Item("recompensas_caos") = .Faccion.RecompensasCaos
260         Objeto.Item("reenlistadas") = .Faccion.Reenlistadas
262         Objeto.Item("fecha_ingreso") = .Faccion.FechaIngreso
264         Objeto.Item("nivel_ingreso") = .Faccion.NivelIngreso
266         Objeto.Item("matados_ingreso") = .Faccion.MatadosIngreso
268         Objeto.Item("siguiente_recompensa") = .Faccion.NextRecompensa
270         Objeto.Item("status") = .Faccion.Status
274         Objeto.Item("guild_index") = .GuildIndex
276         Objeto.Item("chat_combate") = .ChatCombate
278         Objeto.Item("chat_global") = .ChatGlobal
280         Objeto.Item("is_logged") = Not Logout
282         Objeto.Item("warnings") = .Stats.Advertencias
    
        End With
    
284     Set Principal = Objeto

        
        Exit Function

Principal_Err:
        Call RegistrarError(Err.Number, Err.Description, "API_User.Principal", Erl)
        Resume Next
        
End Function

Function Atributos(ByRef UserIndex As Integer) As JS_Array
        
        On Error GoTo Atributos_Err
        
    
100     Call Objeto.Clear
102     Call Matriz.Clear
    
104     With UserList(UserIndex)
    
106         For i = 1 To NUMATRIBUTOS
108             Objeto.Item("user_id") = .Id
110             Objeto.Item("number") = i
112             Objeto.Item("value") = .Stats.UserAtributosBackUP(i)
    
                ' Lo meto en el array de items
114             Matriz.Push Objeto
                    
                ' Limpio el objeto para la proxima iteracion
116             Objeto.Clear
118         Next i
    
        End With

120     Set Atributos = Matriz

        
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
108             Objeto.Item("user_id") = .Id
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
108             Objeto.Item("user_id") = .Id
110             Objeto.Item("number") = i
112             Objeto.Item("item_id") = .Invent.Object(i).ObjIndex
114             Objeto.Item("Amount") = .Invent.Object(i).Amount
116             Objeto.Item("is_equipped") = .Invent.Object(i).Equipped
                
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
108             Objeto.Item("user_id") = .Id
110             Objeto.Item("number") = i
112             Objeto.Item("item_id") = .BancoInvent.Object(i).ObjIndex
114             Objeto.Item("amount") = .BancoInvent.Object(i).Amount
                
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

Function Habilidades(ByRef UserIndex As Integer) As JS_Array
        
        On Error GoTo Habilidades_Err
        
    
100     Call Objeto.Clear
102     Call Matriz.Clear
    
104     With UserList(UserIndex)
    
106         For i = 1 To NUMSKILLS
108             Objeto.Item("user_id") = .Id
110             Objeto.Item("number") = i
112             Objeto.Item("value") = .Stats.UserSkills(i)
                
                ' Lo meto en el array de items
114             Matriz.Push Objeto
                
                ' Limpio el objeto para la proxima iteracion
116             Objeto.Clear
118         Next i
    
        End With

120     Set Habilidades = Matriz

        
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
108             Objeto.Item("user_id") = .Id
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

104         Objeto.Item("user_id") = .Id
106         Objeto.Item("ip") = .ip
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
108             Objeto.Item("user_id") = .Id
110             Objeto.Item("number") = i
112             Objeto.Item("quest_id") = .QuestStats.Quests(i).QuestIndex
            
114             If .QuestStats.Quests(i).QuestIndex > 0 Then
            
116                 Tmp = QuestList(.QuestStats.Quests(i).QuestIndex).RequiredNPCs
118                 tempString = vbNullString
                
120                 If Tmp Then

122                     For LoopK = 1 To Tmp
124                         tempString = tempString & CStr(.QuestStats.Quests(i).NPCsKilled(LoopK))
                        
126                         If LoopK < Tmp Then
128                             tempString = tempString & "-"
                            End If

130                     Next LoopK
                    
                    End If
                
132                 Objeto.Item("npcs") = tempString
                
134                 Tmp = QuestList(.QuestStats.Quests(i).QuestIndex).RequiredTargetNPCs
136                 tempString = vbNullString
                
138                 If Tmp Then
                
140                     For LoopK = 1 To Tmp
142                         tempString = tempString & CStr(.QuestStats.Quests(i).NPCsTarget(LoopK))
                        
144                         If LoopK < Tmp Then
146                             tempString = tempString & "-"
                            End If
                    
148                     Next LoopK

                    End If
                
150                 Objeto.Item("npcstarget") = .Stats.UserSkills(i)
                
                End If

                ' Lo meto en el array de items
152             Matriz.Push Objeto
                
                ' Limpio el objeto para la proxima iteracion
154             Objeto.Clear
156         Next i
    
        End With

158     Set Quest = Matriz

        
        Exit Function

Quest_Err:
        Call RegistrarError(Err.Number, Err.Description, "API_User.Quest", Erl)
        Resume Next
        
End Function

Function QuestTerminadas(ByRef UserIndex As Integer) As JS_Array
        
        On Error GoTo QuestTerminadas_Err
        

100     With UserList(UserIndex)

102         Call Objeto.Clear
104         Call Matriz.Clear
                
106         For i = 1 To .QuestStats.NumQuestsDone
108             Objeto.Item("user_id") = .Id
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
