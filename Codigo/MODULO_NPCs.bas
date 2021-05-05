Attribute VB_Name = "NPCs"
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Public Const MaxRespawn             As Integer = 255

Public RespawnList(1 To MaxRespawn) As npc

Option Explicit

Sub QuitarMascotaNpc(ByVal Maestro As Integer)
        
        On Error GoTo QuitarMascotaNpc_Err
        
100     NpcList(Maestro).Mascotas = NpcList(Maestro).Mascotas - 1

        
        Exit Sub

QuitarMascotaNpc_Err:
102     Call RegistrarError(Err.Number, Err.Description, "NPCs.QuitarMascotaNpc", Erl)
104     Resume Next
        
End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

        '********************************************************
        'Author: Unknown
        'Llamado cuando la vida de un NPC llega a cero.
        'Last Modify Date: 24/01/2007
        '22/06/06: (Nacho) Chequeamos si es pretoriano
        '24/01/2007: Pablo (ToxicWaste): Agrego para actualización de tag si cambia de status.
        '********************************************************
        On Error GoTo ErrHandler
        
        ' Objetivo de pruebas nunca muere
100     If NpcList(NpcIndex).NPCtype = DummyTarget Then
102         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("¡¡Auch!!", NpcList(NpcIndex).Char.CharIndex, vbRed, "Barrin"))
104         NpcList(NpcIndex).Stats.MinHp = NpcList(NpcIndex).Stats.MaxHp
            Exit Sub
        End If

        Dim MiNPC As npc

106     MiNPC = NpcList(NpcIndex)

        Dim EraCriminal As Byte

        Dim TiempoRespw As Long

108     TiempoRespw = NpcList(NpcIndex).Contadores.IntervaloRespawn

        ' Es pretoriano?
110     If MiNPC.NPCtype = eNPCType.Pretoriano Then
112         Call ClanPretoriano(MiNPC.ClanIndex).MuerePretoriano(NpcIndex)
            
        ' Es NPC de la invasión?
114     ElseIf MiNPC.flags.InvasionIndex Then
116         Call MuereNpcInvasion(MiNPC.flags.InvasionIndex, MiNPC.flags.IndexInInvasion)
        End If

        'Quitamos el npc
118     Call QuitarNPC(NpcIndex)
    
120     If UserIndex > 0 Then ' Lo mato un usuario?
122         If MiNPC.flags.Snd3 > 0 Then
124             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.Pos.X, MiNPC.Pos.Y))
            Else
126             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("28", MiNPC.Pos.X, MiNPC.Pos.Y))
            End If

128         UserList(UserIndex).flags.TargetNPC = 0
130         UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        
            'El user que lo mato tiene mascotas?
132         Call AllFollowAmo(UserIndex)
            
134         If UserList(UserIndex).ChatCombate = 1 Then
136             Call WriteLocaleMsg(UserIndex, "184", FontTypeNames.FONTTYPE_FIGHT, "la criatura")
            End If

138         If UserList(UserIndex).Stats.NPCsMuertos < 32000 Then UserList(UserIndex).Stats.NPCsMuertos = UserList(UserIndex).Stats.NPCsMuertos + 1
            
140         If MiNPC.MaestroUser > 0 Then Exit Sub
            
142         Call SubirSkill(UserIndex, eSkill.Supervivencia)

144         If MiNPC.flags.ExpCount > 0 Then

146             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
148                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + MiNPC.flags.ExpCount

150                 If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
                    
152                 Call WriteTextOverTile(UserIndex, "+" & PonerPuntos(MiNPC.flags.ExpCount), UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, RGB(0, 169, 255))
154                 Call WriteUpdateExp(UserIndex)
156                 Call CheckUserLevel(UserIndex)

                End If
        
158             MiNPC.flags.ExpCount = 0

            End If
        
160         EraCriminal = Status(UserIndex)
        
162         If MiNPC.GiveEXPClan > 0 Then
164             If UserList(UserIndex).GuildIndex > 0 Then
166                 Call modGuilds.CheckClanExp(UserIndex, MiNPC.GiveEXPClan)

                    ' Else
                    ' Call WriteConsoleMsg(UserIndex, "No perteneces a ningun clan, experiencia perdida.", FontTypeNames.FONTTYPE_INFOIAO)
                End If

            End If
        
            Dim i As Long, j As Long
        
168         For i = 1 To MAXUSERQUESTS
        
170             With UserList(UserIndex).QuestStats.Quests(i)
        
172                 If .QuestIndex Then
174                     If QuestList(.QuestIndex).RequiredNPCs Then
        
176                         For j = 1 To QuestList(.QuestIndex).RequiredNPCs
        
178                             If QuestList(.QuestIndex).RequiredNPC(j).NpcIndex = MiNPC.Numero Then
180                                 If QuestList(.QuestIndex).RequiredNPC(j).amount > .NPCsKilled(j) Then
182                                     .NPCsKilled(j) = .NPCsKilled(j) + 1
        
                                    End If
                                    
184                                 If QuestList(.QuestIndex).RequiredNPC(j).amount = .NPCsKilled(j) Then
186                                     Call WriteConsoleMsg(UserIndex, "Ya has matado todos los " & MiNPC.name & " que la misión " & QuestList(.QuestIndex).nombre & " requería. Revisa si ya estás listo para recibir la recompensa.", FontTypeNames.FONTTYPE_INFOIAO)
                                    
                                    End If
        
                                End If
        
188                         Next j
        
                        End If
        
                    End If
        
                End With
        
190         Next i

            'Tiramos el oro
192         Call NPCTirarOro(MiNPC, UserIndex)

194         Call DropObjQuest(MiNPC, UserIndex)
    
            'Item Magico!
196         Call NpcDropeo(MiNPC, UserIndex)
            
            'Tiramos el inventario
198         Call NPC_TIRAR_ITEMS(MiNPC)
            
        End If ' UserIndex > 0
        
        ' Mascotas y npcs de entrenamiento no respawnean
200     If MiNPC.MaestroNPC > 0 Or MiNPC.MaestroUser > 0 Then Exit Sub
        
202     If NpcIndex = npc_index_evento Then
204         BusquedaNpcActiva = False
206         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Evento> El NPC ha sido asesinado.", FontTypeNames.FONTTYPE_CITIZEN))

        End If
        
        'ReSpawn o no
208     If TiempoRespw = 0 Then
210         Call ReSpawnNpc(MiNPC)

        Else

            Dim Indice As Integer

212         MiNPC.flags.NPCActive = True
214         Indice = ObtenerIndiceRespawn
216         RespawnList(Indice) = MiNPC

        End If
    
        Exit Sub

ErrHandler:
218     Call RegistrarError(Err.Number, Err.Description, "NPCs.MuereNpc", Erl())

End Sub

Sub ResetNpcFlags(ByVal NpcIndex As Integer)
        'Clear the npc's flags
        
        On Error GoTo ResetNpcFlags_Err
        
    
100     With NpcList(NpcIndex).flags
102         .AfectaParalisis = 0
104         .AguaValida = 0
106         .AttackedBy = vbNullString
108         .AttackedFirstBy = vbNullString
110         .Attacking = 0
112         .backup = 0
114         .Bendicion = 0
116         .Domable = 0
118         .Envenenado = 0
120         .Faccion = 0
122         .Follow = False
124         .LanzaSpells = 0
126         .GolpeExacto = 0
128         .invisible = 0
130         .OldHostil = 0
132         .OldMovement = 0
134         .Paralizado = 0
136         .Inmovilizado = 0
138         .Respawn = 0
140         .RespawnOrigPos = 0
142         .Snd1 = 0
144         .Snd2 = 0
146         .Snd3 = 0
148         .TierraInvalida = 0
150         .AtacaUsuarios = True
152         .AtacaNPCs = True
154         .AIAlineacion = e_Alineacion.ninguna
156         .NPCIdle = False
        End With

        
        Exit Sub

ResetNpcFlags_Err:
158     Call RegistrarError(Err.Number, Err.Description, "NPCs.ResetNpcFlags", Erl)
160     Resume Next
        
End Sub

Sub ResetNpcCounters(ByVal NpcIndex As Integer)
        
        On Error GoTo ResetNpcCounters_Err
        

100     NpcList(NpcIndex).Contadores.Paralisis = 0
102     NpcList(NpcIndex).Contadores.TiempoExistencia = 0
104     NpcList(NpcIndex).Contadores.IntervaloMovimiento = 0
106     NpcList(NpcIndex).Contadores.IntervaloAtaque = 0
108     NpcList(NpcIndex).Contadores.IntervaloLanzarHechizo = 0
110     NpcList(NpcIndex).Contadores.IntervaloRespawn = 0

        
        Exit Sub

ResetNpcCounters_Err:
112     Call RegistrarError(Err.Number, Err.Description, "NPCs.ResetNpcCounters", Erl)
114     Resume Next
        
End Sub

Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)
        
        On Error GoTo ResetNpcCharInfo_Err
        

100     NpcList(NpcIndex).Char.Body = 0
102     NpcList(NpcIndex).Char.CascoAnim = 0
104     NpcList(NpcIndex).Char.CharIndex = 0
106     NpcList(NpcIndex).Char.FX = 0
108     NpcList(NpcIndex).Char.Head = 0
110     NpcList(NpcIndex).Char.Heading = 0
112     NpcList(NpcIndex).Char.loops = 0
114     NpcList(NpcIndex).Char.ShieldAnim = 0
116     NpcList(NpcIndex).Char.WeaponAnim = 0

        
        Exit Sub

ResetNpcCharInfo_Err:
118     Call RegistrarError(Err.Number, Err.Description, "NPCs.ResetNpcCharInfo", Erl)
120     Resume Next
        
End Sub

Sub ResetNpcCriatures(ByVal NpcIndex As Integer)
        
        On Error GoTo ResetNpcCriatures_Err
        

        Dim j As Integer

100     For j = 1 To NpcList(NpcIndex).NroCriaturas
102         NpcList(NpcIndex).Criaturas(j).NpcIndex = 0
104         NpcList(NpcIndex).Criaturas(j).NpcName = vbNullString
106     Next j

108     NpcList(NpcIndex).NroCriaturas = 0

        
        Exit Sub

ResetNpcCriatures_Err:
110     Call RegistrarError(Err.Number, Err.Description, "NPCs.ResetNpcCriatures", Erl)
112     Resume Next
        
End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)
        
        On Error GoTo ResetExpresiones_Err
        

        Dim j As Integer

100     For j = 1 To NpcList(NpcIndex).NroExpresiones
102         NpcList(NpcIndex).Expresiones(j) = vbNullString
104     Next j

106     NpcList(NpcIndex).NroExpresiones = 0

        
        Exit Sub

ResetExpresiones_Err:
108     Call RegistrarError(Err.Number, Err.Description, "NPCs.ResetExpresiones", Erl)
110     Resume Next
        
End Sub

Sub ResetDrop(ByVal NpcIndex As Integer)
        
        On Error GoTo ResetDrop_Err
        

        Dim j As Integer

100     For j = 1 To NpcList(NpcIndex).NumQuiza
102         NpcList(NpcIndex).QuizaDropea(j) = 0
104     Next j

106     NpcList(NpcIndex).NumQuiza = 0

        
        Exit Sub

ResetDrop_Err:
108     Call RegistrarError(Err.Number, Err.Description, "NPCs.ResetDrop", Erl)
110     Resume Next
        
End Sub

Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)
        
        On Error GoTo ResetNpcMainInfo_Err
        
    
100     NpcList(NpcIndex).Attackable = 0
102     NpcList(NpcIndex).Comercia = 0
104     NpcList(NpcIndex).GiveEXP = 0
106     NpcList(NpcIndex).GiveEXPClan = 0
108     NpcList(NpcIndex).GiveGLD = 0
110     NpcList(NpcIndex).Hostile = 0
112     NpcList(NpcIndex).InvReSpawn = 0
114     NpcList(NpcIndex).level = 0

    
116     NpcList(NpcIndex).MaestroNPC = 0
    
118     NpcList(NpcIndex).Mascotas = 0
120     NpcList(NpcIndex).Movement = 0
122     NpcList(NpcIndex).name = "NPC SIN INICIAR"
124     NpcList(NpcIndex).NPCtype = 0
126     NpcList(NpcIndex).Numero = 0
128     NpcList(NpcIndex).Orig.Map = 0
130     NpcList(NpcIndex).Orig.X = 0
132     NpcList(NpcIndex).Orig.Y = 0
134     NpcList(NpcIndex).PoderAtaque = 0
136     NpcList(NpcIndex).PoderEvasion = 0
138     NpcList(NpcIndex).Pos.Map = 0
140     NpcList(NpcIndex).Pos.X = 0
142     NpcList(NpcIndex).Pos.Y = 0
144     NpcList(NpcIndex).Target = 0
146     NpcList(NpcIndex).TargetNPC = 0
148     NpcList(NpcIndex).TipoItems = 0
150     NpcList(NpcIndex).Veneno = 0
152     NpcList(NpcIndex).Desc = vbNullString
154     NpcList(NpcIndex).NumDropQuest = 0
        
156     If NpcList(NpcIndex).MaestroUser > 0 Then Call QuitarMascota(NpcList(NpcIndex).MaestroUser, NpcIndex)
158     If NpcList(NpcIndex).MaestroNPC > 0 Then Call QuitarMascotaNpc(NpcList(NpcIndex).MaestroNPC)

160     NpcList(NpcIndex).MaestroUser = 0
162     NpcList(NpcIndex).MaestroNPC = 0

164     NpcList(NpcIndex).CaminataActual = 0
    
        Dim j As Integer

166     For j = 1 To NpcList(NpcIndex).NroSpells
168         NpcList(NpcIndex).Spells(j) = 0
170     Next j
        
172     Call ResetNpcCharInfo(NpcIndex)
174     Call ResetNpcCriatures(NpcIndex)
176     Call ResetExpresiones(NpcIndex)
178     Call ResetDrop(NpcIndex)

        
        Exit Sub

ResetNpcMainInfo_Err:
180     Call RegistrarError(Err.Number, Err.Description, "NPCs.ResetNpcMainInfo", Erl)
182     Resume Next
        
End Sub

Sub QuitarNPC(ByVal NpcIndex As Integer)

        On Error GoTo ErrHandler

100     NpcList(NpcIndex).flags.NPCActive = False
    
102     If InMapBounds(NpcList(NpcIndex).Pos.Map, NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y) Then
104         Call EraseNPCChar(NpcIndex)
        End If
    
        ' Es pretoriano?
106     If NpcList(NpcIndex).NPCtype = eNPCType.Pretoriano Then
108         Call ClanPretoriano(NpcList(NpcIndex).ClanIndex).MuerePretoriano(NpcIndex)
        End If
    
        'Nos aseguramos de que el inventario sea removido...
        'asi los lobos no volveran a tirar armaduras ;))
110     Call ResetNpcInv(NpcIndex)
112     Call ResetNpcFlags(NpcIndex)
114     Call ResetNpcCounters(NpcIndex)
    
116     Call ResetNpcMainInfo(NpcIndex)
    
118     If NpcIndex = LastNPC Then

120         Do Until NpcList(LastNPC).flags.NPCActive
122             LastNPC = LastNPC - 1

124             If LastNPC < 1 Then Exit Do
            Loop

        End If
      
126     If NumNPCs <> 0 Then
128         NumNPCs = NumNPCs - 1

        End If

        Exit Sub

ErrHandler:
130     NpcList(NpcIndex).flags.NPCActive = False
132     Call LogError("Error en QuitarNPC")

End Sub

Function TestSpawnTrigger(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

        On Error GoTo TestSpawnTrigger_Err

100     TestSpawnTrigger = MapData(Map, X, Y).trigger <> 3 And MapData(Map, X, Y).trigger <> 2 And MapData(Map, X, Y).trigger <> 1

        Exit Function

TestSpawnTrigger_Err:
102     Call RegistrarError(Err.Number, Err.Description, "NPCs.TestSpawnTrigger", Erl)
104     Resume Next
        
End Function

Public Function CrearNPC(NroNPC As Integer, Mapa As Integer, OrigPos As WorldPos, Optional ByVal CustomHead As Integer)
        'Call LogTarea("Sub CrearNPC")
        'Crea un NPC del tipo NRONPC
        
        On Error GoTo CrearNPC_Err

        Dim NpcIndex       As Integer
        Dim Iteraciones    As Long

        Dim PuedeAgua      As Boolean
        Dim PuedeTierra    As Boolean

        Dim Map            As Integer
        Dim X              As Integer
        Dim Y              As Integer

100     NpcIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    
102     If NpcIndex = 0 Then Exit Function

104     With NpcList(NpcIndex)
        
            ' Cabeza customizada
106         If CustomHead <> 0 Then .Char.Head = CustomHead
    
108         PuedeAgua = .flags.AguaValida = 1
110         PuedeTierra = .flags.TierraInvalida = 0
        
            'Necesita ser respawned en un lugar especifico
112         If .flags.RespawnOrigPos And InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
            
114             Map = OrigPos.Map
116             X = OrigPos.X
118             Y = OrigPos.Y
120             .Orig = OrigPos
122             .Pos = OrigPos
           
            Else
                ' Primera búsqueda: buscamos una posición ideal hasta llegar al máximo de iteraciones
                Do
124                 .Pos.Map = Mapa
126                 .Pos.X = RandomNumber(MinXBorder + 2, MaxXBorder - 2) 'Obtenemos posicion al azar en x
128                 .Pos.Y = RandomNumber(MinYBorder + 2, MaxYBorder - 2) 'Obtenemos posicion al azar en y
    
130                 .Pos = ClosestLegalPosNPC(NpcIndex, 10)       'Nos devuelve la posicion valida mas cercana
                    
132                 Iteraciones = Iteraciones + 1
    
134             Loop While .Pos.X = 0 And .Pos.Y = 0 And Iteraciones < MAXSPAWNATTEMPS
    
                ' Si no encontramos una posición válida en la primera instancia
136             If Iteraciones >= MAXSPAWNATTEMPS Then
                    ' Hacemos una búsqueda exhaustiva partiendo desde el centro del mapa
138                 .Pos.Map = Mapa
140                 .Pos.X = (XMaxMapSize - XMinMapSize) \ 2
142                 .Pos.Y = (YMaxMapSize - YMinMapSize) \ 2
                    
144                 .Pos = ClosestLegalPosNPC(NpcIndex, (XMaxMapSize - XMinMapSize) \ 2)
                    
                    ' Si sigue fallando
146                 If .Pos.X = 0 And .Pos.Y = 0 Then
                        ' Hacemos una última búsqueda exhaustiva, ignorando los usuarios
148                     .Pos.Map = Mapa
150                     .Pos.X = (XMaxMapSize - XMinMapSize) \ 2
152                     .Pos.Y = (YMaxMapSize - YMinMapSize) \ 2
                        
154                     .Pos = ClosestLegalPosNPC(NpcIndex, (XMaxMapSize - XMinMapSize) \ 2, True)
                        
                        ' Si falló, borramos el NPC y salimos
156                     If .Pos.X = 0 And .Pos.Y = 0 Then
158                         Call QuitarNPC(NpcIndex)
                            Exit Function
                        End If
                    End If
                End If
            
                'asignamos las nuevas coordenas
160             Map = .Pos.Map
162             X = .Pos.X
164             Y = .Pos.Y
    
            End If

        End With
    
        'Crea el NPC
166     Call MakeNPCChar(True, Map, NpcIndex, Map, X, Y)
        
168     CrearNPC = NpcIndex
        
        Exit Function

CrearNPC_Err:
170     Call RegistrarError(Err.Number, Err.Description, "NPCs.CrearNPC", Erl)

End Function

Sub MakeNPCChar(ByVal toMap As Boolean, sndIndex As Integer, NpcIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo MakeNPCChar_Err
        
100     With NpcList(NpcIndex)

            Dim CharIndex As Integer
    
102         If .Char.CharIndex = 0 Then
104             CharIndex = NextOpenCharIndex
106             .Char.CharIndex = CharIndex
108             CharList(CharIndex) = NpcIndex
    
            End If
        
110         MapData(Map, X, Y).NpcIndex = NpcIndex
        
            Dim Simbolo As Byte
        
            Dim GG      As String
            Dim tmpByte As Byte
       
112         GG = IIf(.showName > 0, .name & .SubName, vbNullString)
        
114         If Not toMap Then
116             If .NumQuest > 0 Then
    
                    Dim q As Byte
                    Dim HayFinalizada As Boolean
                    Dim HayDisponible As Boolean
                    Dim HayPendiente As Boolean
            
118                 For q = 1 To .NumQuest
120                      tmpByte = TieneQuest(sndIndex, .QuestNumber(q))
                    
122                      If tmpByte Then
124                         If FinishQuestCheck(sndIndex, .QuestNumber(q), tmpByte) Then
126                             Simbolo = 3
128                             HayFinalizada = True
                            Else
130                             HayPendiente = True
132                             Simbolo = 4
                            End If
                        Else
134                                 If UserDoneQuest(sndIndex, .QuestNumber(q)) Or Not UserDoneQuest(sndIndex, QuestList(.QuestNumber(q)).RequiredQuest) Or UserList(sndIndex).Stats.ELV < QuestList(.QuestNumber(q)).RequiredLevel Then
136                                     Simbolo = 2
                            Else
138                                     Simbolo = 1
140                             HayDisponible = True
                            End If
        
                        End If
        
142                 Next q
                    
                    
                    'Para darle prioridad a ciertos simbolos
                    
144                 If HayDisponible Then
146                     Simbolo = 1
                    End If
                    
148                 If HayPendiente Then
150                     Simbolo = 4
                    End If
                    
152                 If HayFinalizada Then
154                     Simbolo = 3
                    End If
                    'Para darle prioridad a ciertos simbolos
                    
                End If
                
                
156             If UserList(sndIndex).Stats.UserSkills(eSkill.Supervivencia) >= 90 Then
158                 Call WriteCharacterCreate(sndIndex, IIf(.flags.NPCIdle, .Char.BodyIdle, .Char.Body), .Char.Head, .Char.Heading, .Char.CharIndex, X, Y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, GG, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, .Char.speeding, IIf(.MaestroUser = sndIndex, 2, 1), False, 0, 0, 0, 0, .Stats.MinHp, .Stats.MaxHp, Simbolo, .flags.NPCIdle)
                Else
160                 Call WriteCharacterCreate(sndIndex, IIf(.flags.NPCIdle, .Char.BodyIdle, .Char.Body), .Char.Head, .Char.Heading, .Char.CharIndex, X, Y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, GG, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, .Char.speeding, IIf(.MaestroUser = sndIndex, 2, 1), False, 0, 0, 0, 0, 0, 0, Simbolo, .flags.NPCIdle)
                
                End If
            Else
162             Call AgregarNpc(NpcIndex)
    
            End If

        End With
        
        Exit Sub

MakeNPCChar_Err:
164     Call RegistrarError(Err.Number, Err.Description, "NPCs.MakeNPCChar", Erl)
166     Resume Next
        
End Sub

Sub ChangeNPCChar(ByVal NpcIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading)
        
        On Error GoTo ChangeNPCChar_Err
        
100     With NpcList(NpcIndex)

102         If NpcIndex > 0 Then
104             If .flags.NPCIdle Then
106                 Body = .Char.BodyIdle
                End If
    
108             .Char.Head = Head
110             .Char.Heading = Heading
            
112             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(Body, Head, Heading, .Char.CharIndex, 0, 0, 0, 0, 0, .flags.NPCIdle, False))
    
            End If
        
        End With
        
        Exit Sub

ChangeNPCChar_Err:
114     Call RegistrarError(Err.Number, Err.Description, "NPCs.ChangeNPCChar", Erl)
116     Resume Next
        
End Sub

Sub EraseNPCChar(ByVal NpcIndex As Integer)
        
        On Error GoTo EraseNPCChar_Err
        

100     If NpcList(NpcIndex).Char.CharIndex <> 0 Then CharList(NpcList(NpcIndex).Char.CharIndex) = 0

102     If NpcList(NpcIndex).Char.CharIndex = LastChar Then

104         Do Until CharList(LastChar) > 0
106             LastChar = LastChar - 1

108             If LastChar <= 1 Then Exit Do
            Loop

        End If

        'Quitamos del mapa
110     MapData(NpcList(NpcIndex).Pos.Map, NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y).NpcIndex = 0

        'Actualizamos los clientes
112     Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterRemove(NpcList(NpcIndex).Char.CharIndex, True))

        'Update la lista npc
114     NpcList(NpcIndex).Char.CharIndex = 0

        'update NumChars
116     NumChars = NumChars - 1

        
        Exit Sub

EraseNPCChar_Err:
118     Call RegistrarError(Err.Number, Err.Description, "NPCs.EraseNPCChar", Erl)
120     Resume Next
        
End Sub

Public Function MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte) As Boolean
        On Error GoTo errh

        Dim nPos      As WorldPos
        Dim UserIndex As Integer
        Dim esGuardia As Boolean
100     With NpcList(NpcIndex)
102         If .flags.Paralizado + .flags.Inmovilizado > 0 Then Exit Function
            
104         nPos = .Pos
106         Call HeadtoPos(nHeading, nPos)
108         esGuardia = .NPCtype = eNPCType.GuardiaReal Or .NPCtype = eNPCType.GuardiasCaos
            ' es una posicion legal
            
110         If LegalWalkNPC(nPos.Map, nPos.X, nPos.Y, nHeading, .flags.AguaValida = 1, .flags.TierraInvalida = 0, .MaestroUser <> 0, , esGuardia) Then
            
112             UserIndex = MapData(.Pos.Map, nPos.X, nPos.Y).UserIndex
    
                ' Si hay un usuario a donde se mueve el npc, entonces esta muerto o es un gm invisible
114             If UserIndex > 0 Then

116                 With UserList(UserIndex)
                
                        ' Actualizamos posicion y mapa
118                     MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
120                     .Pos.X = NpcList(NpcIndex).Pos.X
122                     .Pos.Y = NpcList(NpcIndex).Pos.Y
124                     MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                        
                        ' Avisamos a los usuarios del area, y al propio usuario lo forzamos a moverse
126                     Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, .Pos.X, .Pos.Y))
128                     Call WriteForceCharMove(UserIndex, InvertHeading(nHeading))

                    End With

                End If
                
130             If HayPuerta(nPos.Map, nPos.X, nPos.Y) Then
132                 Call AccionParaPuertaNpc(nPos.Map, nPos.X, nPos.Y, NpcIndex)
134             ElseIf HayPuerta(nPos.Map, nPos.X + 1, nPos.Y) Then
136                 Call AccionParaPuertaNpc(nPos.Map, nPos.X + 1, nPos.Y, NpcIndex)
138             ElseIf HayPuerta(nPos.Map, nPos.X + 1, nPos.Y - 1) Then
140                 Call AccionParaPuertaNpc(nPos.Map, nPos.X + 1, nPos.Y - 1, NpcIndex)
142             ElseIf HayPuerta(nPos.Map, nPos.X, nPos.Y - 1) Then
144                 Call AccionParaPuertaNpc(nPos.Map, nPos.X, nPos.Y - 1, NpcIndex)
                End If
146             Call AnimacionIdle(NpcIndex, False)
            
148             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))

                'Update map and user pos
150             MapData(.Pos.Map, .Pos.X, .Pos.Y).NpcIndex = 0
152             .Pos = nPos
154             .Char.Heading = nHeading
156             MapData(.Pos.Map, nPos.X, nPos.Y).NpcIndex = NpcIndex
            
158             Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
                ' Npc has moved
160             MoveNPCChar = True

            End If

        End With
    
        Exit Function

errh:
162     LogError ("Error en move npc " & NpcIndex & ". Error: " & Err.Number & " - " & Err.Description)

End Function

Function NextOpenNPC() As Integer
        'Call LogTarea("Sub NextOpenNPC")

        On Error GoTo ErrHandler

        Dim LoopC As Integer
  
100     For LoopC = 1 To MaxNPCs + 1

102         If LoopC > MaxNPCs Then Exit For

104         If Not NpcList(LoopC).flags.NPCActive Then Exit For

106     Next LoopC
  
108     NextOpenNPC = LoopC

        Exit Function
ErrHandler:
110     Call LogError("Error en NextOpenNPC")

End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer, ByVal VenenoNivel As Byte)
        
        On Error GoTo NpcEnvenenarUser_Err
        

        Dim n As Integer

100     n = RandomNumber(1, 100)

102     If n < 30 Then
104         UserList(UserIndex).flags.Envenenado = VenenoNivel

            'Call WriteConsoleMsg(UserIndex, "¡¡La criatura te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
106         If UserList(UserIndex).ChatCombate = 1 Then
108             Call WriteLocaleMsg(UserIndex, "182", FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If

        
        Exit Sub

NpcEnvenenarUser_Err:
110     Call RegistrarError(Err.Number, Err.Description, "NPCs.NpcEnvenenarUser", Erl)
112     Resume Next
        
End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean, Optional Avisar As Boolean = False, Optional ByVal MaestroUser As Integer = 0) As Integer
        
        On Error GoTo SpawnNpc_Err
        

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 23/01/2007
        '23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
        '***************************************************
        Dim newpos         As WorldPos

        Dim altpos         As WorldPos

        Dim nIndex         As Integer

        Dim PuedeAgua      As Boolean

        Dim PuedeTierra    As Boolean

        Dim Map            As Integer

        Dim X              As Integer

        Dim Y              As Integer

100     nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

102     If nIndex = 0 Then
104         SpawnNpc = 0
            Exit Function
        End If

106     PuedeAgua = NpcList(nIndex).flags.AguaValida = 1
108     PuedeTierra = NpcList(nIndex).flags.TierraInvalida = 0

110         Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
            'Si X e Y son iguales a 0 significa que no se encontro posicion valida

112     If newpos.X <> 0 And newpos.Y <> 0 Then
            'Asignamos las nuevas coordenas solo si son validas
114         NpcList(nIndex).Pos.Map = newpos.Map
116         NpcList(nIndex).Pos.X = newpos.X
118         NpcList(nIndex).Pos.Y = newpos.Y
            
        Else
120         Call QuitarNPC(nIndex)
122         SpawnNpc = 0
            Exit Function
        End If

        'asignamos las nuevas coordenas
124     Map = newpos.Map
126     X = NpcList(nIndex).Pos.X
128     Y = NpcList(nIndex).Pos.Y

        ' WyroX: Asignamos el dueño
130     NpcList(nIndex).MaestroUser = MaestroUser
        
        ' WyroX: Caminata de NPCs
132     If NpcList(nIndex).Movement = Caminata Or NpcList(nIndex).Movement = GuardiaPersigueNpc Then
134         NpcList(nIndex).Orig = NpcList(nIndex).Pos
        End If

        'Crea el NPC
136     Call MakeNPCChar(True, Map, nIndex, Map, X, Y)

138     If FX Then
140         Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
142         Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(NpcList(nIndex).Char.CharIndex, FXIDs.FXWARP, 0))

        End If

144     If Avisar Then
146         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(NpcList(nIndex).name & " ha aparecido en " & DarNameMapa(Map) & " , todo indica que puede tener una gran recompensa para el que logre sobrevivir a él.", FontTypeNames.FONTTYPE_CITIZEN))
        End If

148     SpawnNpc = nIndex

        
        Exit Function

SpawnNpc_Err:
150     Call RegistrarError(Err.Number, Err.Description, "NPCs.SpawnNpc", Erl)
152     Resume Next
        
End Function

Sub ReSpawnNpc(MiNPC As npc)
        
        On Error GoTo ReSpawnNpc_Err
        

100     If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.Map, MiNPC.Orig)

        
        Exit Sub

ReSpawnNpc_Err:
102     Call RegistrarError(Err.Number, Err.Description, "NPCs.ReSpawnNpc", Erl)
104     Resume Next
        
End Sub

'Devuelve el nro de enemigos que hay en el Mapa Map
Function NPCHostiles(ByVal Map As Integer) As Integer
        
        On Error GoTo NPCHostiles_Err
        

        Dim NpcIndex As Integer

        Dim cont     As Integer

        'Contador
100     cont = 0

102     For NpcIndex = 1 To LastNPC

            '¿esta vivo?
104         If NpcList(NpcIndex).flags.NPCActive And NpcList(NpcIndex).Pos.Map = Map And NpcList(NpcIndex).Hostile = 1 Then
106             cont = cont + 1
           
            End If
    
108     Next NpcIndex

110     NPCHostiles = cont

        
        Exit Function

NPCHostiles_Err:
112     Call RegistrarError(Err.Number, Err.Description, "NPCs.NPCHostiles", Erl)
114     Resume Next
        
End Function

Sub NPCTirarOro(MiNPC As npc, ByVal UserIndex As Integer)
        
            On Error GoTo NPCTirarOro_Err
            
100         If UserIndex = 0 Then Exit Sub
            
102         If MiNPC.GiveGLD > 0 Then

                Dim Oro As Long
104                 Oro = MiNPC.GiveGLD * OroMult * UserList(UserIndex).flags.ScrollOro
        
106             If UserList(UserIndex).Grupo.EnGrupo Then

108                 Select Case UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
                        Case 2: Oro = Oro * 1.2
110                     Case 3: Oro = Oro * 1.4
112                     Case 4: Oro = Oro * 1.6
114                     Case 5: Oro = Oro * 1.8
116                     Case 6: Oro = Oro * 2
                    End Select
                    
                End If

                Dim MiObj As obj
118             MiObj.ObjIndex = iORO

120             While (Oro > 0)
122                 If Oro > MAX_INVENTORY_OBJS Then
124                     MiObj.amount = MAX_INVENTORY_OBJS
126                     Oro = Oro - MAX_INVENTORY_OBJS
                    Else
128                     MiObj.amount = Oro
130                     Oro = 0
                    End If

132                 Call TirarItemAlPiso(MiNPC.Pos, MiObj, MiNPC.flags.AguaValida = 1)
                Wend

134             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso("87", MiNPC.Pos.X, MiNPC.Pos.Y))
            End If

        
            Exit Sub

NPCTirarOro_Err:
136         Call RegistrarError(Err.Number, Err.Description, "NPCs.NPCTirarOro", Erl)
138         Resume Next
        
End Sub

Function OpenNPC(ByVal NpcNumber As Integer, _
                 Optional ByVal Respawn = True, _
                 Optional ByVal Reload As Boolean = False) As Integer
        
        On Error GoTo OpenNPC_Err
        

        '###################################################
        '#               ATENCION PELIGRO                  #
        '###################################################
        '
        '    ¡¡¡¡ NO USAR GetVar PARA LEER LOS NPCS !!!!
        '
        'El que ose desafiar esta LEY, se las tendrá que ver
        'conmigo. Para leer los NPCS se deberá usar la
        'nueva clase clsIniReader.
        '
        'Alejo
        '
        '###################################################

        Dim NpcIndex As Integer
    
        Dim Leer As clsIniReader
100     Set Leer = LeerNPCs

        'If requested index is invalid, abort
102     If Not Leer.KeyExists("NPC" & NpcNumber) Then
104         OpenNPC = 0
            Exit Function
        End If

106     NpcIndex = NextOpenNPC

108     If NpcIndex > MaxNPCs Then 'Limite de npcs
110         OpenNPC = 0
            Exit Function
        End If

        Dim LoopC As Long
        Dim ln    As String
        Dim aux As String
        Dim Field() As String
        
112     With NpcList(NpcIndex)

114         .Numero = NpcNumber
116         .name = Leer.GetValue("NPC" & NpcNumber, "Name")
118         .SubName = Leer.GetValue("NPC" & NpcNumber, "SubName")
120         .Desc = Leer.GetValue("NPC" & NpcNumber, "Desc")
    
122         .Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
124         .flags.OldMovement = .Movement
    
126         .flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
128         .flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
130         .flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))
    
132         .NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))
    
134         .Char.Body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
136         .Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
138         .Char.Heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))
140         .Char.BodyIdle = val(Leer.GetValue("NPC" & NpcNumber, "BodyIdle"))
    
142         .Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "Arma"))
144         .Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "Escudo"))
146         .Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "Casco"))
    
148         .Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
150         .Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
152         .Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
154         .flags.OldHostil = .Hostile
    
156         .GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP"))
    
158         .Distancia = val(Leer.GetValue("NPC" & NpcNumber, "Distancia"))
    
160         .GiveEXPClan = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXPClan"))
    
            '.flags.ExpDada = .GiveEXP
162         .flags.ExpCount = .GiveEXP
    
164         .Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))
    
166         .flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))
    
168         .GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD"))
    
        '166     .QuestNumber = val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber"))
    
    
    
    
170         .PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
172         .PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
    
174         .InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
    
176         .showName = val(Leer.GetValue("NPC" & NpcNumber, "ShowName"))
        
178         .GobernadorDe = val(Leer.GetValue("NPC" & NpcNumber, "GobernadorDe"))
    
180         .SoundOpen = val(Leer.GetValue("NPC" & NpcNumber, "SoundOpen"))
182         .SoundClose = val(Leer.GetValue("NPC" & NpcNumber, "SoundClose"))
    
184         .IntervaloAtaque = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloAtaque"))
186         .IntervaloMovimiento = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloMovimiento"))
188         .IntervaloLanzarHechizo = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloLanzarHechizo"))

190         .Contadores.IntervaloRespawn = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloRespawn"))
    
192         .InformarRespawn = val(Leer.GetValue("NPC" & NpcNumber, "InformarRespawn"))
    
194         .QuizaProb = val(Leer.GetValue("NPC" & NpcNumber, "QuizaProb"))
    
196         .SubeSupervivencia = val(Leer.GetValue("NPC" & NpcNumber, "SubeSupervivencia"))
    
198         If .IntervaloMovimiento = 0 Then
200             .IntervaloMovimiento = 380
202             .Char.speeding = 0.552631578947368
            Else
204             .Char.speeding = 210 / .IntervaloMovimiento
            End If
    
206         If .IntervaloLanzarHechizo = 0 Then
208             .IntervaloLanzarHechizo = 8000
    
            End If
    
210         If .IntervaloAtaque = 0 Then
212             .IntervaloAtaque = 2000
    
            End If
    
214         .Stats.MaxHp = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
216         .Stats.MinHp = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
218         .Stats.MaxHit = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
220         .Stats.MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
222         .Stats.def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
224         .Stats.defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
226         .flags.AIAlineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))
    
228         .Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
    
230         For LoopC = 1 To .Invent.NroItems
232             ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
234             .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
236             .Invent.Object(LoopC).amount = val(ReadField(2, ln, 45))
238         Next LoopC
    
240         .flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
    
242         If .flags.LanzaSpells > 0 Then
244             ReDim .Spells(1 To .flags.LanzaSpells)
            End If
    
246         For LoopC = 1 To .flags.LanzaSpells
248             .Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
250         Next LoopC
    
252         If .NPCtype = eNPCType.Entrenador Then
254             .NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
                
256             If .NroCriaturas > 0 Then
258                 ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador
    
260                 For LoopC = 1 To .NroCriaturas
262                     .Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
264                     .Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
266                 Next LoopC
                End If
    
            End If
    
268         .flags.NPCActive = True

270         Select Case val(Leer.GetValue("NPC" & NpcNumber, "RestriccionDeAtaque"))
                Case 0 ' Todos
272                 .flags.AtacaNPCs = True
274                 .flags.AtacaUsuarios = True
276             Case 1 ' Usuarios solamente
278                 .flags.AtacaNPCs = False
280                 .flags.AtacaUsuarios = True
282             Case 2 ' NPCs solamente
284                 .flags.AtacaNPCs = True
286                 .flags.AtacaUsuarios = False
                    
            End Select
    
288         If Respawn Then
290             .flags.Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
            Else
292             .flags.Respawn = 1
            End If
    
294         .flags.backup = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
296         .flags.RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
298         .flags.AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
300         .flags.GolpeExacto = val(Leer.GetValue("NPC" & NpcNumber, "GolpeExacto"))
    
302         .flags.Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
304         .flags.Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
306         .flags.Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))
    
            '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
308         aux = Leer.GetValue("NPC" & NpcNumber, "NROEXP")
    
310         If LenB(aux) = 0 Then
312             .NroExpresiones = 0
            
            Else
        
314             .NroExpresiones = val(aux)
            
316             ReDim .Expresiones(1 To .NroExpresiones) As String
    
318             For LoopC = 1 To .NroExpresiones
320                 .Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
322             Next LoopC
    
            End If
    
            '<<<<<<<<<<<<<< Sistema de Dropeo NUEVO >>>>>>>>>>>>>>>>
324         aux = Leer.GetValue("NPC" & NpcNumber, "NumQuiza")
    
326         If LenB(aux) = 0 Then
328             .NumQuiza = 0
            
            Else
        
330             .NumQuiza = val(aux)
            
332             ReDim .QuizaDropea(1 To .NumQuiza) As String
    
334             For LoopC = 1 To .NumQuiza
336                 .QuizaDropea(LoopC) = Leer.GetValue("NPC" & NpcNumber, "QuizaDropea" & LoopC)
338             Next LoopC
    
            End If
    
    
        'Ladder
        'Nuevo sistema de Quest
    
340         aux = Leer.GetValue("NPC" & NpcNumber, "NumQuest")
    
342         If LenB(aux) = 0 Then
344             .NumQuest = 0
            
            Else
        
346             .NumQuest = val(aux)
                
348             ReDim .QuestNumber(1 To .NumQuest) As Byte
                
350             For LoopC = 1 To .NumQuest
352                 .QuestNumber(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber" & LoopC))
354             Next LoopC
    
            End If
            
        'Nuevo sistema de Quest
    
        'Nuevo sistema de Drop Quest
            
356         aux = Leer.GetValue("NPC" & NpcNumber, "NumDropQuest")
    
358         If LenB(aux) = 0 Then
360             .NumDropQuest = 0
            
            Else
        
362             .NumDropQuest = val(aux)
                
364             ReDim .DropQuest(1 To .NumDropQuest) As tQuestObj
                
366             For LoopC = 1 To .NumDropQuest
368                 .DropQuest(LoopC).QuestIndex = val(ReadField(1, Leer.GetValue("NPC" & NpcNumber, "DropQuest" & LoopC), Asc("-")))
370                 .DropQuest(LoopC).ObjIndex = val(ReadField(2, Leer.GetValue("NPC" & NpcNumber, "DropQuest" & LoopC), Asc("-")))
372                 .DropQuest(LoopC).amount = val(ReadField(3, Leer.GetValue("NPC" & NpcNumber, "DropQuest" & LoopC), Asc("-")))
374                 .DropQuest(LoopC).Probabilidad = val(ReadField(4, Leer.GetValue("NPC" & NpcNumber, "DropQuest" & LoopC), Asc("-")))
376             Next LoopC
    
            End If
        
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< PATHFINDING >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
378         .pathFindingInfo.RangoVision = val(Leer.GetValue("NPC" & NpcNumber, "Distancia"))
380         If .pathFindingInfo.RangoVision = 0 Then .pathFindingInfo.RangoVision = RANGO_VISION_X
        
382         .pathFindingInfo.Inteligencia = val(Leer.GetValue("NPC" & NpcNumber, "Inteligencia"))
384         If .pathFindingInfo.Inteligencia = 0 Then .pathFindingInfo.Inteligencia = 30
        
386         ReDim .pathFindingInfo.Path(1 To .pathFindingInfo.Inteligencia + RANGO_VISION_X * 3)
    
            '<<<<<<<<<<<<<< Sistema de Viajes NUEVO >>>>>>>>>>>>>>>>
388         aux = Leer.GetValue("NPC" & NpcNumber, "NumDestinos")
    
390         If LenB(aux) = 0 Then
392             .NumDestinos = 0
            
            Else
        
394             .NumDestinos = val(aux)
            
396             ReDim .Dest(1 To .NumDestinos) As String
    
398             For LoopC = 1 To .NumDestinos
400                 .Dest(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Dest" & LoopC)
402             Next LoopC
    
            End If
    
            '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
404         .Interface = val(Leer.GetValue("NPC" & NpcNumber, "Interface"))
    
            'Tipo de items con los que comercia
406         .TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))
    
    
            '<<<<<<<<<<<<<< Animaciones >>>>>>>>>>>>>>>>
    
            ' Por defecto la animación es idle
408         Call AnimacionIdle(NpcIndex, True)
    
            ' Si el tipo de movimiento es Caminata
410         If .Movement = Caminata Then
                ' Leemos la cantidad de indicaciones
                Dim cant As Byte
412             cant = val(Leer.GetValue("NPC" & NpcNumber, "CaminataLen"))
                ' Prevengo NPCs rotos
414             If cant = 0 Then
416                 .Movement = Estatico
                Else
                    ' Redimenciono el array
418                 ReDim .Caminata(1 To cant)
                    ' Leo todas las indicaciones
420                 For LoopC = 1 To cant
422                     Field = Split(Leer.GetValue("NPC" & NpcNumber, "Caminata" & LoopC), ":")
    
424                     .Caminata(LoopC).Offset.X = val(Field(0))
426                     .Caminata(LoopC).Offset.Y = val(Field(1))
428                     .Caminata(LoopC).Espera = val(Field(2))
                    Next
                    
430                 .CaminataActual = 1
                End If
            End If
            '<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>
            
        End With

        'Si NO estamos actualizando los NPC's activos, actualizamos el contador.
432     If Reload = False Then
434         If NpcIndex > LastNPC Then LastNPC = NpcIndex
436         NumNPCs = NumNPCs + 1
        End If
    
        'Devuelve el nuevo Indice
438     OpenNPC = NpcIndex

        Exit Function

OpenNPC_Err:
440     Call RegistrarError(Err.Number, Err.Description, "NPCs.OpenNPC", Erl)
442     Resume Next
        
End Function

Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)
        
        On Error GoTo DoFollow_Err
        
100     With NpcList(NpcIndex)
    
102         If .flags.Follow Then
        
104             .flags.AttackedBy = vbNullString
106             .Target = 0
108             .flags.Follow = False
110             .Movement = .flags.OldMovement
112             .Hostile = .flags.OldHostil
   
            Else
        
114             .flags.AttackedBy = UserName
116             .Target = NameIndex(UserName)
118             .flags.Follow = True
120             .Movement = TipoAI.NpcDefensa
122             .Hostile = 0

            End If
    
        End With
        
        Exit Sub

DoFollow_Err:
124     Call RegistrarError(Err.Number, Err.Description, "NPCs.DoFollow", Erl)
126     Resume Next
        
End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)
        On Error GoTo FollowAmo_Err

100     With NpcList(NpcIndex)
102         .flags.Follow = True
104         .Movement = TipoAI.SigueAmo
106         .Hostile = 0
108         .Target = 0
110         .TargetNPC = 0
        End With

        Exit Sub

FollowAmo_Err:
112     Call RegistrarError(Err.Number, Err.Description, "NPCs.FollowAmo", Erl)
End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)
            On Error GoTo AllFollowAmo_Err

            Dim j As Long

100         For j = 1 To MAXMASCOTAS
102             If UserList(UserIndex).MascotasIndex(j) > 0 Then
104                 Call FollowAmo(UserList(UserIndex).MascotasIndex(j))
                End If
106         Next j

            Exit Sub

AllFollowAmo_Err:
108         Call RegistrarError(Err.Number, Err.Description, "SistemaCombate.AllFollowAmo", Erl)

End Sub


Public Function ObtenerIndiceRespawn() As Integer

        On Error GoTo ErrHandler

        Dim LoopC As Integer

100     For LoopC = 1 To MaxRespawn
102         If Not RespawnList(LoopC).flags.NPCActive Then Exit For
104     Next LoopC
  
106     ObtenerIndiceRespawn = LoopC

        Exit Function
ErrHandler:
108     Call LogError("Error en ObtenerIndiceRespawn")

End Function

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        
        On Error GoTo QuitarMascota_Err
    
        

        Dim i As Integer
    
100     For i = 1 To MAXMASCOTAS

102         If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
104             UserList(UserIndex).MascotasIndex(i) = 0
106             UserList(UserIndex).MascotasType(i) = 0
         
108             UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
                Exit For

            End If

110     Next i

        
        Exit Sub

QuitarMascota_Err:
112     Call RegistrarError(Err.Number, Err.Description, "NPCs.QuitarMascota", Erl)

        
End Sub

Sub AnimacionIdle(ByVal NpcIndex As Integer, ByVal Show As Boolean)
    
        On Error GoTo Handler
    
100     With NpcList(NpcIndex)
    
102         If .Char.BodyIdle = 0 Then Exit Sub
        
104         If .flags.NPCIdle = Show Then Exit Sub

106         .flags.NPCIdle = Show
        
108         Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, .Char.Heading)
        
        End With
    
        Exit Sub
Handler:
110     Call RegistrarError(Err.Number, Err.Description, "NPCs.AnimacionIdle", Erl)
112     Resume Next
End Sub

Sub WarpNpcChar(ByVal NpcIndex As Integer, ByVal Map As Byte, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)

        Dim NuevaPos                    As WorldPos
        Dim FuturePos                   As WorldPos

100     Call EraseNPCChar(NpcIndex)

102     FuturePos.Map = Map
104     FuturePos.X = X
106     FuturePos.Y = Y
108     Call ClosestLegalPos(FuturePos, NuevaPos, True, True)

110     If NuevaPos.Map = 0 Or NuevaPos.X = 0 Or NuevaPos.Y = 0 Then
112         Debug.Print "Error al tepear NPC"
114         Call QuitarNPC(NpcIndex)
        Else
116         NpcList(NpcIndex).Pos = NuevaPos
118         Call MakeNPCChar(True, 0, NpcIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

120         If FX Then                                    'FX
122             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_WARP, NuevaPos.X, NuevaPos.Y))
124             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(NpcList(NpcIndex).Char.CharIndex, FXIDs.FXWARP, 0))
            End If

        End If

End Sub

' Autor: WyroX - 20/01/2021
' Intenta moverlo hacia un "costado" según el heading indicado. Se usa para mover NPCs del camino de otro char.
' Si no hay un lugar válido a los lados, lo mueve a la posición válida más cercana.
Sub MoveNpcToSide(ByVal NpcIndex As Integer, ByVal Heading As eHeading)

        On Error GoTo Handler

100     With NpcList(NpcIndex)

            ' Elegimos un lado al azar
            Dim R As Integer
102         R = RandomNumber(0, 1) * 2 - 1 ' -1 o 1

            ' Roto el heading original hacia ese lado
104         Heading = RotateHeading(Heading, R)

            ' Intento moverlo para ese lado
106         If MoveNPCChar(NpcIndex, Heading) Then Exit Sub
        
            ' Si falló, intento moverlo para el lado opuesto
108         Heading = InvertHeading(Heading)
110         If MoveNPCChar(NpcIndex, Heading) Then Exit Sub
        
            ' Si ambos fallan, entonces lo dejo en la posición válida más cercana
            Dim NuevaPos As WorldPos
112         Call ClosestLegalPos(.Pos, NuevaPos, .flags.AguaValida, .flags.TierraInvalida = 0)
114         Call WarpNpcChar(NpcIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

        End With

        Exit Sub
    
Handler:
116     Call RegistrarError(Err.Number, Err.Description, "NPCs.MoveNpcToSide", Erl)
118     Resume Next
End Sub
