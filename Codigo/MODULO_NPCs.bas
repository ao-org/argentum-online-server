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
        
100     Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1

        
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
        If Npclist(NpcIndex).NPCtype = DummyTarget Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead("¡¡Auch!!", Npclist(NpcIndex).Char.CharIndex, vbRed, "Barrin"))
            Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MaxHp
            Exit Sub
        End If

        Dim MiNPC As npc

100     MiNPC = Npclist(NpcIndex)

        Dim EraCriminal As Byte

        Dim TiempoRespw As Integer

102     TiempoRespw = Npclist(NpcIndex).Contadores.InvervaloRespawn

        ' Es pretoriano?
        If MiNPC.NPCtype = eNPCType.Pretoriano Then
            Call ClanPretoriano(MiNPC.ClanIndex).MuerePretoriano(NpcIndex)
        End If

        'Quitamos el npc
104     Call QuitarNPC(NpcIndex)
    
106     If UserIndex > 0 Then ' Lo mato un usuario?
108         If MiNPC.flags.Snd3 > 0 Then
110             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.Pos.X, MiNPC.Pos.Y))
            Else
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("28", MiNPC.Pos.X, MiNPC.Pos.Y))
            End If

114         UserList(UserIndex).flags.TargetNPC = 0
116         UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        
            'El user que lo mato tiene mascotas?
120         Call AllFollowAmo(UserIndex)
            
140         If UserList(UserIndex).ChatCombate = 1 Then
150             Call WriteLocaleMsg(UserIndex, "184", FontTypeNames.FONTTYPE_FIGHT, "la criatura")
            End If

152         If UserList(UserIndex).Stats.NPCsMuertos < 32000 Then UserList(UserIndex).Stats.NPCsMuertos = UserList(UserIndex).Stats.NPCsMuertos + 1
            
            If MiNPC.MaestroUser > 0 Then Exit Sub
            
118         Call SubirSkill(UserIndex, eSkill.Supervivencia)

132         If MiNPC.flags.ExpCount > 0 Then

134             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
136                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + MiNPC.flags.ExpCount

138                 If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
                    
                    Call WriteTextOverTile(UserIndex, "+" & PonerPuntos(MiNPC.flags.ExpCount), UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, RGB(0, 169, 255))
142                 Call WriteUpdateExp(UserIndex)
144                 Call CheckUserLevel(UserIndex)

                End If
        
146             MiNPC.flags.ExpCount = 0

            End If
        
154         EraCriminal = Status(UserIndex)
        
            'If MiNPC.Stats.Alineacion = 0 Then
            '  If MiNPC.Numero = Guardias Then
            'UserList(UserIndex).Reputacion.NobleRep = 0
               
            'End If
            'ElseIf MiNPC.Stats.Alineacion = 1 Then
            ' UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep + vlCAZADOR
            ' ElseIf MiNPC.Stats.Alineacion = 2 Then
            'UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep + vlASESINO / 2
            
            'ElseIf MiNPC.Stats.Alineacion = 4 Then
            ' UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep + vlCAZADOR
            
            'End If
            ' If Status(UserIndex) = 0 And esArmada(UserIndex) Then Call ExpulsarFaccionReal(UserIndex)
            ' If Status(UserIndex) = 2 And esCaos(UserIndex) Then Call ExpulsarFaccionCaos(UserIndex)
        
            'If EraCriminal = 2 And Status(UserIndex) < 2 Then
            '    Call RefreshCharStatus(UserIndex)
            'ElseIf EraCriminal < 2 And Status(UserIndex) = 2 Then
            '    Call RefreshCharStatus(UserIndex)
            'End If
        
156         If MiNPC.GiveEXPClan > 0 Then
158             If UserList(UserIndex).GuildIndex > 0 Then
160                 Call modGuilds.CheckClanExp(UserIndex, MiNPC.GiveEXPClan)

                    ' Else
                    ' Call WriteConsoleMsg(UserIndex, "No perteneces a ningun clan, experiencia perdida.", FontTypeNames.FONTTYPE_INFOIAO)
                End If

            End If
        
            Dim i As Long, j As Long
        
162         For i = 1 To MAXUSERQUESTS
        
164             With UserList(UserIndex).QuestStats.Quests(i)
        
166                 If .QuestIndex Then
168                     If QuestList(.QuestIndex).RequiredNPCs Then
        
170                         For j = 1 To QuestList(.QuestIndex).RequiredNPCs
        
172                             If QuestList(.QuestIndex).RequiredNPC(j).NpcIndex = MiNPC.Numero Then
174                                 If QuestList(.QuestIndex).RequiredNPC(j).Amount > .NPCsKilled(j) Then
176                                     .NPCsKilled(j) = .NPCsKilled(j) + 1
        
                                    End If
                                    
178                                 If QuestList(.QuestIndex).RequiredNPC(j).Amount = .NPCsKilled(j) Then
180                                     Call WriteConsoleMsg(UserIndex, "Ya has matado todos los " & MiNPC.name & " que la mision " & QuestList(.QuestIndex).nombre & " requeria. Chequeá si ya estas listo para recibir la recompensa.", FontTypeNames.FONTTYPE_INFOIAO)
                                    
                                    End If
        
                                End If
        
182                         Next j
        
                        End If
        
                    End If
        
                End With
        
184         Next i

            'Tiramos el oro
188         Call NPCTirarOro(MiNPC, UserIndex)

190         Call DropObjQuest(MiNPC, UserIndex)
    
            'Item Magico!
192         Call NpcDropeo(MiNPC, UserIndex)
            
            'Tiramos el inventario
194         Call NPC_TIRAR_ITEMS(MiNPC)
            
        End If ' UserIndex > 0
        
        ' Mascotas y npcs de entrenamiento no respawnean
        If MiNPC.MaestroNPC > 0 Or MiNPC.MaestroUser > 0 Then Exit Sub

        'ReSpawn o no
196     If TiempoRespw = 0 Then
198         Call ReSpawnNpc(MiNPC)

        Else

            Dim Indice As Integer

200         MiNPC.flags.NPCActive = True
202         Indice = ObtenerIndiceRespawn
204         RespawnList(Indice) = MiNPC

        End If
    
        Exit Sub

ErrHandler:
206     Call RegistrarError(Err.Number, Err.Description, "NPCs.MuereNpc", Erl())

End Sub

Sub ResetNpcFlags(ByVal NpcIndex As Integer)
        'Clear the npc's flags
        
        On Error GoTo ResetNpcFlags_Err
        
    
100     With Npclist(NpcIndex).flags
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
150         .UseAINow = False
152         .AtacaAPJ = 0
154         .AtacaANPC = 0
156         .AIAlineacion = e_Alineacion.ninguna
158         .AIPersonalidad = e_Personalidad.ninguna
160         .NPCIdle = False
        End With

        
        Exit Sub

ResetNpcFlags_Err:
162     Call RegistrarError(Err.Number, Err.Description, "NPCs.ResetNpcFlags", Erl)
164     Resume Next
        
End Sub

Sub ResetNpcCounters(ByVal NpcIndex As Integer)
        
        On Error GoTo ResetNpcCounters_Err
        

100     Npclist(NpcIndex).Contadores.Paralisis = 0
102     Npclist(NpcIndex).Contadores.TiempoExistencia = 0
104     Npclist(NpcIndex).Contadores.IntervaloMovimiento = 0
106     Npclist(NpcIndex).Contadores.IntervaloAtaque = 0
108     Npclist(NpcIndex).Contadores.InvervaloLanzarHechizo = 0
110     Npclist(NpcIndex).Contadores.InvervaloRespawn = 0

        
        Exit Sub

ResetNpcCounters_Err:
112     Call RegistrarError(Err.Number, Err.Description, "NPCs.ResetNpcCounters", Erl)
114     Resume Next
        
End Sub

Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)
        
        On Error GoTo ResetNpcCharInfo_Err
        

100     Npclist(NpcIndex).Char.Body = 0
102     Npclist(NpcIndex).Char.CascoAnim = 0
104     Npclist(NpcIndex).Char.CharIndex = 0
106     Npclist(NpcIndex).Char.FX = 0
108     Npclist(NpcIndex).Char.Head = 0
110     Npclist(NpcIndex).Char.Heading = 0
112     Npclist(NpcIndex).Char.loops = 0
114     Npclist(NpcIndex).Char.ShieldAnim = 0
116     Npclist(NpcIndex).Char.WeaponAnim = 0

        
        Exit Sub

ResetNpcCharInfo_Err:
118     Call RegistrarError(Err.Number, Err.Description, "NPCs.ResetNpcCharInfo", Erl)
120     Resume Next
        
End Sub

Sub ResetNpcCriatures(ByVal NpcIndex As Integer)
        
        On Error GoTo ResetNpcCriatures_Err
        

        Dim j As Integer

100     For j = 1 To Npclist(NpcIndex).NroCriaturas
102         Npclist(NpcIndex).Criaturas(j).NpcIndex = 0
104         Npclist(NpcIndex).Criaturas(j).NpcName = vbNullString
106     Next j

108     Npclist(NpcIndex).NroCriaturas = 0

        
        Exit Sub

ResetNpcCriatures_Err:
110     Call RegistrarError(Err.Number, Err.Description, "NPCs.ResetNpcCriatures", Erl)
112     Resume Next
        
End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)
        
        On Error GoTo ResetExpresiones_Err
        

        Dim j As Integer

100     For j = 1 To Npclist(NpcIndex).NroExpresiones
102         Npclist(NpcIndex).Expresiones(j) = vbNullString
104     Next j

106     Npclist(NpcIndex).NroExpresiones = 0

        
        Exit Sub

ResetExpresiones_Err:
108     Call RegistrarError(Err.Number, Err.Description, "NPCs.ResetExpresiones", Erl)
110     Resume Next
        
End Sub

Sub ResetDrop(ByVal NpcIndex As Integer)
        
        On Error GoTo ResetDrop_Err
        

        Dim j As Integer

100     For j = 1 To Npclist(NpcIndex).NumQuiza
102         Npclist(NpcIndex).QuizaDropea(j) = 0
104     Next j

106     Npclist(NpcIndex).NumQuiza = 0

        
        Exit Sub

ResetDrop_Err:
108     Call RegistrarError(Err.Number, Err.Description, "NPCs.ResetDrop", Erl)
110     Resume Next
        
End Sub

Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)
        
        On Error GoTo ResetNpcMainInfo_Err
        
    
100     Npclist(NpcIndex).Attackable = 0
102     Npclist(NpcIndex).CanAttack = 0
104     Npclist(NpcIndex).Comercia = 0
106     Npclist(NpcIndex).GiveEXP = 0
108     Npclist(NpcIndex).GiveEXPClan = 0
110     Npclist(NpcIndex).GiveGLD = 0
112     Npclist(NpcIndex).Hostile = 0
114     Npclist(NpcIndex).InvReSpawn = 0
116     Npclist(NpcIndex).level = 0

    
118     Npclist(NpcIndex).MaestroNPC = 0
    
120     Npclist(NpcIndex).Mascotas = 0
122     Npclist(NpcIndex).Movement = 0
124     Npclist(NpcIndex).name = "NPC SIN INICIAR"
126     Npclist(NpcIndex).NPCtype = 0
128     Npclist(NpcIndex).Numero = 0
130     Npclist(NpcIndex).Orig.Map = 0
132     Npclist(NpcIndex).Orig.X = 0
134     Npclist(NpcIndex).Orig.Y = 0
136     Npclist(NpcIndex).PoderAtaque = 0
138     Npclist(NpcIndex).PoderEvasion = 0
140     Npclist(NpcIndex).Pos.Map = 0
142     Npclist(NpcIndex).Pos.X = 0
144     Npclist(NpcIndex).Pos.Y = 0
146     Npclist(NpcIndex).Target = 0
148     Npclist(NpcIndex).TargetNPC = 0
150     Npclist(NpcIndex).TipoItems = 0
152     Npclist(NpcIndex).Veneno = 0
154     Npclist(NpcIndex).Desc = vbNullString
156     Npclist(NpcIndex).NumDropQuest = 0
        
158     If Npclist(NpcIndex).MaestroUser > 0 Then Call QuitarMascota(Npclist(NpcIndex).MaestroUser, NpcIndex)
160     If Npclist(NpcIndex).MaestroNPC > 0 Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNPC)

162     Npclist(NpcIndex).MaestroUser = 0
164     Npclist(NpcIndex).MaestroNPC = 0

        Npclist(NpcIndex).CaminataActual = 0
    
        Dim j As Integer

166     For j = 1 To Npclist(NpcIndex).NroSpells
168         Npclist(NpcIndex).Spells(j) = 0
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

100     Npclist(NpcIndex).flags.NPCActive = False
    
102     If InMapBounds(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then
104         Call EraseNPCChar(NpcIndex)
        End If
    
        ' Es pretoriano?
106     If Npclist(NpcIndex).NPCtype = eNPCType.Pretoriano Then
108         Call ClanPretoriano(Npclist(NpcIndex).ClanIndex).MuerePretoriano(NpcIndex)
        End If
    
        'Nos aseguramos de que el inventario sea removido...
        'asi los lobos no volveran a tirar armaduras ;))
110     Call ResetNpcInv(NpcIndex)
112     Call ResetNpcFlags(NpcIndex)
114     Call ResetNpcCounters(NpcIndex)
    
116     Call ResetNpcMainInfo(NpcIndex)
    
118     If NpcIndex = LastNPC Then

120         Do Until Npclist(LastNPC).flags.NPCActive
122             LastNPC = LastNPC - 1

124             If LastNPC < 1 Then Exit Do
            Loop

        End If
      
126     If NumNPCs <> 0 Then
128         NumNPCs = NumNPCs - 1

        End If

        Exit Sub

ErrHandler:
130     Npclist(NpcIndex).flags.NPCActive = False
132     Call LogError("Error en QuitarNPC")

End Sub

Function TestSpawnTrigger(Pos As WorldPos, Optional PuedeAgua As Boolean = False) As Boolean
        
        On Error GoTo TestSpawnTrigger_Err
        
    
100     If LegalPos(Pos.Map, Pos.X, Pos.Y, PuedeAgua) Then
102         TestSpawnTrigger = MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 3 And MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 2 And MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 1

        End If

        
        Exit Function

TestSpawnTrigger_Err:
104     Call RegistrarError(Err.Number, Err.Description, "NPCs.TestSpawnTrigger", Erl)
106     Resume Next
        
End Function

Public Function CrearNPC(NroNPC As Integer, Mapa As Integer, OrigPos As WorldPos, Optional ByVal CustomHead As Integer)
        'Call LogTarea("Sub CrearNPC")
        'Crea un NPC del tipo NRONPC
        
        On Error GoTo CrearNPC_Err
        

        Dim Pos            As WorldPos
        Dim newpos         As WorldPos
        Dim altpos         As WorldPos

        Dim nIndex         As Integer
        Dim PosicionValida As Boolean
        Dim Iteraciones    As Long

        Dim PuedeAgua      As Boolean
        Dim PuedeTierra    As Boolean

        Dim Map            As Integer
        Dim X              As Integer
        Dim Y              As Integer

100     nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    
102     If nIndex = 0 Then Exit Function
        
        ' Cabeza customizada
104     If CustomHead <> 0 Then Npclist(nIndex).Char.Head = CustomHead

106     PuedeAgua = Npclist(nIndex).flags.AguaValida = 1
108     PuedeTierra = Npclist(nIndex).flags.TierraInvalida = 0
    
        'Necesita ser respawned en un lugar especifico
110     If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
        
112         Map = OrigPos.Map
114         X = OrigPos.X
116         Y = OrigPos.Y
118         Npclist(nIndex).Orig = OrigPos
120         Npclist(nIndex).Pos = OrigPos
       
        Else
        
122         Pos.Map = Mapa 'mapa
124         altpos.Map = Mapa
        
126         Do While Not PosicionValida
128             Pos.X = RandomNumber(MinXBorder + 2, MaxXBorder - 2) 'Obtenemos posicion al azar en x
130             Pos.Y = RandomNumber(MinYBorder + 2, MaxYBorder - 2) 'Obtenemos posicion al azar en y
            
132             Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana

134             If newpos.X <> 0 And newpos.Y <> 0 Then
136                 altpos.X = newpos.X
138                 altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn, pero intentando qeu si tenía que ser en el agua, sea en el agua.)
                Else
140                 Call ClosestLegalPos(Pos, newpos, PuedeAgua)

142                 If newpos.X <> 0 And newpos.Y <> 0 Then
144                     altpos.X = newpos.X
146                     altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn)

                    End If

                End If

                'Si X e Y son iguales a 0 significa que no se encontro posicion valida
148             If LegalPosNPC(newpos.Map, newpos.X, newpos.Y, PuedeAgua) And Not HayPCarea(newpos) And TestSpawnTrigger(newpos, PuedeAgua) Then

                    'Asignamos las nuevas coordenas solo si son validas
150                 Npclist(nIndex).Pos.Map = newpos.Map
152                 Npclist(nIndex).Pos.X = newpos.X
154                 Npclist(nIndex).Pos.Y = newpos.Y
156                 PosicionValida = True

                Else
                
158                 newpos.X = 0
160                 newpos.Y = 0
            
                End If
                
                'for debug
162             Iteraciones = Iteraciones + 1

164             If Iteraciones > MAXSPAWNATTEMPS Then

166                 If altpos.X <> 0 And altpos.Y <> 0 Then

168                     Map = altpos.Map
170                     X = altpos.X
172                     Y = altpos.Y

174                     Npclist(nIndex).Pos.Map = Map
176                     Npclist(nIndex).Pos.X = X
178                     Npclist(nIndex).Pos.Y = Y

180                     Call MakeNPCChar(True, Map, nIndex, Map, X, Y)
182                     CrearNPC = nIndex
                        
                        Exit Function
                    
                    Else
                    
184                     altpos.X = 50
186                     altpos.Y = 50
188                     Call ClosestLegalPos(altpos, newpos)

190                     If newpos.X <> 0 And newpos.Y <> 0 Then

192                         Npclist(nIndex).Pos.Map = newpos.Map
194                         Npclist(nIndex).Pos.X = newpos.X
196                         Npclist(nIndex).Pos.Y = newpos.Y

198                         Call MakeNPCChar(True, newpos.Map, nIndex, newpos.Map, newpos.X, newpos.Y)
200                         CrearNPC = nIndex
                            
                            Exit Function
                            
                        Else
                        
202                         Call QuitarNPC(nIndex)
204                         Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & Mapa & " NroNpc:" & NroNPC)
                            Exit Function

                        End If

                    End If

                End If

            Loop
        
            'asignamos las nuevas coordenas
206         Map = newpos.Map
208         X = Npclist(nIndex).Pos.X
210         Y = Npclist(nIndex).Pos.Y

        End If
    
        'Crea el NPC
212     Call MakeNPCChar(True, Map, nIndex, Map, X, Y)
        
214     CrearNPC = nIndex
        
        Exit Function

CrearNPC_Err:
216     Call RegistrarError(Err.Number, Err.Description, "NPCs.CrearNPC", Erl)
218     Resume Next
        
End Function

Sub MakeNPCChar(ByVal toMap As Boolean, sndIndex As Integer, NpcIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo MakeNPCChar_Err
        
100     With Npclist(NpcIndex)

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
                
                
                If UserList(sndIndex).Stats.UserSkills(eSkill.Supervivencia) >= 90 Then
156                 Call WriteCharacterCreate(sndIndex, IIf(.flags.NPCIdle, .Char.BodyIdle, .Char.Body), .Char.Head, .Char.Heading, .Char.CharIndex, X, Y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, GG, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, .Char.speeding, IIf(.MaestroUser = sndIndex, 2, 1), False, 0, 0, 0, 0, .Stats.MinHp, .Stats.MaxHp, Simbolo, .flags.NPCIdle)
                Else
                    Call WriteCharacterCreate(sndIndex, IIf(.flags.NPCIdle, .Char.BodyIdle, .Char.Body), .Char.Head, .Char.Heading, .Char.CharIndex, X, Y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, GG, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, .Char.speeding, IIf(.MaestroUser = sndIndex, 2, 1), False, 0, 0, 0, 0, 0, 0, Simbolo, .flags.NPCIdle)
                
                End If
            Else
158             Call AgregarNpc(NpcIndex)
    
            End If

        End With
        
        Exit Sub

MakeNPCChar_Err:
160     Call RegistrarError(Err.Number, Err.Description, "NPCs.MakeNPCChar", Erl)
162     Resume Next
        
End Sub

Sub ChangeNPCChar(ByVal NpcIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading)
        
        On Error GoTo ChangeNPCChar_Err
        
100     With Npclist(NpcIndex)

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
        

100     If Npclist(NpcIndex).Char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

102     If Npclist(NpcIndex).Char.CharIndex = LastChar Then

104         Do Until CharList(LastChar) > 0
106             LastChar = LastChar - 1

108             If LastChar <= 1 Then Exit Do
            Loop

        End If

        'Quitamos del mapa
110     MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

        'Actualizamos los clientes
112     Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterRemove(Npclist(NpcIndex).Char.CharIndex, True))

        'Update la lista npc
114     Npclist(NpcIndex).Char.CharIndex = 0

        'update NumChars
116     NumChars = NumChars - 1

        
        Exit Sub

EraseNPCChar_Err:
118     Call RegistrarError(Err.Number, Err.Description, "NPCs.EraseNPCChar", Erl)
120     Resume Next
        
End Sub

Public Function MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte) As Boolean
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 06/04/2009
        '06/04/2009: ZaMa - Now npcs can force to change position with dead character
        '01/08/2009: ZaMa - Now npcs can't force to chance position with a dead character if that means to change the terrain the character is in
        '26/09/2010: ZaMa - Turn sub into function to know if npc has moved or not.
        '***************************************************

        On Error GoTo errh

        Dim nPos      As WorldPos
        Dim UserIndex As Integer
    
100     With Npclist(NpcIndex)
102         nPos = .Pos
104         Call HeadtoPos(nHeading, nPos)
        
            ' es una posicion legal
106         If LegalWalkNPC(nPos.Map, nPos.X, nPos.Y, nHeading, .flags.AguaValida = 1, .flags.TierraInvalida = 0, .MaestroUser <> 0) Then
            
108             UserIndex = MapData(.Pos.Map, nPos.X, nPos.Y).UserIndex

                ' Si hay un usuario a donde se mueve el npc, entonces esta muerto o es un gm invisible
110             If UserIndex > 0 Then

112                 With UserList(UserIndex)
                
                        ' Actualizamos posicion y mapa
114                     MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
116                     .Pos.X = Npclist(NpcIndex).Pos.X
118                     .Pos.Y = Npclist(NpcIndex).Pos.Y
120                     MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                        
                        ' Avisamos a los usuarios del area, y al propio usuario lo forzamos a moverse
122                     Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, .Pos.X, .Pos.Y))
124                     Call WriteForceCharMove(UserIndex, InvertHeading(nHeading))

                    End With

                End If
            
126             Call AnimacionIdle(NpcIndex, False)
            
128             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))

                'Update map and user pos
130             MapData(.Pos.Map, .Pos.X, .Pos.Y).NpcIndex = 0
132             .Pos = nPos
134             .Char.Heading = nHeading
136             MapData(.Pos.Map, nPos.X, nPos.Y).NpcIndex = NpcIndex
            
138             Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
                ' Npc has moved
140             MoveNPCChar = True
        
142         ElseIf .MaestroUser = 0 Then

144             If .Movement = TipoAI.NpcPathfinding Then
                    'Someone has blocked the npc's way, we must to seek a new path!
146                 .PFINFO.PathLenght = 0
                End If

            End If

        End With
    
        Exit Function

errh:
148     LogError ("Error en move npc " & NpcIndex & ". Error: " & Err.Number & " - " & Err.Description)

End Function

Function NextOpenNPC() As Integer
        'Call LogTarea("Sub NextOpenNPC")

        On Error GoTo ErrHandler

        Dim LoopC As Integer
  
100     For LoopC = 1 To MAXNPCS + 1

102         If LoopC > MAXNPCS Then Exit For

104         If Not Npclist(LoopC).flags.NPCActive Then Exit For

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

        Dim PosicionValida As Boolean

        Dim PuedeAgua      As Boolean

        Dim PuedeTierra    As Boolean

        Dim Map            As Integer

        Dim X              As Integer

        Dim Y              As Integer

        Dim it             As Integer

100     nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

102     If nIndex > MAXNPCS Then
104         SpawnNpc = 0
            Exit Function

        End If

106     PuedeAgua = Npclist(nIndex).flags.AguaValida = 1
108     PuedeTierra = Npclist(nIndex).flags.TierraInvalida = 0

110     it = 0

112     Do While Not PosicionValida
        
114         Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
116         Call ClosestLegalPos(Pos, altpos, PuedeAgua)
            'Si X e Y son iguales a 0 significa que no se encontro posicion valida

118         If newpos.X <> 0 And newpos.Y <> 0 Then
                'Asignamos las nuevas coordenas solo si son validas
120             Npclist(nIndex).Pos.Map = newpos.Map
122             Npclist(nIndex).Pos.X = newpos.X
124             Npclist(nIndex).Pos.Y = newpos.Y
126             PosicionValida = True
            Else

128             If altpos.X <> 0 And altpos.Y <> 0 Then
130                 Npclist(nIndex).Pos.Map = altpos.Map
132                 Npclist(nIndex).Pos.X = altpos.X
134                 Npclist(nIndex).Pos.Y = altpos.Y
136                 PosicionValida = True
                Else
138                 PosicionValida = False

                End If

            End If
        
140         it = it + 1
        
142         If it > MAXSPAWNATTEMPS Then
144             Call QuitarNPC(nIndex)
146             SpawnNpc = 0
148             Call LogError("Mas de " & MAXSPAWNATTEMPS & " iteraciones en SpawnNpc Mapa:" & Pos.Map & " Index:" & NpcIndex)
                Exit Function

            End If

        Loop

        'asignamos las nuevas coordenas
150     Map = newpos.Map
152     X = Npclist(nIndex).Pos.X
154     Y = Npclist(nIndex).Pos.Y

        ' WyroX: Asignamos el dueño
        Npclist(nIndex).MaestroUser = MaestroUser
        
        ' WyroX: Caminata de NPCs
        If Npclist(nIndex).Movement = Caminata Then
            Npclist(nIndex).Orig = Npclist(nIndex).Pos
        End If

        'Crea el NPC
156     Call MakeNPCChar(True, Map, nIndex, Map, X, Y)

158     If FX Then
160         Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
162         Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).Char.CharIndex, FXIDs.FXWARP, 0))

        End If

164     If Avisar Then
166         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Npclist(nIndex).name & " ha aparecido en " & DarNameMapa(Map) & " , todo indica que puede tener una gran recompensa para el que logre sobrevivir a él.", FontTypeNames.FONTTYPE_CITIZEN))

        End If

168     SpawnNpc = nIndex

        
        Exit Function

SpawnNpc_Err:
170     Call RegistrarError(Err.Number, Err.Description, "NPCs.SpawnNpc", Erl)
172     Resume Next
        
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
104         If Npclist(NpcIndex).flags.NPCActive And Npclist(NpcIndex).Pos.Map = Map And Npclist(NpcIndex).Hostile = 1 And Npclist(NpcIndex).Stats.Alineacion = 2 Then
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
124                     MiObj.Amount = MAX_INVENTORY_OBJS
126                     Oro = Oro - MAX_INVENTORY_OBJS
                    Else
128                     MiObj.Amount = Oro
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
104         OpenNPC = MAXNPCS + 1
            Exit Function
        End If

106     NpcIndex = NextOpenNPC

108     If NpcIndex > MAXNPCS Then 'Limite de npcs
110         OpenNPC = NpcIndex
            Exit Function
        End If

        Dim LoopC As Long
        Dim ln    As String
        Dim Aux As String
        Dim Field() As String
        
        With Npclist(NpcIndex)

112         .Numero = NpcNumber
114         .name = Leer.GetValue("NPC" & NpcNumber, "Name")
116         .SubName = Leer.GetValue("NPC" & NpcNumber, "SubName")
118         .Desc = Leer.GetValue("NPC" & NpcNumber, "Desc")
    
120         .Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
122         .flags.OldMovement = .Movement
    
124         .flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
126         .flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
128         .flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))
    
130         .NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))
    
132         .Char.Body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
134         .Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
136         .Char.Heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))
138         .Char.BodyIdle = val(Leer.GetValue("NPC" & NpcNumber, "BodyIdle"))
    
140         .Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "Arma"))
142         .Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "Escudo"))
144         .Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "Casco"))
    
146         .Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
148         .Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
150         .Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
152         .flags.OldHostil = .Hostile
    
154         .GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP"))
    
156         .Distancia = val(Leer.GetValue("NPC" & NpcNumber, "Distancia"))
    
158         .GiveEXPClan = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXPClan"))
    
            '.flags.ExpDada = .GiveEXP
160         .flags.ExpCount = .GiveEXP
    
162         .Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))
    
164         .flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))
    
166         .GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD"))
    
        '166     .QuestNumber = val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber"))
    
    
    
    
168         .PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
170         .PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
    
172         .InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
    
174         .showName = val(Leer.GetValue("NPC" & NpcNumber, "ShowName"))
        
176         .GobernadorDe = val(Leer.GetValue("NPC" & NpcNumber, "GobernadorDe"))
    
178         .SoundOpen = val(Leer.GetValue("NPC" & NpcNumber, "SoundOpen"))
180         .SoundClose = val(Leer.GetValue("NPC" & NpcNumber, "SoundClose"))
    
182         .IntervaloAtaque = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloAtaque"))
184         .IntervaloMovimiento = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloMovimiento"))
186         .InvervaloLanzarHechizo = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloLanzarHechizo"))
188         .Contadores.InvervaloRespawn = val(Leer.GetValue("NPC" & NpcNumber, "InvervaloRespawn"))
    
190         .InformarRespawn = val(Leer.GetValue("NPC" & NpcNumber, "InformarRespawn"))
    
192         .QuizaProb = val(Leer.GetValue("NPC" & NpcNumber, "QuizaProb"))
    
194         .SubeSupervivencia = val(Leer.GetValue("NPC" & NpcNumber, "SubeSupervivencia"))
    
196         If .IntervaloMovimiento = 0 Then
198             .IntervaloMovimiento = 380
200             .Char.speeding = 0.552631578947368
            Else
202             .Char.speeding = 210 / .IntervaloMovimiento
            End If
    
204         If .InvervaloLanzarHechizo = 0 Then
206             .InvervaloLanzarHechizo = 8000
    
            End If
    
208         If .IntervaloAtaque = 0 Then
210             .IntervaloAtaque = 2000
    
            End If
    
212         .Stats.MaxHp = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
214         .Stats.MinHp = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
216         .Stats.MaxHit = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
218         .Stats.MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
220         .Stats.def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
222         .Stats.defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
224         .Stats.Alineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))
    
226         .Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
    
228         For LoopC = 1 To .Invent.NroItems
230             ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
232             .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
234             .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
236         Next LoopC
    
238         .flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
    
240         If .flags.LanzaSpells > 0 Then
242             ReDim .Spells(1 To .flags.LanzaSpells)
            End If
    
244         For LoopC = 1 To .flags.LanzaSpells
246             .Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
248         Next LoopC
    
250         If .NPCtype = eNPCType.Entrenador Then
        
252             .NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
            
254             ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador
    
256             For LoopC = 1 To .NroCriaturas
258                 .Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
260                 .Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
262             Next LoopC
    
            End If
    
264         .flags.NPCActive = True
266         .flags.NPCActive = True
268         .flags.UseAINow = False
    
270         If Respawn Then
272             .flags.Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
            Else
274             .flags.Respawn = 1
            End If
    
276         .flags.backup = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
278         .flags.RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
280         .flags.AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
282         .flags.GolpeExacto = val(Leer.GetValue("NPC" & NpcNumber, "GolpeExacto"))
    
284         .flags.Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
286         .flags.Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
288         .flags.Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))
    
            
    
    
            '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
290         Aux = Leer.GetValue("NPC" & NpcNumber, "NROEXP")
    
292         If LenB(Aux) = 0 Then
294             .NroExpresiones = 0
            
            Else
        
296             .NroExpresiones = val(Aux)
            
298             ReDim .Expresiones(1 To .NroExpresiones) As String
    
300             For LoopC = 1 To .NroExpresiones
302                 .Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
304             Next LoopC
    
            End If
    
            '<<<<<<<<<<<<<< Sistema de Dropeo NUEVO >>>>>>>>>>>>>>>>
306         Aux = Leer.GetValue("NPC" & NpcNumber, "NumQuiza")
    
308         If LenB(Aux) = 0 Then
310             .NumQuiza = 0
            
            Else
        
312             .NumQuiza = val(Aux)
            
314             ReDim .QuizaDropea(1 To .NumQuiza) As String
    
316             For LoopC = 1 To .NumQuiza
318                 .QuizaDropea(LoopC) = Leer.GetValue("NPC" & NpcNumber, "QuizaDropea" & LoopC)
320             Next LoopC
    
            End If
    
    
        'Ladder
        'Nuevo sistema de Quest
    
322         Aux = Leer.GetValue("NPC" & NpcNumber, "NumQuest")
    
324         If LenB(Aux) = 0 Then
326             .NumQuest = 0
            
            Else
        
328             .NumQuest = val(Aux)
                
330             ReDim .QuestNumber(1 To .NumQuest) As Byte
                
332             For LoopC = 1 To .NumQuest
334                 .QuestNumber(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber" & LoopC))
336             Next LoopC
    
            End If
            
        'Nuevo sistema de Quest
    
        'Nuevo sistema de Drop Quest
            
338         Aux = Leer.GetValue("NPC" & NpcNumber, "NumDropQuest")
    
340         If LenB(Aux) = 0 Then
342             .NumDropQuest = 0
            
            Else
        
344             .NumDropQuest = val(Aux)
                
346             ReDim .DropQuest(1 To .NumDropQuest) As tQuestObj
                
348             For LoopC = 1 To .NumDropQuest
350                 .DropQuest(LoopC).QuestIndex = val(ReadField(1, Leer.GetValue("NPC" & NpcNumber, "DropQuest" & LoopC), Asc("-")))
352                 .DropQuest(LoopC).ObjIndex = val(ReadField(2, Leer.GetValue("NPC" & NpcNumber, "DropQuest" & LoopC), Asc("-")))
354                 .DropQuest(LoopC).Amount = val(ReadField(3, Leer.GetValue("NPC" & NpcNumber, "DropQuest" & LoopC), Asc("-")))
356                 .DropQuest(LoopC).Probabilidad = val(ReadField(4, Leer.GetValue("NPC" & NpcNumber, "DropQuest" & LoopC), Asc("-")))
358             Next LoopC
    
            End If
    
        'Nuevo sistema de Drop Quest
        'Ladder
    
    
    
    
    
    
            '<<<<<<<<<<<<<< Sistema de Viajes NUEVO >>>>>>>>>>>>>>>>
360         Aux = Leer.GetValue("NPC" & NpcNumber, "NumDestinos")
    
362         If LenB(Aux) = 0 Then
364             .NumDestinos = 0
            
            Else
        
366             .NumDestinos = val(Aux)
            
368             ReDim .Dest(1 To .NumDestinos) As String
    
370             For LoopC = 1 To .NumDestinos
372                 .Dest(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Dest" & LoopC)
374             Next LoopC
    
            End If
    
            '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
376         .Interface = val(Leer.GetValue("NPC" & NpcNumber, "Interface"))
    
            'Tipo de items con los que comercia
378         .TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))
    
    
            '<<<<<<<<<<<<<< Animaciones >>>>>>>>>>>>>>>>
    
            ' Por defecto la animación es idle
380         Call AnimacionIdle(NpcIndex, True)
    
            ' Si el tipo de movimiento es Caminata
            If .Movement = Caminata Then
                ' Leemos la cantidad de indicaciones
                Dim Cant As Byte
                Cant = val(Leer.GetValue("NPC" & NpcNumber, "CaminataLen"))
                ' Prevengo NPCs rotos
                If Cant = 0 Then
                    .Movement = ESTATICO
                Else
                    ' Redimenciono el array
                    ReDim .Caminata(1 To Cant)
                    ' Leo todas las indicaciones
                    For LoopC = 1 To Cant
                        Field = Split(Leer.GetValue("NPC" & NpcNumber, "Caminata" & LoopC), ":")
    
                        .Caminata(LoopC).Offset.X = val(Field(0))
                        .Caminata(LoopC).Offset.Y = val(Field(1))
                        .Caminata(LoopC).Espera = val(Field(2))
                    Next
                    
                    .CaminataActual = 1
                End If
            End If
            '<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>
            
        End With

        'Si NO estamos actualizando los NPC's activos, actualizamos el contador.
382     If Reload = False Then
384         If NpcIndex > LastNPC Then LastNPC = NpcIndex
386         NumNPCs = NumNPCs + 1
        End If
    
        'Devuelve el nuevo Indice
388     OpenNPC = NpcIndex

        Exit Function

OpenNPC_Err:
390     Call RegistrarError(Err.Number, Err.Description, "NPCs.OpenNPC", Erl)
392     Resume Next
        
End Function

Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)
        
        On Error GoTo DoFollow_Err
        

100     If Npclist(NpcIndex).flags.Follow Then
102         Npclist(NpcIndex).flags.AttackedBy = vbNullString
104         Npclist(NpcIndex).flags.Follow = False
106         Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
108         Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
        Else
110         Npclist(NpcIndex).flags.AttackedBy = UserName
112         Npclist(NpcIndex).flags.Follow = True
114         Npclist(NpcIndex).Movement = 4 'follow
116         Npclist(NpcIndex).Hostile = 0

        End If

        
        Exit Sub

DoFollow_Err:
118     Call RegistrarError(Err.Number, Err.Description, "NPCs.DoFollow", Erl)
120     Resume Next
        
End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
        
        On Error GoTo FollowAmo_Err
    
        

100     With Npclist(NpcIndex)
102         .flags.Follow = True
104         .Movement = TipoAI.SigueAmo
106         .Target = .MaestroUser
108         .PFINFO.TargetUser = .MaestroUser
110         .PFINFO.PathLenght = 0
112         .Hostile = 0
114         .Target = 0
116         .TargetNPC = 0
        End With

        
        Exit Sub

FollowAmo_Err:
118     Call RegistrarError(Err.Number, Err.Description, "NPCs.FollowAmo", Erl)

        
End Sub

Public Function ObtenerIndiceRespawn() As Integer

        On Error GoTo ErrHandler

        Dim LoopC As Integer

100     For LoopC = 1 To MaxRespawn

            'If LoopC > MaxRespawn Then Exit For
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
    
100     With Npclist(NpcIndex)
    
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
116         Npclist(NpcIndex).Pos = NuevaPos
118         Call MakeNPCChar(True, 0, NpcIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

120         If FX Then                                    'FX
122             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_WARP, NuevaPos.X, NuevaPos.Y))
124             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.CharIndex, FXIDs.FXWARP, 0))
            End If

        End If

End Sub

' Autor: WyroX - 20/01/2021
' Intenta moverlo hacia un "costado" según el heading indicado. Se usa para mover NPCs del camino de otro char.
' Si no hay un lugar válido a los lados, lo mueve a la posición válida más cercana.
Sub MoveNpcToSide(ByVal NpcIndex As Integer, ByVal Heading As eHeading)

    On Error GoTo Handler

    With Npclist(NpcIndex)

        ' Elegimos un lado al azar
        Dim R As Integer
        R = RandomNumber(0, 1) * 2 - 1 ' -1 o 1

        ' Roto el heading original hacia ese lado
        Heading = RotateHeading(Heading, R)

        ' Intento moverlo para ese lado
        If MoveNPCChar(NpcIndex, Heading) Then Exit Sub
        
        ' Si falló, intento moverlo para el lado opuesto
        Heading = InvertHeading(Heading)
        If MoveNPCChar(NpcIndex, Heading) Then Exit Sub
        
        ' Si ambos fallan, entonces lo dejo en la posición válida más cercana
        Dim NuevaPos As WorldPos
        Call ClosestLegalPos(.Pos, NuevaPos, .flags.AguaValida, .flags.TierraInvalida = 0)
        Call WarpNpcChar(NpcIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

    End With

    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.Description, "NPCs.MoveNpcToSide", Erl)
    Resume Next
End Sub
