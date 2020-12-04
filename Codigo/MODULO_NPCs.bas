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
        Call RegistrarError(Err.Number, Err.description, "NPCs.QuitarMascotaNpc", Erl)
        Resume Next
        
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

    Dim MiNPC As npc

    MiNPC = Npclist(NpcIndex)

    Dim EraCriminal As Byte

    Dim TiempoRespw As Integer

    TiempoRespw = Npclist(NpcIndex).Contadores.InvervaloRespawn
    'Familiares
    '  If UserList(UserIndex).Familiar.Existe = 1 Then
    '  If UserList(UserIndex).Familiar.Invocado = 1 Then
    '   If NpcIndex = UserList(UserIndex).Familiar.Id Then
    'Call WriteConsoleMsg(UserIndex, "Tu familiar a muerto, deberas resucitarlo.", FontTypeNames.FONTTYPE_WARNING)
    '   Call WriteLocaleMsg(UserIndex, "181", FontTypeNames.FONTTYPE_WARNING)
    '   UserList(UserIndex).Familiar.Muerto = 1
    ' End If
    ' End If
    ' End If
    'Familiares
    'Quitamos el npc
    Call QuitarNPC(NpcIndex)
    
    If UserIndex > 0 Then ' Lo mato un usuario?
        If MiNPC.flags.Snd3 > 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.Pos.X, MiNPC.Pos.Y))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave("28", MiNPC.Pos.X, MiNPC.Pos.Y))
        
        End If
        
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        
        'If MiNPC.SubeSupervivencia = 1 Then
        Call SubirSkill(UserIndex, eSkill.Supervivencia)
        'End If
        
        '[KEVIN]
        If MiNPC.flags.ExpCount > 0 Then

            If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + MiNPC.flags.ExpCount

                If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
                    
                Call WriteRenderValueMsg(UserIndex, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, MiNPC.flags.ExpCount, 6)
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)

            End If
        
            MiNPC.flags.ExpCount = 0

        End If
        
        '[/KEVIN]
        ' Call WriteConsoleMsg(UserIndex, "Has matado a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
        If UserList(UserIndex).ChatCombate = 1 Then
            Call WriteLocaleMsg(UserIndex, "184", FontTypeNames.FONTTYPE_FIGHT, "la criatura")

        End If
        
        'Particula al matar
        ' Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(MiNPC.Pos.X, MiNPC.Pos.Y, 84, 2))
        
        'Call WriteEfectOverHead(SendTarget.ToPCArea, MiNPC.GiveGLD, CStr(Npclist(NpcIndex).Char.CharIndex))
        ' Call WriteConsoleMsg(UserIndex, MiNPC.GiveGLD, FontTypeNames.FONTTYPE_FIGHT)
        
        If UserList(UserIndex).Stats.NPCsMuertos < 32000 Then UserList(UserIndex).Stats.NPCsMuertos = UserList(UserIndex).Stats.NPCsMuertos + 1
        ' Call CheckearRecompesas(UserIndex, 1)
        
        EraCriminal = Status(UserIndex)
        
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
        
        If MiNPC.GiveEXPClan > 0 Then
            If UserList(UserIndex).GuildIndex > 0 Then
                Call modGuilds.CheckClanExp(UserIndex, MiNPC.GiveEXPClan)

                ' Else
                ' Call WriteConsoleMsg(UserIndex, "No perteneces a ningun clan, experiencia perdida.", FontTypeNames.FONTTYPE_INFOIAO)
            End If

        End If
        
        Dim i As Long, j As Long
        
        For i = 1 To MAXUSERQUESTS
        
            With UserList(UserIndex).QuestStats.Quests(i)
        
                If .QuestIndex Then
                    If QuestList(.QuestIndex).RequiredNPCs Then
        
                        For j = 1 To QuestList(.QuestIndex).RequiredNPCs
        
                            If QuestList(.QuestIndex).RequiredNPC(j).NpcIndex = MiNPC.Numero Then
                                If QuestList(.QuestIndex).RequiredNPC(j).Amount > .NPCsKilled(j) Then
                                    .NPCsKilled(j) = .NPCsKilled(j) + 1
        
                                End If
                                    
                                If QuestList(.QuestIndex).RequiredNPC(j).Amount = .NPCsKilled(j) Then
                                    Call WriteConsoleMsg(UserIndex, "Ya has matado todos los " & MiNPC.name & " que la mision " & QuestList(.QuestIndex).nombre & " requeria. Chequeá si ya estas listo para recibir la recompensa.", FontTypeNames.FONTTYPE_INFOIAO)
                                    
                                End If
        
                            End If
        
                        Next j
        
                    End If
        
                End If
        
            End With
        
        Next i
            
    End If ' UserIndex > 0
    
    If Npclist(NpcIndex).MaestroUser > 0 Then Exit Sub

    'Tiramos el oro
    Call NPCTirarOro(MiNPC, UserIndex)
    
    
    Call DropObjQuest(MiNPC, UserIndex)

    'Item Magico!
    Call NpcDropeo(MiNPC, UserIndex)
        
    'Tiramos el inventario
    Call NPC_TIRAR_ITEMS(MiNPC)
    'ReSpawn o no

    If TiempoRespw = 0 Then
        Call ReSpawnNpc(MiNPC)
    Else

        Dim Indice As Integer

        MiNPC.flags.NPCActive = True
        Indice = ObtenerIndiceRespawn
        RespawnList(Indice) = MiNPC

    End If
    
    Exit Sub

ErrHandler:
    Call LogError("Error en MuereNpc - Error: " & Err.Number & " - Desc: " & Err.description)

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
        End With

        
        Exit Sub

ResetNpcFlags_Err:
        Call RegistrarError(Err.Number, Err.description, "NPCs.ResetNpcFlags", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "NPCs.ResetNpcCounters", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "NPCs.ResetNpcCharInfo", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "NPCs.ResetNpcCriatures", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "NPCs.ResetExpresiones", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "NPCs.ResetDrop", Erl)
        Resume Next
        
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
    
118     If Npclist(NpcIndex).MaestroNPC > 0 Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNPC)
    
120     Npclist(NpcIndex).MaestroNPC = 0
    
122     Npclist(NpcIndex).Mascotas = 0
124     Npclist(NpcIndex).Movement = 0
126     Npclist(NpcIndex).name = "NPC SIN INICIAR"
128     Npclist(NpcIndex).NPCtype = 0
130     Npclist(NpcIndex).Numero = 0
132     Npclist(NpcIndex).Orig.Map = 0
134     Npclist(NpcIndex).Orig.X = 0
136     Npclist(NpcIndex).Orig.Y = 0
138     Npclist(NpcIndex).PoderAtaque = 0
140     Npclist(NpcIndex).PoderEvasion = 0
142     Npclist(NpcIndex).Pos.Map = 0
144     Npclist(NpcIndex).Pos.X = 0
146     Npclist(NpcIndex).Pos.Y = 0
148     Npclist(NpcIndex).Target = 0
150     Npclist(NpcIndex).TargetNPC = 0
152     Npclist(NpcIndex).TipoItems = 0
154     Npclist(NpcIndex).Veneno = 0
156     Npclist(NpcIndex).Desc = vbNullString
        
        If Npclist(NpcIndex).MaestroUser > 0 Then Call QuitarMascota(Npclist(NpcIndex).MaestroUser, NpcIndex)
        If Npclist(NpcIndex).MaestroNPC > 0 Then Call QuitarMascotaNpc(Npclist(NpcIndex).MaestroNPC)

        Npclist(NpcIndex).MaestroUser = 0
        Npclist(NpcIndex).MaestroNPC = 0
    
        Dim j As Integer

158     For j = 1 To Npclist(NpcIndex).NroSpells
160         Npclist(NpcIndex).Spells(j) = 0
162     Next j
    
164     Call ResetNpcCharInfo(NpcIndex)
166     Call ResetNpcCriatures(NpcIndex)
168     Call ResetExpresiones(NpcIndex)
170     Call ResetDrop(NpcIndex)

        
        Exit Sub

ResetNpcMainInfo_Err:
        Call RegistrarError(Err.Number, Err.description, "NPCs.ResetNpcMainInfo", Erl)
        Resume Next
        
End Sub

Sub QuitarNPC(ByVal NpcIndex As Integer)

    On Error GoTo ErrHandler

    Npclist(NpcIndex).flags.NPCActive = False
    
    If InMapBounds(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y) Then
        Call EraseNPCChar(NpcIndex)
    End If
    
    ' Es pretoriano?
    If Npclist(NpcIndex).NPCtype = eNPCType.Pretoriano Then
        Call ClanPretoriano(Npclist(NpcIndex).ClanIndex).MuerePretoriano(NpcIndex)
    End If
    
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    
    Call ResetNpcMainInfo(NpcIndex)
    
    If NpcIndex = LastNPC Then

        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1

            If LastNPC < 1 Then Exit Do
        Loop

    End If
      
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1

    End If

    Exit Sub

ErrHandler:
    Npclist(NpcIndex).flags.NPCActive = False
    Call LogError("Error en QuitarNPC")

End Sub

Function TestSpawnTrigger(Pos As WorldPos, Optional PuedeAgua As Boolean = False) As Boolean
        
        On Error GoTo TestSpawnTrigger_Err
        
    
100     If LegalPos(Pos.Map, Pos.X, Pos.Y, PuedeAgua) Then
102         TestSpawnTrigger = MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 3 And MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 2 And MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 1

        End If

        
        Exit Function

TestSpawnTrigger_Err:
        Call RegistrarError(Err.Number, Err.description, "NPCs.TestSpawnTrigger", Erl)
        Resume Next
        
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

106     PuedeAgua = Npclist(nIndex).flags.AguaValida
108     PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)
    
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
216     Call RegistrarError(Err.Number, Err.description, "NPCs.CrearNPC", Erl)
218     Resume Next
        
End Function

Sub MakeNPCChar(ByVal toMap As Boolean, sndIndex As Integer, NpcIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo MakeNPCChar_Err
        

        Dim CharIndex As Integer

100     If Npclist(NpcIndex).Char.CharIndex = 0 Then
102         CharIndex = NextOpenCharIndex
104         Npclist(NpcIndex).Char.CharIndex = CharIndex
106         CharList(CharIndex) = NpcIndex

        End If
    
108     MapData(Map, X, Y).NpcIndex = NpcIndex
    
        Dim Simbolo As Byte
    
        Dim GG      As String
        Dim tmpByte As Byte
   
110     GG = IIf(Npclist(NpcIndex).showName > 0, Npclist(NpcIndex).name & Npclist(NpcIndex).SubName, vbNullString)
    
112     If Not toMap Then
114         If Npclist(NpcIndex).NumQuest > 0 Then

            Dim q As Byte
            Dim HayFinalizada As Boolean
            Dim HayDisponible As Boolean
            Dim HayPendiente As Boolean
    
                For q = 1 To Npclist(NpcIndex).NumQuest
                     tmpByte = TieneQuest(sndIndex, Npclist(NpcIndex).QuestNumber(q))
                
                     If tmpByte Then
                            If FinishQuestCheck(sndIndex, Npclist(NpcIndex).QuestNumber(q), tmpByte) Then
                                Simbolo = 3
                                HayFinalizada = True
                            Else
                                HayPendiente = True
                                Simbolo = 4
                            End If
                        Else
116                         If UserDoneQuest(sndIndex, Npclist(NpcIndex).QuestNumber(q)) Or UserList(sndIndex).Stats.ELV < QuestList(Npclist(NpcIndex).QuestNumber(q)).RequiredLevel Then
118                             Simbolo = 2
                            Else
120                             Simbolo = 1
                                HayDisponible = True
                            End If
                    End If
            
                
                Next q
                
                
                'Para darle prioridad a ciertos simbolos
                
                If HayDisponible Then
                    Simbolo = 1
                End If
                
                If HayPendiente Then
                    Simbolo = 4
                End If
                
                If HayFinalizada Then
                    Simbolo = 3
                End If
                'Para darle prioridad a ciertos simbolos
                
            End If
            
            

         Call WriteCharacterCreate(sndIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, Npclist(NpcIndex).Char.Heading, Npclist(NpcIndex).Char.CharIndex, X, Y, Npclist(NpcIndex).Char.WeaponAnim, Npclist(NpcIndex).Char.ShieldAnim, 0, 0, Npclist(NpcIndex).Char.CascoAnim, GG, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1#, True, False, 0, 0, 0, 0, Npclist(NpcIndex).Stats.MinHp, Npclist(NpcIndex).Stats.MaxHp, Simbolo)
        
        Else
124         Call AgregarNpc(NpcIndex)

        End If

        
        Exit Sub

MakeNPCChar_Err:
        Call RegistrarError(Err.Number, Err.description, "NPCs.MakeNPCChar", Erl)
        Resume Next
        
End Sub

Sub ChangeNPCChar(ByVal NpcIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading)
        
        On Error GoTo ChangeNPCChar_Err
        

100     If NpcIndex > 0 Then
102         Npclist(NpcIndex).Char.Body = Body
104         Npclist(NpcIndex).Char.Head = Head
106         Npclist(NpcIndex).Char.Heading = Heading
        
108         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(Body, Head, Heading, Npclist(NpcIndex).Char.CharIndex, 0, 0, 0, 0, 0))

        End If

        
        Exit Sub

ChangeNPCChar_Err:
        Call RegistrarError(Err.Number, Err.description, "NPCs.ChangeNPCChar", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "NPCs.EraseNPCChar", Erl)
        Resume Next
        
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
    
    With Npclist(NpcIndex)
        nPos = .Pos
        Call HeadtoPos(nHeading, nPos)
        
        ' es una posicion legal
        If LegalWalkNPC(nPos.Map, nPos.X, nPos.Y, nHeading, .flags.AguaValida = 1, .flags.TierraInvalida = 0, .MaestroUser <> 0) Then
            
            UserIndex = MapData(.Pos.Map, nPos.X, nPos.Y).UserIndex

            ' Si hay un usuario a donde se mueve el npc, entonces esta muerto o es un gm invisible
            If UserIndex > 0 Then

                With UserList(UserIndex)
                
                    ' Actualizamos posicion y mapa
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
                    .Pos.X = Npclist(NpcIndex).Pos.X
                    .Pos.Y = Npclist(NpcIndex).Pos.Y
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                        
                    ' Avisamos a los usuarios del area, y al propio usuario lo forzamos a moverse
                    Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, .Pos.X, .Pos.Y))
                    Call WriteForceCharMove(UserIndex, InvertHeading(nHeading))

                End With

            End If
            
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))

            'Update map and user pos
            MapData(.Pos.Map, .Pos.X, .Pos.Y).NpcIndex = 0
            .Pos = nPos
            .Char.Heading = nHeading
            MapData(.Pos.Map, nPos.X, nPos.Y).NpcIndex = NpcIndex
            
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
            ' Npc has moved
            MoveNPCChar = True
        
        ElseIf .MaestroUser = 0 Then

            If .Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                .PFINFO.PathLenght = 0
            End If

        End If

    End With
    
    Exit Function

errh:
    LogError ("Error en move npc " & NpcIndex & ". Error: " & Err.Number & " - " & Err.description)

End Function

Function NextOpenNPC() As Integer
    'Call LogTarea("Sub NextOpenNPC")

    On Error GoTo ErrHandler

    Dim LoopC As Integer
  
    For LoopC = 1 To MAXNPCS + 1

        If LoopC > MAXNPCS Then Exit For
        If Not Npclist(LoopC).flags.NPCActive Then Exit For
    Next LoopC
  
    NextOpenNPC = LoopC

    Exit Function
ErrHandler:
    Call LogError("Error en NextOpenNPC")

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
        Call RegistrarError(Err.Number, Err.description, "NPCs.NpcEnvenenarUser", Erl)
        Resume Next
        
End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean, Optional Avisar As Boolean = False) As Integer
        
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

106     PuedeAgua = Npclist(nIndex).flags.AguaValida
108     PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)

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
        Call RegistrarError(Err.Number, Err.description, "NPCs.SpawnNpc", Erl)
        Resume Next
        
End Function

Sub ReSpawnNpc(MiNPC As npc)
        
        On Error GoTo ReSpawnNpc_Err
        

100     If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.Map, MiNPC.Orig)

        
        Exit Sub

ReSpawnNpc_Err:
        Call RegistrarError(Err.Number, Err.description, "NPCs.ReSpawnNpc", Erl)
        Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "NPCs.NPCHostiles", Erl)
        Resume Next
        
End Function

Sub NPCTirarOro(MiNPC As npc, ByVal UserIndex As Integer)
        
        On Error GoTo NPCTirarOro_Err

        If MiNPC.GiveGLD > 0 Then
            Dim Oro As Long
            Oro = MiNPC.GiveGLD * OroMult * UserList(UserIndex).flags.ScrollOro
        
            If UserList(UserIndex).Grupo.EnGrupo Then
                Select Case UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
                    Case 2: Oro = Oro * 1.2
                    Case 3: Oro = Oro * 1.4
                    Case 4: Oro = Oro * 1.6
                    Case 5: Oro = Oro * 1.8
                    Case 6: Oro = Oro * 2
                End Select
            End If

            Dim MiObj As obj
            MiObj.ObjIndex = iORO

            While (Oro > 0)
                If Oro > MAX_INVENTORY_OBJS Then
                    MiObj.Amount = MAX_INVENTORY_OBJS
                    Oro = Oro - MAX_INVENTORY_OBJS
                Else
                    MiObj.Amount = Oro
                    Oro = 0
                End If

                Call TirarItemAlPiso(MiNPC.Pos, MiObj)
            Wend

            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso("87", MiNPC.Pos.X, MiNPC.Pos.Y))
        End If

        
        Exit Sub

NPCTirarOro_Err:
        Call RegistrarError(Err.Number, Err.description, "NPCs.NPCTirarOro", Erl)
        Resume Next
        
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

112     Npclist(NpcIndex).Numero = NpcNumber
114     Npclist(NpcIndex).name = Leer.GetValue("NPC" & NpcNumber, "Name")
116     Npclist(NpcIndex).SubName = Leer.GetValue("NPC" & NpcNumber, "SubName")
118     Npclist(NpcIndex).Desc = Leer.GetValue("NPC" & NpcNumber, "Desc")

120     Npclist(NpcIndex).Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
122     Npclist(NpcIndex).flags.OldMovement = Npclist(NpcIndex).Movement

124     Npclist(NpcIndex).flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
126     Npclist(NpcIndex).flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
128     Npclist(NpcIndex).flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))

130     Npclist(NpcIndex).NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))

132     Npclist(NpcIndex).Char.Body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
134     Npclist(NpcIndex).Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
136     Npclist(NpcIndex).Char.Heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))

138     Npclist(NpcIndex).Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "Arma"))
140     Npclist(NpcIndex).Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "Escudo"))
142     Npclist(NpcIndex).Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "Casco"))

144     Npclist(NpcIndex).Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
146     Npclist(NpcIndex).Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
148     Npclist(NpcIndex).Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
150     Npclist(NpcIndex).flags.OldHostil = Npclist(NpcIndex).Hostile

152     Npclist(NpcIndex).GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP"))

154     Npclist(NpcIndex).Distancia = val(Leer.GetValue("NPC" & NpcNumber, "Distancia"))

156     Npclist(NpcIndex).GiveEXPClan = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXPClan"))

        'Npclist(NpcIndex).flags.ExpDada = Npclist(NpcIndex).GiveEXP
158     Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).GiveEXP

160     Npclist(NpcIndex).Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))

162     Npclist(NpcIndex).flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))

164     Npclist(NpcIndex).GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD"))

'166     Npclist(NpcIndex).QuestNumber = val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber"))




168     Npclist(NpcIndex).PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
170     Npclist(NpcIndex).PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))

172     Npclist(NpcIndex).InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))

174     Npclist(NpcIndex).showName = val(GetVar(DatPath & "NPCs.dat", "NPC" & NpcNumber, "ShowName"))
    
176     Npclist(NpcIndex).GobernadorDe = val(GetVar(DatPath & "NPCs.dat", "NPC" & NpcNumber, "GobernadorDe"))

178     Npclist(NpcIndex).SoundOpen = val(GetVar(DatPath & "NPCs.dat", "NPC" & NpcNumber, "SoundOpen"))
180     Npclist(NpcIndex).SoundClose = val(GetVar(DatPath & "NPCs.dat", "NPC" & NpcNumber, "SoundClose"))

182     Npclist(NpcIndex).IntervaloAtaque = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloAtaque"))
184     Npclist(NpcIndex).IntervaloMovimiento = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloMovimiento"))
186     Npclist(NpcIndex).InvervaloLanzarHechizo = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloLanzarHechizo"))
188     Npclist(NpcIndex).Contadores.InvervaloRespawn = val(Leer.GetValue("NPC" & NpcNumber, "InvervaloRespawn"))

190     Npclist(NpcIndex).InformarRespawn = val(Leer.GetValue("NPC" & NpcNumber, "InformarRespawn"))

192     Npclist(NpcIndex).QuizaProb = val(Leer.GetValue("NPC" & NpcNumber, "QuizaProb"))

194     Npclist(NpcIndex).SubeSupervivencia = val(Leer.GetValue("NPC" & NpcNumber, "SubeSupervivencia"))

196     If Npclist(NpcIndex).IntervaloMovimiento = 0 Then
198         Npclist(NpcIndex).IntervaloMovimiento = 380

        End If

200     If Npclist(NpcIndex).InvervaloLanzarHechizo = 0 Then
202         Npclist(NpcIndex).InvervaloLanzarHechizo = 8000

        End If

204     If Npclist(NpcIndex).IntervaloAtaque = 0 Then
206         Npclist(NpcIndex).IntervaloAtaque = 2000

        End If

208     Npclist(NpcIndex).Stats.MaxHp = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
210     Npclist(NpcIndex).Stats.MinHp = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
212     Npclist(NpcIndex).Stats.MaxHit = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
214     Npclist(NpcIndex).Stats.MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
216     Npclist(NpcIndex).Stats.def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
218     Npclist(NpcIndex).Stats.defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
220     Npclist(NpcIndex).Stats.Alineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))

        Dim LoopC As Long
        Dim ln    As String

222     Npclist(NpcIndex).Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))

224     For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
226         ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
228         Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
230         Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
232     Next LoopC

234     Npclist(NpcIndex).flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))

236     If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
238         ReDim Npclist(NpcIndex).Spells(1 To Npclist(NpcIndex).flags.LanzaSpells)
        End If

240     For LoopC = 1 To Npclist(NpcIndex).flags.LanzaSpells
242         Npclist(NpcIndex).Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
244     Next LoopC

246     If Npclist(NpcIndex).NPCtype = eNPCType.Entrenador Then
    
248         Npclist(NpcIndex).NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
        
250         ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador

252         For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
254             Npclist(NpcIndex).Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
256             Npclist(NpcIndex).Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
258         Next LoopC

        End If

260     Npclist(NpcIndex).flags.NPCActive = True
262     Npclist(NpcIndex).flags.NPCActive = True
264     Npclist(NpcIndex).flags.UseAINow = False

266     If Respawn Then
268         Npclist(NpcIndex).flags.Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
        Else
270         Npclist(NpcIndex).flags.Respawn = 1
        End If

272     Npclist(NpcIndex).flags.backup = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
274     Npclist(NpcIndex).flags.RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
276     Npclist(NpcIndex).flags.AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
278     Npclist(NpcIndex).flags.GolpeExacto = val(Leer.GetValue("NPC" & NpcNumber, "GolpeExacto"))

280     Npclist(NpcIndex).flags.Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
282     Npclist(NpcIndex).flags.Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
284     Npclist(NpcIndex).flags.Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))

        


        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        Dim aux As String
286         aux = Leer.GetValue("NPC" & NpcNumber, "NROEXP")

288     If LenB(aux) = 0 Then
290         Npclist(NpcIndex).NroExpresiones = 0
        
        Else
    
292         Npclist(NpcIndex).NroExpresiones = val(aux)
        
294         ReDim Npclist(NpcIndex).Expresiones(1 To Npclist(NpcIndex).NroExpresiones) As String

296         For LoopC = 1 To Npclist(NpcIndex).NroExpresiones
298             Npclist(NpcIndex).Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
300         Next LoopC

        End If

        '<<<<<<<<<<<<<< Sistema de Dropeo NUEVO >>>>>>>>>>>>>>>>
302     aux = Leer.GetValue("NPC" & NpcNumber, "NumQuiza")

304     If LenB(aux) = 0 Then
306         Npclist(NpcIndex).NumQuiza = 0
        
        Else
    
308         Npclist(NpcIndex).NumQuiza = val(aux)
        
310         ReDim Npclist(NpcIndex).QuizaDropea(1 To Npclist(NpcIndex).NumQuiza) As String

312         For LoopC = 1 To Npclist(NpcIndex).NumQuiza
314             Npclist(NpcIndex).QuizaDropea(LoopC) = Leer.GetValue("NPC" & NpcNumber, "QuizaDropea" & LoopC)
316         Next LoopC

        End If


'Ladder
'Nuevo sistema de Quest

        aux = Leer.GetValue("NPC" & NpcNumber, "NumQuest")

        If LenB(aux) = 0 Then
            Npclist(NpcIndex).NumQuest = 0
        
        Else
    
            Npclist(NpcIndex).NumQuest = val(aux)
            
            ReDim Npclist(NpcIndex).QuestNumber(1 To Npclist(NpcIndex).NumQuest) As Byte
            
            For LoopC = 1 To Npclist(NpcIndex).NumQuest
                Npclist(NpcIndex).QuestNumber(LoopC) = Leer.GetValue("NPC" & NpcNumber, "QuestNumber" & LoopC)
            Next LoopC

        End If
        
        
        
        Npclist(NpcIndex).DropQuest = Leer.GetValue("NPC" & NpcNumber, "DropQuest")


'Ladder
'Nuevo sistema de Quest





        '<<<<<<<<<<<<<< Sistema de Viajes NUEVO >>>>>>>>>>>>>>>>
318     aux = Leer.GetValue("NPC" & NpcNumber, "NumDestinos")

320     If LenB(aux) = 0 Then
322         Npclist(NpcIndex).NumDestinos = 0
        
        Else
    
324         Npclist(NpcIndex).NumDestinos = val(aux)
        
326         ReDim Npclist(NpcIndex).Dest(1 To Npclist(NpcIndex).NumDestinos) As String

328         For LoopC = 1 To Npclist(NpcIndex).NumDestinos
330             Npclist(NpcIndex).Dest(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Dest" & LoopC)
332         Next LoopC

        End If

        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
334     Npclist(NpcIndex).Interface = val(Leer.GetValue("NPC" & NpcNumber, "Interface"))

        'Tipo de items con los que comercia
336     Npclist(NpcIndex).TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))

        'Si NO estamos actualizando los NPC's activos, actualizamos el contador.
338     If Reload = False Then
340         If NpcIndex > LastNPC Then LastNPC = NpcIndex
342         NumNPCs = NumNPCs + 1
        End If
    
        'Devuelve el nuevo Indice
344     OpenNPC = NpcIndex

        Exit Function

OpenNPC_Err:
346     Call RegistrarError(Err.Number, Err.description, "NPCs.OpenNPC", Erl)
348     Resume Next
        
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
        Call RegistrarError(Err.Number, Err.description, "NPCs.DoFollow", Erl)
        Resume Next
        
End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With Npclist(NpcIndex)
        .flags.Follow = True
        .Movement = TipoAI.SigueAmo
        .Target = .MaestroUser
        .PFINFO.TargetUser = .MaestroUser
        .PFINFO.PathLenght = 0
        .Hostile = 0
        .Target = 0
        .TargetNPC = 0
    End With

End Sub

Public Function ObtenerIndiceRespawn() As Integer

    On Error GoTo ErrHandler

    Dim LoopC As Integer

    For LoopC = 1 To MaxRespawn

        'If LoopC > MaxRespawn Then Exit For
        If Not RespawnList(LoopC).flags.NPCActive Then Exit For
    Next LoopC
  
    ObtenerIndiceRespawn = LoopC

    Exit Function
ErrHandler:
    Call LogError("Error en ObtenerIndiceRespawn")
    
End Function

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer
    
    For i = 1 To MAXMASCOTAS

        If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
            UserList(UserIndex).MascotasIndex(i) = 0
            UserList(UserIndex).MascotasType(i) = 0
         
            UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
            Exit For

        End If

    Next i

End Sub

