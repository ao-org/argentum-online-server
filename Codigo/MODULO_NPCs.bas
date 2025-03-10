Attribute VB_Name = "NPCs"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Public Const MaxRespawn             As Integer = 255
Public Const NpcIndexHeapSize       As Integer = 10000

Public RespawnList(1 To MaxRespawn) As t_Npc
Private IdNpcLibres                 As t_IndexHeap
Option Explicit

Public Sub InitializeNpcIndexHeap(Optional ByVal Size As Integer = NpcIndexHeapSize)

    On Error GoTo ErrHandler_InitizlizeNpcIndex

    ReDim IdNpcLibres.IndexInfo(Size)

    Dim i As Integer

    For i = 1 To Size
        IdNpcLibres.IndexInfo(i) = Size - (i - 1)
    Next i

    IdNpcLibres.currentIndex = Size
    Exit Sub
ErrHandler_InitizlizeNpcIndex:
    Call TraceError(Err.Number, Err.Description, "NPCs.InitializeNpcIndexHeap", Erl)

End Sub

Public Function IsValidNpcRef(ByRef NpcRef As t_NpcReference) As Boolean
    IsValidNpcRef = False

    If NpcRef.ArrayIndex < LBound(NpcList) Or NpcRef.ArrayIndex > UBound(NpcList) Then
        Exit Function

    End If

    If NpcList(NpcRef.ArrayIndex).VersionId <> NpcRef.VersionId Then
        Exit Function

    End If

    IsValidNpcRef = True

End Function

Public Function SetNpcRef(ByRef NpcRef As t_NpcReference, _
                          ByVal Index As Integer) As Boolean
    SetNpcRef = False
    NpcRef.ArrayIndex = Index

    If Index < LBound(NpcList) Or NpcRef.ArrayIndex > UBound(NpcList) Then
        Exit Function

    End If

    NpcRef.VersionId = NpcList(Index).VersionId
    SetNpcRef = True

End Function

Public Sub ClearNpcRef(ByRef NpcRef As t_NpcReference)
    NpcRef.ArrayIndex = 0
    NpcRef.VersionId = -1

End Sub

Public Sub IncreaseNpcVersionId(ByVal NpcIndex As Integer)

    With NpcList(NpcIndex)

        If .VersionId > 32760 Then
            .VersionId = 0
        Else
            .VersionId = .VersionId + 1

        End If

    End With

End Sub

Public Function ReleaseNpc(ByVal NpcIndex As Integer, _
                           ByVal Reason As e_DeleteSource) As Boolean

    On Error GoTo ErrHandler

    If Not NpcList(NpcIndex).flags.NPCActive Then
        Call TraceError(Err.Number, "Trying to release the id twice, last reset reason: " & NpcList(NpcIndex).LastReset & " current reason " & Reason, "NPCs.ReleaseNpc", Erl)
        ReleaseNpc = False
        Exit Function

    End If

    NpcList(NpcIndex).flags.NPCActive = False
    NpcList(NpcIndex).LastReset = Reason
    Call IncreaseNpcVersionId(NpcIndex)
    IdNpcLibres.currentIndex = IdNpcLibres.currentIndex + 1
    Debug.Assert IdNpcLibres.currentIndex <= NpcIndexHeapSize
    IdNpcLibres.IndexInfo(IdNpcLibres.currentIndex) = NpcIndex
    ReleaseNpc = True
    Exit Function
ErrHandler:
    ReleaseNpc = False
    Call TraceError(Err.Number, Err.Description, "NPCs.ReleaseNpc", Erl)

End Function

Public Function GetAvailableNpcIndex() As Integer
    GetAvailableNpcIndex = IdNpcLibres.currentIndex

End Function

Public Function GetNextAvailableNpc() As Integer

    On Error GoTo ErrHandler

    If (IdNpcLibres.currentIndex = 0) Then
        GetNextAvailableNpc = 0
        Return

    End If

    GetNextAvailableNpc = IdNpcLibres.IndexInfo(IdNpcLibres.currentIndex)
    IdNpcLibres.currentIndex = IdNpcLibres.currentIndex - 1

    If NpcList(GetNextAvailableNpc).flags.NPCActive Then
        Call TraceError(Err.Number, "Trying to active the same id twice", "NPCs.GetNextAvailableNpc", Erl)

    End If

    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "NPCs.GetNextAvailableNpc", Erl)

End Function

Sub QuitarMascotaNpc(ByVal Maestro As Integer)

        On Error GoTo QuitarMascotaNpc_Err

100     NpcList(Maestro).Mascotas = NpcList(Maestro).Mascotas - 1
        Exit Sub
QuitarMascotaNpc_Err:
102     Call TraceError(Err.Number, Err.Description, "NPCs.QuitarMascotaNpc", Erl)

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

        Dim MiNPC       As t_Npc
        Dim EraCriminal As Byte
        Dim TiempoRespw As Long
        Dim i           As Long, j As Long
        Dim Indice      As Integer

        ' Objetivo de pruebas nunca muere
100     If NpcList(NpcIndex).npcType = DummyTarget Then
102         Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageChatOverHead("¡¡Auch!!", NpcList(NpcIndex).Char.charindex, vbRed))

104         If UBound(NpcList(NpcIndex).Char.Animation) > 0 Then
106             Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageDoAnimation(NpcList(NpcIndex).Char.charindex, NpcList(NpcIndex).Char.Animation(1)))

            End If

108         NpcList(NpcIndex).Stats.MinHp = NpcList(NpcIndex).Stats.MaxHp
            NpcList(NpcIndex).Stats.shield = 0
            Exit Sub

        End If

110     MiNPC = NpcList(NpcIndex)
112     TiempoRespw = NpcList(NpcIndex).Contadores.IntervaloRespawn

        ' Es NPC de la invasión?
118     If MiNPC.flags.InvasionIndex Then
120         Call MuereNpcInvasion(MiNPC.flags.InvasionIndex, MiNPC.flags.IndexInInvasion)

        End If

        'Quitamos el npc
122     Call QuitarNPC(NpcIndex, eDie)

124     If UserIndex > 0 Then ' Lo mato un usuario?
126         If MiNPC.flags.Snd3 > 0 Then
128             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.pos.x, MiNPC.pos.y))
            Else
130             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave("28", MiNPC.pos.x, MiNPC.pos.y))

            End If

132         Call ClearNpcRef(UserList(UserIndex).flags.TargetNPC)
134         UserList(UserIndex).flags.TargetNpcTipo = e_NPCType.Comun

            ' El user que lo mato tiene mascotas?
            If UserList(UserIndex).NroMascotas > 0 Then

                ' Me fijo si alguna de sus mascotas le estaba pegando al NPC
                For i = 1 To UBound(UserList(UserIndex).MascotasIndex)

                    If UserList(UserIndex).MascotasIndex(i).ArrayIndex > 0 Then
                        If IsValidNpcRef(UserList(UserIndex).MascotasIndex(i)) Then
135                         If NpcList(UserList(UserIndex).MascotasIndex(i).ArrayIndex).TargetNPC.ArrayIndex = NpcIndex Then
136                             Call FollowAmo(UserList(UserIndex).MascotasIndex(i).ArrayIndex)

                            End If

                        Else
                            Call ClearNpcRef(UserList(UserIndex).MascotasIndex(i))

                        End If

                    End If

                Next

            End If

138         If UserList(UserIndex).ChatCombate = 1 Then
140             Call WriteLocaleMsg(UserIndex, "184", e_FontTypeNames.FONTTYPE_DIOS, "la criatura")

            End If

142         If UserList(UserIndex).Stats.NPCsMuertos < 32000 Then UserList(UserIndex).Stats.NPCsMuertos = UserList(UserIndex).Stats.NPCsMuertos + 1
144         If IsValidUserRef(MiNPC.MaestroUser) Then Exit Sub
146         Call SubirSkill(UserIndex, e_Skill.Supervivencia)

148         If MiNPC.flags.ExpCount > 0 Then
150             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
152                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + MiNPC.flags.ExpCount

154                 If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
156                 Call WriteTextOverTile(UserIndex, "+" & PonerPuntos(MiNPC.flags.ExpCount), UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, RGB(0, 169, 255))
158                 Call WriteUpdateExp(UserIndex)
160                 Call CheckUserLevel(UserIndex)

                End If

162             MiNPC.flags.ExpCount = 0

            End If

164         EraCriminal = Status(UserIndex)

166         If MiNPC.GiveEXPClan > 0 Then
168             If UserList(UserIndex).GuildIndex > 0 Then
170                 Call modGuilds.CheckClanExp(UserIndex, MiNPC.GiveEXPClan * SvrConfig.GetValue("ExpMult"))

                End If

            End If

172         For i = 1 To MAXUSERQUESTS

174             With UserList(UserIndex).QuestStats.Quests(i)

176                 If .QuestIndex Then
178                     If QuestList(.QuestIndex).RequiredNPCs Then

180                         For j = 1 To QuestList(.QuestIndex).RequiredNPCs

182                             If QuestList(.QuestIndex).RequiredNPC(j).NpcIndex = MiNPC.Numero Then
184                                 If QuestList(.QuestIndex).RequiredNPC(j).amount >= .NPCsKilled(j) Then
186                                     .NPCsKilled(j) = .NPCsKilled(j) + 1 '
188                                     Call WriteConsoleMsg(UserIndex, MiNPC.name & " matados/as: " & .NPCsKilled(j) & " de " & QuestList(.QuestIndex).RequiredNPC(j).amount, e_FontTypeNames.FONTTYPE_INFOIAO)
190                                     Call WriteChatOverHead(UserIndex, "NOCONSOLA*" & .NPCsKilled(j) & "/" & QuestList(.QuestIndex).RequiredNPC(j).amount & " " & MiNPC.name, UserList(UserIndex).Char.charindex, RGB(180, 180, 180))
                                    Else
192                                     Call WriteConsoleMsg(UserIndex, "Ya has matado todos los " & MiNPC.name & " que la misión " & QuestList(.QuestIndex).nombre & " requería. Revisa si ya estás listo para recibir la recompensa.", e_FontTypeNames.FONTTYPE_INFOIAO)
194                                     Call WriteChatOverHead(UserIndex, "NOCONSOLA*" & QuestList(.QuestIndex).RequiredNPC(j).amount & "/" & QuestList(.QuestIndex).RequiredNPC(j).amount & " " & MiNPC.name, UserList(UserIndex).Char.charindex, RGB(180, 180, 180))

                                    End If

                                End If

196                         Next j

                        End If

                        UserList(UserIndex).flags.ModificoQuests = True

                    End If

                End With

198         Next i

            'Tiramos el oro
200         Call NPCTirarOro(MiNPC, UserIndex)
202         Call DropObjQuest(MiNPC, UserIndex)
            'Item Magico!
204         Call NpcDropeo(MiNPC, UserIndex)
            Call DropFromGlobalDropTable(MiNPC, UserIndex)
            'Tiramos el inventario
206         Call NPC_TIRAR_ITEMS(MiNPC)
        End If ' UserIndex > 0

        ' Mascotas y npcs de entrenamiento no respawnean
208     If MiNPC.MaestroNPC.ArrayIndex > 0 Or IsValidUserRef(MiNPC.MaestroUser) Then Exit Sub
210     If NpcIndex = npc_index_evento Then
212         BusquedaNpcActiva = False
214         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Evento> El NPC ha sido asesinado.", e_FontTypeNames.FONTTYPE_CITIZEN))
            npc_index_evento = 0

        End If

        'ReSpawn o no
216     If TiempoRespw = 0 Then
218         Call ReSpawnNpc(MiNPC)
        Else
220         MiNPC.flags.NPCActive = True
222         Indice = ObtenerIndiceRespawn
224         RespawnList(Indice) = MiNPC

        End If

        Exit Sub
ErrHandler:
226     Call TraceError(Err.Number, Err.Description & "->" & Erl(), "NPCs.MuereNpc", Erl())

End Sub

Sub ResetNpcFlags(ByVal NpcIndex As Integer)

        'Clear the npc's flags
        On Error GoTo ResetNpcFlags_Err

100     With NpcList(NpcIndex).flags
102         .AfectaParalisis = 0
104         .AguaValida = 0
106         .AttackedBy = vbNullString
107         .AttackedTime = 0
108         .AttackedFirstBy = vbNullString
112         .backup = 0
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
150         Call ResetMask(.BehaviorFlags)
152         Call SetMask(.BehaviorFlags, e_BehaviorFlags.eAttackNpc Or e_BehaviorFlags.eAttackUsers)
154         .AIAlineacion = e_Alineacion.ninguna
156         .NPCIdle = False
159         Call ClearNpcRef(.Summoner)

        End With

        Exit Sub
ResetNpcFlags_Err:
158     Call TraceError(Err.Number, Err.Description, "NPCs.ResetNpcFlags", Erl)

End Sub

Sub ResetNpcCounters(ByVal NpcIndex As Integer)

        On Error GoTo ResetNpcCounters_Err

100     NpcList(NpcIndex).Contadores.Paralisis = 0
102     NpcList(NpcIndex).Contadores.TiempoExistencia = 0
104     NpcList(NpcIndex).Contadores.IntervaloMovimiento = 0
106     NpcList(NpcIndex).Contadores.IntervaloAtaque = 0
108     NpcList(NpcIndex).Contadores.IntervaloLanzarHechizo = 0
110     NpcList(NpcIndex).Contadores.IntervaloRespawn = 0
112     NpcList(NpcIndex).Contadores.CriaturasInvocadas = 0
        Exit Sub
ResetNpcCounters_Err:
114     Call TraceError(Err.Number, Err.Description, "NPCs.ResetNpcCounters", Erl)

End Sub

Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)

        On Error GoTo ResetNpcCharInfo_Err

100     NpcList(NpcIndex).Char.body = 0
102     NpcList(NpcIndex).Char.CascoAnim = 0
104     NpcList(NpcIndex).Char.charindex = 0
106     NpcList(NpcIndex).Char.FX = 0
108     NpcList(NpcIndex).Char.head = 0
110     NpcList(NpcIndex).Char.Heading = 0
112     NpcList(NpcIndex).Char.loops = 0
114     NpcList(NpcIndex).Char.ShieldAnim = 0
116     NpcList(NpcIndex).Char.WeaponAnim = 0
118     NpcList(NpcIndex).Char.CartAnim = 0
        Exit Sub
ResetNpcCharInfo_Err:
120     Call TraceError(Err.Number, Err.Description, "NPCs.ResetNpcCharInfo", Erl)

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
110     Call TraceError(Err.Number, Err.Description, "NPCs.ResetNpcCriatures", Erl)

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
108     Call TraceError(Err.Number, Err.Description, "NPCs.ResetExpresiones", Erl)

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
108     Call TraceError(Err.Number, Err.Description, "NPCs.ResetDrop", Erl)

End Sub

Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)

        On Error GoTo ResetNpcMainInfo_Err

        With (NpcList(NpcIndex))
100         .Attackable = 0
102         .Comercia = 0
104         .GiveEXP = 0
106         .GiveEXPClan = 0
108         .GiveGLD = 0
110         .Hostile = 0
112         .InvReSpawn = 0
114         .nivel = 0
116         Call ClearNpcRef(.MaestroNPC)
118         .Mascotas = 0
120         .Movement = 0
122         .name = "NPC SIN INICIAR"
124         .npcType = 0
126         .Numero = 0
128         .Orig.Map = 0
130         .Orig.x = 0
132         .Orig.y = 0
134         .PoderAtaque = 0
136         .PoderEvasion = 0
138         .pos.Map = 0
140         .pos.x = 0
142         .pos.y = 0
144         Call SetUserRef(.TargetUser, 0)
146         Call ClearNpcRef(.TargetNPC)
148         .TipoItems = 0
150         .Veneno = 0
152         .Desc = vbNullString
154         .NumDropQuest = 0

156         If IsValidUserRef(.MaestroUser) Then Call QuitarMascota(.MaestroUser.ArrayIndex, NpcIndex)
158         If IsValidNpcRef(.MaestroNPC) Then Call QuitarMascotaNpc(.MaestroNPC.ArrayIndex)
160         Call SetUserRef(.MaestroUser, 0)
162         Call ClearNpcRef(.MaestroNPC)
164         .CaminataActual = 0

            Dim j As Integer

166         For j = 1 To .flags.LanzaSpells
168             .Spells(j).SpellIndex = 0
170         Next j

            Call ClearEffectList(.EffectOverTime)
            Call ClearModifiers(.Modifiers)

        End With

172     Call ResetNpcCharInfo(NpcIndex)
174     Call ResetNpcCriatures(NpcIndex)
176     Call ResetExpresiones(NpcIndex)
178     Call ResetDrop(NpcIndex)
        Exit Sub
ResetNpcMainInfo_Err:
180     Call TraceError(Err.Number, Err.Description, "NPCs.ResetNpcMainInfo", Erl)

End Sub

Sub QuitarNPC(ByVal NpcIndex As Integer, ByVal releaseReason As e_DeleteSource)

        On Error GoTo ErrHandler

        If Not ReleaseNpc(NpcIndex, releaseReason) Then
            Exit Sub

        End If

        If IsValidNpcRef(NpcList(NpcIndex).flags.Summoner) Then
            If NpcList(NpcList(NpcIndex).flags.Summoner.ArrayIndex).Contadores.CriaturasInvocadas > 0 Then
                'Resto 1 Npc invocado al invocador
                NpcList(NpcList(NpcIndex).flags.Summoner.ArrayIndex).Contadores.CriaturasInvocadas = NpcList(NpcList(NpcIndex).flags.Summoner.ArrayIndex).Contadores.CriaturasInvocadas - 1

                'También lo saco de la lista
                Dim LoopC As Long

                For LoopC = 1 To NpcList(NpcList(NpcIndex).flags.Summoner.ArrayIndex).Stats.CantidadInvocaciones

                    If NpcList(NpcList(NpcIndex).flags.Summoner.ArrayIndex).Stats.NpcsInvocados(LoopC).ArrayIndex = NpcIndex Then
                        Call ClearNpcRef(NpcList(NpcList(NpcIndex).flags.Summoner.ArrayIndex).Stats.NpcsInvocados(LoopC))
                        Exit For

                    End If

                Next LoopC

            End If

        ElseIf NpcList(NpcIndex).Contadores.CriaturasInvocadas > 0 Then

            Dim i As Long

            For i = 1 To NpcList(NpcIndex).Stats.CantidadInvocaciones

                If IsValidNpcRef(NpcList(NpcIndex).Stats.NpcsInvocados(i)) Then
                    Call MuereNpc(NpcList(NpcIndex).Stats.NpcsInvocados(i).ArrayIndex, 0)

                End If

            Next i

        End If

102     If InMapBounds(NpcList(NpcIndex).pos.Map, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y) Then
104         Call EraseNPCChar(NpcIndex)

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
132     Call LogError("Error en QuitarNPC")

End Sub

Function TestSpawnTrigger(ByVal Map As Integer, _
                          ByVal x As Integer, _
                          ByVal y As Integer) As Boolean

        On Error GoTo TestSpawnTrigger_Err

100     TestSpawnTrigger = MapData(Map, x, y).trigger < 1 Or (MapData(Map, x, y).trigger > 3 And MapData(Map, x, y).trigger < 12)
        Exit Function
TestSpawnTrigger_Err:
102     Call TraceError(Err.Number, Err.Description, "NPCs.TestSpawnTrigger", Erl)

End Function

Public Function CrearNPC(NroNPC As Integer, _
                         Mapa As Integer, _
                         OrigPos As t_WorldPos, _
                         Optional ByVal CustomHead As Integer)

        'Crea un NPC del tipo NRONPC
        On Error GoTo CrearNPC_Err

        Dim NpcIndex    As Integer
        Dim Iteraciones As Long
        Dim PuedeAgua   As Boolean
        Dim PuedeTierra As Boolean
        Dim Map         As Integer
        Dim x           As Integer
        Dim y           As Integer
100     NpcIndex = OpenNPC(NroNPC) 'Conseguimos un indice

102     If NpcIndex = 0 Then Exit Function

104     With NpcList(NpcIndex)

            ' Cabeza customizada
106         If CustomHead <> 0 Then .Char.head = CustomHead
108         PuedeAgua = .flags.AguaValida = 1
110         PuedeTierra = .flags.TierraInvalida = 0

            'Necesita ser respawned en un lugar especifico
112         If .flags.RespawnOrigPos And InMapBounds(OrigPos.Map, OrigPos.x, OrigPos.y) Then
114             Map = OrigPos.Map
116             x = OrigPos.x
118             y = OrigPos.y
120             .Orig = OrigPos
122             .pos = OrigPos
            Else
                ' Primera búsqueda: buscamos una posición ideal hasta llegar al máximo de iteraciones
                Do
124                 .pos.Map = Mapa
126                 .pos.x = RandomNumber(MinXBorder + 2, MaxXBorder - 2) 'Obtenemos posicion al azar en x
128                 .pos.y = RandomNumber(MinYBorder + 2, MaxYBorder - 2) 'Obtenemos posicion al azar en y
130                 .pos = ClosestLegalPosNPC(NpcIndex, 10, , True)     'Nos devuelve la posicion valida mas cercana
132                 Iteraciones = Iteraciones + 1
134             Loop While .pos.x = 0 And .pos.y = 0 And Iteraciones < MAXSPAWNATTEMPS

                ' Si no encontramos una posición válida en la primera instancia
136             If Iteraciones >= MAXSPAWNATTEMPS Then
                    ' Hacemos una búsqueda exhaustiva partiendo desde el centro del mapa
138                 .pos.Map = Mapa
140                 .pos.x = (XMaxMapSize - XMinMapSize) \ 2
142                 .pos.y = (YMaxMapSize - YMinMapSize) \ 2
144                 .pos = ClosestLegalPosNPC(NpcIndex, (XMaxMapSize - XMinMapSize) \ 2, , True)

                    ' Si sigue fallando
146                 If .pos.x = 0 And .pos.y = 0 Then
                        ' Hacemos una última búsqueda exhaustiva, ignorando los usuarios
148                     .pos.Map = Mapa
150                     .pos.x = (XMaxMapSize - XMinMapSize) \ 2
152                     .pos.y = (YMaxMapSize - YMinMapSize) \ 2
154                     .pos = ClosestLegalPosNPC(NpcIndex, (XMaxMapSize - XMinMapSize) \ 2, True)

                        ' Si falló, borramos el NPC y salimos
156                     If .pos.x = 0 And .pos.y = 0 Then
158                         Call QuitarNPC(NpcIndex, eFailToFindSpawnPos)
                            Exit Function

                        End If

                    End If

                End If

                'asignamos las nuevas coordenas
160             Map = .pos.Map
162             x = .pos.x
164             y = .pos.y
                'Y tambien asignamos su posicion original, para tener una posicion de retorno.
166             .Orig.Map = .pos.Map
168             .Orig.x = .pos.x
170             .Orig.y = .pos.y

            End If

        End With

        'Crea el NPC
172     Call MakeNPCChar(True, Map, NpcIndex, Map, x, y)
174     CrearNPC = NpcIndex
        Exit Function
CrearNPC_Err:
176     Call TraceError(Err.Number, Err.Description, "NPCs.CrearNPC", Erl)

End Function

Sub MakeNPCChar(ByVal toMap As Boolean, _
                sndIndex As Integer, _
                NpcIndex As Integer, _
                ByVal Map As Integer, _
                ByVal x As Integer, _
                ByVal y As Integer)

        On Error GoTo MakeNPCChar_Err

100     With NpcList(NpcIndex)

            Dim charindex As Integer

102         If .Char.charindex = 0 Then
104             charindex = NextOpenCharIndex
106             .Char.charindex = charindex
108             CharList(charindex) = NpcIndex

            End If

110         MapData(Map, x, y).NpcIndex = NpcIndex

            Dim Simbolo As Byte
            Dim GG      As String
            Dim tmpByte As Byte
112         GG = IIf(.showName > 0, .name & .SubName, vbNullString)

114         If Not toMap Then
116             If .NumQuest > 0 Then

                    Dim q             As Byte
                    Dim HayFinalizada As Boolean
                    Dim HayDisponible As Boolean
                    Dim HayPendiente  As Boolean

118                 For q = 1 To .NumQuest
120                     tmpByte = TieneQuest(sndIndex, .QuestNumber(q))

122                     If tmpByte Then
124                         If FinishQuestCheck(sndIndex, .QuestNumber(q), tmpByte) Then
126                             Simbolo = 3
128                             HayFinalizada = True
                            Else
130                             HayPendiente = True
132                             Simbolo = 4

                            End If

                        Else

134                         If UserDoneQuest(sndIndex, .QuestNumber(q)) Or Not UserDoneQuest(sndIndex, QuestList(.QuestNumber(q)).RequiredQuest) Or UserList(sndIndex).Stats.ELV < QuestList(.QuestNumber(q)).RequiredLevel Or UserList(sndIndex).clase = QuestList(.QuestNumber(q)).RequiredClass Then
136                             Simbolo = 2
                            Else
138                             Simbolo = 1
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

                Dim body As Integer

                'Si está muerto el usuario y en zona insegura
                If UserList(sndIndex).flags.Muerto = 1 And MapInfo(UserList(sndIndex).pos.Map).Seguro = 0 Then

                    'Solamente mando el body si es de tipo revividor.
                    If .npcType = e_NPCType.Revividor Then
                        body = IIf(.flags.NPCIdle, .Char.BodyIdle, .Char.body)
                    Else
                        body = 0

                    End If

                Else
                    body = IIf(.flags.NPCIdle, .Char.BodyIdle, .Char.body)

                End If

156             If UserList(sndIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 90 Then
158                 Call WriteCharacterCreate(sndIndex, body, .Char.head, .Char.Heading, .Char.charindex, x, y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, .Char.CartAnim, GG, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, .Char.speeding, IIf(.MaestroUser.ArrayIndex = sndIndex, 2, 1), 0, 0, 0, 0, .Stats.MinHp, .Stats.MaxHp, 0, 0, Simbolo, .flags.NPCIdle, , , .flags.team, , .Char.Ataque1)
                Else
160                 Call WriteCharacterCreate(sndIndex, body, .Char.head, .Char.Heading, .Char.charindex, x, y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, .Char.CartAnim, GG, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, .Char.speeding, IIf(.MaestroUser.ArrayIndex = sndIndex, 2, 1), 0, 0, 0, 0, 0, 0, 0, 0, Simbolo, .flags.NPCIdle, , , .flags.team, , .Char.Ataque1)

                End If

                If IsSet(.flags.StatusMask, e_StatusMask.eDontBlockTile) Then
                    Call SendData(ToIndex, sndIndex, PrepareUpdateCharValue(.Char.charindex, e_CharValue.eDontBlockTile, True))

                End If

            Else
162             Call AgregarNpc(NpcIndex)

            End If

        End With

        Exit Sub
MakeNPCChar_Err:
164     Call TraceError(Err.Number, Err.Description, "NPCs.MakeNPCChar", Erl)

End Sub

Sub ChangeNPCChar(ByVal NpcIndex As Integer, _
                  ByVal body As Integer, _
                  ByVal head As Integer, _
                  ByVal Heading As e_Heading)

        On Error GoTo ChangeNPCChar_Err

100     With NpcList(NpcIndex)

102         If NpcIndex > 0 Then
104             If .flags.NPCIdle Then
106                 body = .Char.BodyIdle

                End If

108             .Char.head = head
110             .Char.Heading = Heading

                If .Char.charindex > 0 Then
112                 Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageCharacterChange(body, head, Heading, .Char.charindex, 0, 0, 0, 0, 0, 0, .flags.NPCIdle, False))

                End If

            End If

        End With

        Exit Sub
ChangeNPCChar_Err:
114     Call TraceError(Err.Number, Err.Description, "NPCs.ChangeNPCChar", Erl)

End Sub

Sub EraseNPCChar(ByVal NpcIndex As Integer)

        On Error GoTo EraseNPCChar_Err

100     If NpcList(NpcIndex).Char.charindex <> 0 Then CharList(NpcList(NpcIndex).Char.charindex) = 0
102     If NpcList(NpcIndex).Char.charindex = LastChar Then

104         Do Until CharList(LastChar) > 0
106             LastChar = LastChar - 1

108             If LastChar <= 1 Then Exit Do
            Loop

        End If

        Call RemoveNpc(NpcIndex)
        'Quitamos del mapa
110     MapData(NpcList(NpcIndex).pos.Map, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y).NpcIndex = 0
        'Actualizamos los clientes
112     Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterRemove(5, NpcList(NpcIndex).Char.charindex, True))
        'Update la lista npc
114     NpcList(NpcIndex).Char.charindex = 0
        'update NumChars
116     NumChars = NumChars - 1
        Exit Sub
EraseNPCChar_Err:
118     Call TraceError(Err.Number, Err.Description, "NPCs.EraseNPCChar", Erl)

End Sub

Public Sub TranslateNpcChar(ByVal NpcIndex As Integer, _
                            ByRef NewPos As t_WorldPos, _
                            ByVal Speed As Long)

    On Error GoTo TranslateNpcChar_Err

    With NpcList(NpcIndex)

        If MapData(.pos.Map, NewPos.x, NewPos.y).UserIndex Then
            Call SwapTargetUserPos(MapData(.pos.Map, NewPos.x, NewPos.y).UserIndex, .pos)

        End If

        'Update map and user pos
        MapData(.pos.Map, .pos.x, .pos.y).NpcIndex = 0

        Dim PrevPos As t_WorldPos
        PrevPos = .pos
        .pos = NewPos
        MapData(.pos.Map, NewPos.x, NewPos.y).NpcIndex = NpcIndex
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareCharacterTranslate(.Char.charindex, NewPos.x, NewPos.y, Speed))
        Call CheckUpdateNeededNpc(NpcIndex, GetHeadingFromWorldPos(PrevPos, NewPos))

    End With

    Exit Sub
TranslateNpcChar_Err:
    Call TraceError(Err.Number, Err.Description, "NPCs.TranslateNpcChar", Erl)

End Sub

Public Function MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte) As Boolean

        On Error GoTo errh

        Dim nPos      As t_WorldPos
        Dim UserIndex As Integer
        Dim esGuardia As Boolean

100     With NpcList(NpcIndex)

102         If Not NPCs.CanMove(.Contadores, .flags) Then Exit Function
104         nPos = .pos
106         Call HeadtoPos(nHeading, nPos)
108         esGuardia = .npcType = e_NPCType.GuardiaReal Or .npcType = e_NPCType.GuardiasCaos

            ' es una posicion legal
110         If LegalWalkNPC(nPos.Map, nPos.x, nPos.y, nHeading, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IsValidUserRef(.MaestroUser), , esGuardia) Then
112             UserIndex = MapData(.pos.Map, nPos.x, nPos.y).UserIndex

                ' Si hay un usuario a donde se mueve el npc, entonces esta muerto o es un gm invisible
114             If UserIndex > 0 Then

116                 With UserList(UserIndex)
                        ' Actualizamos posicion y mapa
118                     MapData(.pos.Map, .pos.x, .pos.y).UserIndex = 0
120                     .pos.x = NpcList(NpcIndex).pos.x
122                     .pos.y = NpcList(NpcIndex).pos.y
124                     MapData(.pos.Map, .pos.x, .pos.y).UserIndex = UserIndex
                        ' Avisamos a los usuarios del area, y al propio usuario lo forzamos a moverse
126                     Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.charindex, .pos.x, .pos.y))
128                     Call WriteForceCharMove(UserIndex, InvertHeading(nHeading))

                    End With

                End If

                ' Solo NPCs hum
                If NpcList(NpcIndex).Humanoide Or NpcList(NpcIndex).npcType = e_NPCType.GuardiaReal Or NpcList(NpcIndex).npcType = e_NPCType.GuardiasCaos Or NpcList(NpcIndex).npcType = e_NPCType.GuardiaNpc Then
130                 If HayPuerta(nPos.Map, nPos.x, nPos.y) Then
132                     Call AccionParaPuertaNpc(nPos.Map, nPos.x, nPos.y, NpcIndex)
134                 ElseIf HayPuerta(nPos.Map, nPos.x + 1, nPos.y) Then
136                     Call AccionParaPuertaNpc(nPos.Map, nPos.x + 1, nPos.y, NpcIndex)
138                 ElseIf HayPuerta(nPos.Map, nPos.x + 1, nPos.y - 1) Then
140                     Call AccionParaPuertaNpc(nPos.Map, nPos.x + 1, nPos.y - 1, NpcIndex)
142                 ElseIf HayPuerta(nPos.Map, nPos.x, nPos.y - 1) Then
144                     Call AccionParaPuertaNpc(nPos.Map, nPos.x, nPos.y - 1, NpcIndex)

                    End If

                End If

146             Call AnimacionIdle(NpcIndex, False)
148             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(.Char.charindex, nPos.x, nPos.y))
                'Update map and user pos
150             MapData(.pos.Map, .pos.x, .pos.y).NpcIndex = 0
152             .pos = nPos
154             .Char.Heading = nHeading
156             MapData(.pos.Map, nPos.x, nPos.y).NpcIndex = NpcIndex
158             Call CheckUpdateNeededNpc(NpcIndex, nHeading)

                If Not MapData(.pos.Map, nPos.x, nPos.y).Trap Is Nothing Then
                    Call ModMap.ActivateTrap(NpcIndex, eNpc, .pos.Map, nPos.x, nPos.y)

                End If

                ' Npc has moved
160             MoveNPCChar = True

            End If

        End With

        Exit Function
errh:
162     LogError ("Error en move npc " & NpcIndex & ". Error: " & Err.Number & " - " & Err.Description)

End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer, ByVal VenenoNivel As Byte)

        On Error GoTo NpcEnvenenarUser_Err

        Dim n As Integer
100     n = RandomNumber(1, 100)

102     If n < 30 Then
104         UserList(UserIndex).flags.Envenenado = VenenoNivel

            'Msg182=¡¡La criatura te ha envenenado!!
106         If UserList(UserIndex).ChatCombate = 1 Then
108             Call WriteLocaleMsg(UserIndex, "182", e_FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If

        Exit Sub
NpcEnvenenarUser_Err:
110     Call TraceError(Err.Number, Err.Description, "NPCs.NpcEnvenenarUser", Erl)

End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, _
                  pos As t_WorldPos, _
                  ByVal FX As Boolean, _
                  ByVal Respawn As Boolean, _
                  Optional Avisar As Boolean = False, _
                  Optional ByVal MaestroUser As Integer = 0) As Integer

        On Error GoTo SpawnNpc_Err

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 23/01/2007
        '23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
        '***************************************************
        Dim NewPos      As t_WorldPos
        Dim altpos      As t_WorldPos
        Dim nIndex      As Integer
        Dim PuedeAgua   As Boolean
        Dim PuedeTierra As Boolean
        Dim Map         As Integer
        Dim x           As Integer
        Dim y           As Integer
100     nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

102     If nIndex = 0 Then
104         SpawnNpc = 0
            Exit Function

        End If

106     PuedeAgua = NpcList(nIndex).flags.AguaValida = 1
108     PuedeTierra = NpcList(nIndex).flags.TierraInvalida = 0
110     Call ClosestLegalPos(pos, NewPos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana

        'Si X e Y son iguales a 0 significa que no se encontro posicion valida
112     If NewPos.x <> 0 And NewPos.y <> 0 Then
            'Asignamos las nuevas coordenas solo si son validas
114         NpcList(nIndex).pos.Map = NewPos.Map
116         NpcList(nIndex).pos.x = NewPos.x
118         NpcList(nIndex).pos.y = NewPos.y
        Else
120         Call QuitarNPC(nIndex, eFailToFindSpawnPos)
122         SpawnNpc = 0
            Exit Function

        End If

        'asignamos las nuevas coordenas
124     Map = NewPos.Map
126     x = NpcList(nIndex).pos.x
128     y = NpcList(nIndex).pos.y
        ' WyroX: Asignamos el dueño
130     Call SetUserRef(NpcList(nIndex).MaestroUser, MaestroUser)
132     NpcList(nIndex).Orig = NpcList(nIndex).pos
        'Crea el NPC
134     Call MakeNPCChar(True, Map, nIndex, Map, x, y)

136     If FX Then
138         Call SendData(SendTarget.ToNPCAliveArea, nIndex, PrepareMessagePlayWave(SND_WARP, x, y))
140         Call SendData(SendTarget.ToNPCAliveArea, nIndex, PrepareMessageCreateFX(NpcList(nIndex).Char.charindex, e_FXIDs.FXWARP, 0))

        End If

142     If Avisar Then
144         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(NpcList(nIndex).name & " ha aparecido en " & get_map_name(Map) & " , todo indica que puede tener una gran recompensa para el que logre sobrevivir a él.", e_FontTypeNames.FONTTYPE_CITIZEN))

        End If

146     SpawnNpc = nIndex
        Exit Function
SpawnNpc_Err:
148     Call TraceError(Err.Number, Err.Description, "NPCs.SpawnNpc", Erl)

End Function

Sub ReSpawnNpc(MiNPC As t_Npc)

        On Error GoTo ReSpawnNpc_Err

100     If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.pos.Map, MiNPC.Orig)
        Exit Sub
ReSpawnNpc_Err:
102     Call TraceError(Err.Number, Err.Description, "NPCs.ReSpawnNpc", Erl)

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
104         If NpcList(NpcIndex).flags.NPCActive And NpcList(NpcIndex).pos.Map = Map And NpcList(NpcIndex).Hostile = 1 Then
106             cont = cont + 1

            End If

108     Next NpcIndex

110     NPCHostiles = cont
        Exit Function
NPCHostiles_Err:
112     Call TraceError(Err.Number, Err.Description, "NPCs.NPCHostiles", Erl)

End Function

Sub NPCTirarOro(MiNPC As t_Npc, ByVal UserIndex As Integer)

        On Error GoTo NPCTirarOro_Err

100     If UserIndex = 0 Then Exit Sub
102     If MiNPC.GiveGLD > 0 Then

            Dim Oro As Long
104         Oro = MiNPC.GiveGLD * SvrConfig.GetValue("GoldMult")

            Dim MiObj As t_Obj
118         MiObj.ObjIndex = iORO

120         While (Oro > 0)

122             If Oro > MAX_INVENTORY_OBJS Then
124                 MiObj.amount = MAX_INVENTORY_OBJS
126                 Oro = Oro - MAX_INVENTORY_OBJS
                Else
128                 MiObj.amount = Oro
130                 Oro = 0

                End If

132             Call TirarItemAlPiso(MiNPC.pos, MiObj, MiNPC.flags.AguaValida = 1)
            Wend
134         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageFxPiso("87", MiNPC.pos.x, MiNPC.pos.y))

        End If

        Exit Sub
NPCTirarOro_Err:
136     Call TraceError(Err.Number, Err.Description, "NPCs.NPCTirarOro", Erl)

End Sub

Function UpdateNpcSpeed(ByVal NpcIndex As Integer)

        With NpcList(NpcIndex)

214         If .IntervaloMovimiento = 0 Then
216             .IntervaloMovimiento = 380
218             .Char.speeding = frmMain.TIMER_AI.Interval / 330
            Else
220             .Char.speeding = 210 / .IntervaloMovimiento

            End If

            .Char.speeding = .Char.speeding * max(0, (1 + .Modifiers.MovementSpeed))
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageSpeedingACT(.Char.charindex, .Char.speeding))

        End With

End Function

Function GetNpcSpeedModifiers(ByVal NpcIndex As Integer) As Single
    GetNpcSpeedModifiers = max(0, (1 + NpcList(NpcIndex).Modifiers.MovementSpeed))

End Function

Function GetNpcName(ByVal NpcNumber As Integer) As String
    GetNpcName = LeerNPCs.GetValue("NPC" & NpcNumber, "Name")

End Function

Function OpenNPC(ByVal NpcNumber As Integer, _
                 Optional ByVal Respawn As Boolean = True, _
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
        'nueva clase clsIniManager.
        '
        'Alejo
        '
        '###################################################
        Dim NpcIndex As Integer
        Dim Leer     As clsIniManager
100     Set Leer = LeerNPCs

        'If requested index is invalid, abort
102     If Not Leer.KeyExists("NPC" & NpcNumber) Then
104         OpenNPC = 0
            Exit Function

        End If

        #If DEBUGGING = 0 Then

            If val(Leer.GetValue("NPC" & NpcNumber, "TESTONLY")) > 0 Then Exit Function
        #End If

        Dim RequireToggle As String
        RequireToggle = Leer.GetValue("NPC" & NpcNumber, "REQUIRETOGGLE")

        If RequireToggle <> "" Then
            If Not IsFeatureEnabled(RequireToggle) Then Exit Function

        End If

106     NpcIndex = GetNextAvailableNpc

108     If NpcIndex > MaxNPCs Then 'Limite de npcs
110         OpenNPC = 0
            Exit Function

        End If

        Dim LoopC   As Long
        Dim ln      As String
        Dim aux     As String
        Dim Field() As String

112     With NpcList(NpcIndex)
114         .Numero = NpcNumber
116         .name = Leer.GetValue("NPC" & NpcNumber, "Name")
118         .SubName = Leer.GetValue("NPC" & NpcNumber, "SubName")
120         .Desc = Leer.GetValue("NPC" & NpcNumber, "Desc")
122         .nivel = val(Leer.GetValue("NPC" & NpcNumber, "Nivel"))
124         Call SetMovement(NpcIndex, val(Leer.GetValue("NPC" & NpcNumber, "Movement")))
126         .flags.OldMovement = .Movement
128         .flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
130         .flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
132         .flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))
134         .npcType = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))
136         .Char.body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
138         .Char.head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
140         .Char.Heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))
142         .Char.BodyIdle = val(Leer.GetValue("NPC" & NpcNumber, "BodyIdle"))
143         .Char.Ataque1 = val(Leer.GetValue("NPC" & NpcNumber, "Ataque1"))
            .Char.CastAnimation = val(Leer.GetValue("NPC" & NpcNumber, "CastAnimation"))

            If .Char.BodyIdle > 0 Then
                .flags.NPCIdle = True

            End If

            Dim CantidadAnimaciones As Integer
144         CantidadAnimaciones = val(Leer.GetValue("NPC" & NpcNumber, "Animaciones"))

146         If CantidadAnimaciones > 0 Then
148             ReDim .Char.Animation(1 To CantidadAnimaciones)

150             For LoopC = 1 To CantidadAnimaciones
152                 .Char.Animation(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Anim" & LoopC))
                Next
            Else
154             ReDim .Char.Animation(0)

            End If

156         .Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "Arma"))
158         .Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "Escudo"))
160         .Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "Casco"))
161         .Char.CartAnim = val(Leer.GetValue("NPC" & NpcNumber, "Cart"))
162         .Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
164         .Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
166         .Craftea = val(Leer.GetValue("NPC" & NpcNumber, "Craftea"))
168         .Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
170         .flags.OldHostil = .Hostile
            .AttackRange = val(Leer.GetValue("NPC" & NpcNumber, "AttackRange"))
            .ProjectileType = val(Leer.GetValue("NPC" & NpcNumber, "ProjectileType"))
            .PreferedRange = val(Leer.GetValue("NPC" & NpcNumber, "PreferedRange"))
172         .GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP"))
174         .Distancia = val(Leer.GetValue("NPC" & NpcNumber, "Distancia"))
176         .GiveEXPClan = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXPClan"))
            '.flags.ExpDada = .GiveEXP
178         .flags.ExpCount = .GiveEXP
180         .Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))
182         .flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))
184         .GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD"))
            '166        .QuestNumber = val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber"))
186         .PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
188         .PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
190         .InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
192         .showName = val(Leer.GetValue("NPC" & NpcNumber, "ShowName"))
194         .GobernadorDe = val(Leer.GetValue("NPC" & NpcNumber, "GobernadorDe"))
196         .SoundOpen = val(Leer.GetValue("NPC" & NpcNumber, "SoundOpen"))
198         .SoundClose = val(Leer.GetValue("NPC" & NpcNumber, "SoundClose"))
200         .IntervaloAtaque = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloAtaque"))
202         .IntervaloMovimiento = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloMovimiento"))
204         .IntervaloLanzarHechizo = val(Leer.GetValue("NPC" & NpcNumber, "IntervaloLanzarHechizo"))
206         .Contadores.IntervaloRespawn = RandomNumber(val(Leer.GetValue("NPC" & NpcNumber, "IntervaloRespawnMin")), val(Leer.GetValue("NPC" & NpcNumber, "IntervaloRespawn")))
208         .InformarRespawn = val(Leer.GetValue("NPC" & NpcNumber, "InformarRespawn"))
210         .QuizaProb = val(Leer.GetValue("NPC" & NpcNumber, "QuizaProb"))

214         If .IntervaloMovimiento = 0 Then
216             .IntervaloMovimiento = 380
218             .Char.speeding = frmMain.TIMER_AI.Interval / 330
            Else
220             .Char.speeding = 210 / .IntervaloMovimiento

            End If

222         If .IntervaloLanzarHechizo = 0 Then
224             .IntervaloLanzarHechizo = 8000

            End If

226         If .IntervaloAtaque = 0 Then
228             .IntervaloAtaque = 2000

            End If

230         .Stats.MaxHp = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
232         .Stats.MinHp = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
234         .Stats.MaxHit = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
236         .Stats.MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
238         .Stats.def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
239         .Stats.defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
240         .Stats.MagicResistance = val(Leer.GetValue("NPC" & NpcNumber, "MagicResistance"))
241         .Stats.MagicDef = val(Leer.GetValue("NPC" & NpcNumber, "MagicDef"))
242         .Stats.CantidadInvocaciones = val(Leer.GetValue("NPC" & NpcNumber, "CantidadInvocaciones"))
243         .Stats.MagicBonus = val(Leer.GetValue("NPC" & NpcNumber, "MagicBonus"))

            If .Stats.CantidadInvocaciones > 0 Then
244             ReDim .Stats.NpcsInvocados(1 To .Stats.CantidadInvocaciones)

                For LoopC = 1 To .Stats.CantidadInvocaciones
                    Call ClearNpcRef(.Stats.NpcsInvocados(LoopC))
                Next LoopC

            End If

245         .flags.AIAlineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))
246         .invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))

            If .invent.NroItems > UBound(.invent.Object) Then
                Debug.Print "Error in NPC " & .name & " configuration, too many items in inventory! Max possible is " & UBound(.invent.Object)

            End If

            Debug.Assert .invent.NroItems <= UBound(.invent.Object) ' NPC not properly configured, it cannot have more so many items in the inventory
            .invent.NroItems = Min(UBound(.invent.Object), .invent.NroItems) 'Clamp the number of items to make sure it's not too big
247         .Humanoide = CBool(val(Leer.GetValue("NPC" & NpcNumber, "Humanoide")))

248         For LoopC = 1 To .invent.NroItems
249             ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
250             .invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
252             .invent.Object(LoopC).amount = val(ReadField(2, ln, 45))
254         Next LoopC

256         .flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))

258         If .flags.LanzaSpells > 0 Then
260             ReDim .Spells(1 To .flags.LanzaSpells)
                .SpellRange = val(Leer.GetValue("NPC" & NpcNumber, "RangoSpell"))

            End If

262         For LoopC = 1 To .flags.LanzaSpells
264             .Spells(LoopC).SpellIndex = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
265             .Spells(LoopC).Cd = val(Leer.GetValue("NPC" & NpcNumber, "Cd" & LoopC))
                .Spells(LoopC).LastUse = 0
266         Next LoopC

268         If .npcType = e_NPCType.Entrenador Then
270             .NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))

272             If .NroCriaturas > 0 Then
274                 ReDim .Criaturas(1 To .NroCriaturas) As t_CriaturasEntrenador

276                 For LoopC = 1 To .NroCriaturas
278                     .Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
280                     .Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
282                 Next LoopC

                End If

            End If

            Call ResetMask(.flags.StatusMask)
284         .flags.NPCActive = True
            Call ResetMask(.flags.BehaviorFlags)

286         Select Case val(Leer.GetValue("NPC" & NpcNumber, "RestriccionDeAtaque"))

                Case 0 ' Todos
288                 Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eAttackNpc)
290                 Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eAttackUsers)

292             Case 1 ' Usuarios solamente
296                 Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eAttackUsers)

298             Case 2 ' NPCs solamente
300                 Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eAttackNpc)

            End Select

            Select Case val(Leer.GetValue("NPC" & NpcNumber, "RestriccionDeAyuda"))

                Case 1
                    Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eHelpNpc)

                Case 2
                    Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eHelpUsers)

                Case 3
                    Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eHelpNpc)
                    Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eHelpUsers)

            End Select

304         If Respawn Then
306             .flags.Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
            Else
308             .flags.Respawn = 1

            End If

            If val(Leer.GetValue("NPC" & NpcNumber, "DontHitVisiblePlayers")) > 0 Then Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eDontHitVisiblePlayers)
            If val(Leer.GetValue("NPC" & NpcNumber, "AddToMapAiList")) > 0 Then Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eConsideredByMapAi)
            If val(Leer.GetValue("NPC" & NpcNumber, "DisplayCastMessage")) > 0 Then Call SetMask(.flags.BehaviorFlags, e_BehaviorFlags.eDisplayCastMessage)
            .flags.team = val(Leer.GetValue("NPC" & NpcNumber, "Team"))
310         .flags.backup = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
312         .flags.RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
314         .flags.AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
316         .flags.GolpeExacto = val(Leer.GetValue("NPC" & NpcNumber, "GolpeExacto"))

            If val(Leer.GetValue("NPC" & NpcNumber, "TranslationInmune")) > 0 Then Call SetMask(.flags.EffectInmunity, e_Inmunities.eTranslation)
318         .flags.Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
320         .flags.Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
322         .flags.Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))
            '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
324         aux = Leer.GetValue("NPC" & NpcNumber, "NROEXP")

326         If LenB(aux) = 0 Then
328             .NroExpresiones = 0
            Else
330             .NroExpresiones = val(aux)
332             ReDim .Expresiones(1 To .NroExpresiones) As String

334             For LoopC = 1 To .NroExpresiones
336                 .Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
338             Next LoopC

            End If

            '<<<<<<<<<<<<<< Sistema de Dropeo NUEVO >>>>>>>>>>>>>>>>
340         .NumQuiza = val(Leer.GetValue("NPC" & NpcNumber, "NumQuiza"))

342         If .NumQuiza > 0 Then
348             ReDim .QuizaDropea(1 To .NumQuiza) As String

350             For LoopC = 1 To .NumQuiza
352                 .QuizaDropea(LoopC) = Leer.GetValue("NPC" & NpcNumber, "QuizaDropea" & LoopC)
354             Next LoopC

            End If

            'Ladder
            'Nuevo sistema de Quest
356         aux = Leer.GetValue("NPC" & NpcNumber, "NumQuest")

358         If LenB(aux) = 0 Then
360             .NumQuest = 0
            Else
362             .NumQuest = val(aux)
364             ReDim .QuestNumber(1 To .NumQuest) As Integer

366             For LoopC = 1 To .NumQuest
368                 .QuestNumber(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber" & LoopC))
370             Next LoopC

            End If

            'Nuevo sistema de Quest
            'Nuevo sistema de Drop Quest
372         .NumDropQuest = val(Leer.GetValue("NPC" & NpcNumber, "NumDropQuest"))

374         If .NumDropQuest Then
380             ReDim .DropQuest(1 To .NumDropQuest) As t_QuestObj

382             For LoopC = 1 To .NumDropQuest
384                 .DropQuest(LoopC).QuestIndex = val(ReadField(1, Leer.GetValue("NPC" & NpcNumber, "DropQuest" & LoopC), Asc("-")))
386                 .DropQuest(LoopC).ObjIndex = val(ReadField(2, Leer.GetValue("NPC" & NpcNumber, "DropQuest" & LoopC), Asc("-")))
388                 .DropQuest(LoopC).amount = val(ReadField(3, Leer.GetValue("NPC" & NpcNumber, "DropQuest" & LoopC), Asc("-")))
390                 .DropQuest(LoopC).Probabilidad = val(ReadField(4, Leer.GetValue("NPC" & NpcNumber, "DropQuest" & LoopC), Asc("-")))
392             Next LoopC

            End If

            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< PATHFINDING >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
394         .pathFindingInfo.RangoVision = val(Leer.GetValue("NPC" & NpcNumber, "Distancia"))

396         If .pathFindingInfo.RangoVision = 0 Then .pathFindingInfo.RangoVision = RANGO_VISION_X
398         .pathFindingInfo.Inteligencia = val(Leer.GetValue("NPC" & NpcNumber, "Inteligencia"))

400         If .pathFindingInfo.Inteligencia = 0 Then .pathFindingInfo.Inteligencia = 10
402         ReDim .pathFindingInfo.Path(1 To .pathFindingInfo.Inteligencia + RANGO_VISION_X * 3)
            '<<<<<<<<<<<<<< Sistema de Viajes NUEVO >>>>>>>>>>>>>>>>
404         aux = Leer.GetValue("NPC" & NpcNumber, "NumDestinos")

406         If LenB(aux) = 0 Then
408             .NumDestinos = 0
            Else
410             .NumDestinos = val(aux)
412             ReDim .dest(1 To .NumDestinos) As String

414             For LoopC = 1 To .NumDestinos
416                 .dest(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Dest" & LoopC)
418             Next LoopC

            End If

            '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
420         .Interface = val(Leer.GetValue("NPC" & NpcNumber, "Interface"))
            'Tipo de items con los que comercia
            .TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))
            'PuedeInvocar -> NPCs que solo ven los SemiDioses
            .PuedeInvocar = val(Leer.GetValue("NPC" & NpcNumber, "PuedeInvocar"))

            '<<<<<<<<<<<<<< Animaciones >>>>>>>>>>>>>>>>
            ' Por defecto la animación es idle
            If NumUsers > 0 Then
424             Call AnimacionIdle(NpcIndex, True)

            End If

            ' Si el tipo de movimiento es Caminata
426         If .Movement = Caminata Then

                ' Leemos la cantidad de indicaciones
                Dim cant As Byte
428             cant = val(Leer.GetValue("NPC" & NpcNumber, "CaminataLen"))

                ' Prevengo NPCs rotos
430             If cant = 0 Then
432                 Call SetMovement(NpcIndex, Estatico)
                Else
                    ' Redimenciono el array
434                 ReDim .Caminata(1 To cant)

                    ' Leo todas las indicaciones
436                 For LoopC = 1 To cant
438                     Field = Split(Leer.GetValue("NPC" & NpcNumber, "Caminata" & LoopC), ":")
440                     .Caminata(LoopC).offset.x = val(Field(0))
442                     .Caminata(LoopC).offset.y = val(Field(1))
444                     .Caminata(LoopC).Espera = val(Field(2))
                    Next
446                 .CaminataActual = 1

                End If

            End If

            '<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>
        End With

        'Si NO estamos actualizando los NPC's activos, actualizamos el contador.
448     If Reload = False Then
450         If NpcIndex > LastNPC Then LastNPC = NpcIndex
452         NumNPCs = NumNPCs + 1

        End If

        'Devuelve el nuevo Indice
454     OpenNPC = NpcIndex
        Exit Function
OpenNPC_Err:
456     Call TraceError(Err.Number, Err.Description, "NPCs.OpenNPC", Erl)

End Function

Function NpcSellsItem(ByVal NpcNumber As Integer, ByVal NroObjeto As Integer) As Boolean

        On Error GoTo NpcSellsItem_Err

        Dim Leer As clsIniManager
100     Set Leer = LeerNPCs

        'If requested index is invalid, abort
102     If Not Leer.KeyExists("NPC" & NpcNumber) Then
104         NpcSellsItem = False
            Exit Function

        End If

        Dim LoopC    As Long
        Dim ln       As String
        Dim Field()  As String
        Dim NroItems As Long
244     NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))

246     For LoopC = 1 To NroItems
248         ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)

            If NroObjeto = val(ReadField(1, ln, 45)) Then
                NpcSellsItem = True
                Exit Function

            End If

254     Next LoopC

        NpcSellsItem = False
        Exit Function
NpcSellsItem_Err:
456     Call TraceError(Err.Number, Err.Description, "NPCs.NpcSellsItem", Erl)

End Function

Sub DoFollow(ByVal NpcIndex As Integer, ByVal username As String)

        On Error GoTo DoFollow_Err

100     With NpcList(NpcIndex)

102         If .flags.Follow Then
104             .flags.AttackedBy = vbNullString
106             Call SetUserRef(.TargetUser, 0)
108             .flags.Follow = False
110             Call SetMovement(NpcIndex, .flags.OldMovement)
112             .Hostile = .flags.OldHostil
            Else

                Dim player As t_UserReference
                player = NameIndex(username)

                If IsValidUserRef(player) Then
114                 .flags.AttackedBy = username
115                 .flags.AttackedTime = GlobalFrameTime
116                 .TargetUser = player
118                 .flags.Follow = True
120                 Call SetMovement(NpcIndex, e_TipoAI.NpcDefensa)
122                 .Hostile = 0

                End If

            End If

        End With

        Exit Sub
DoFollow_Err:
124     Call TraceError(Err.Number, Err.Description, "NPCs.DoFollow", Erl)

End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)

        On Error GoTo FollowAmo_Err

100     With NpcList(NpcIndex)
102         .flags.Follow = True
104         Call SetMovement(NpcIndex, e_TipoAI.SigueAmo)
106         .Hostile = 0
108         Call ClearUserRef(.TargetUser)
110         Call ClearNpcRef(.TargetNPC)

        End With

        Exit Sub
FollowAmo_Err:
112     Call TraceError(Err.Number, Err.Description, "NPCs.FollowAmo", Erl)

End Sub

Public Sub AllFollowAmo(ByVal UserIndex As Integer)

        On Error GoTo AllFollowAmo_Err

        Dim j As Long

100     For j = 1 To MAXMASCOTAS

102         If UserList(UserIndex).MascotasIndex(j).ArrayIndex > 0 Then
                If IsValidNpcRef(UserList(UserIndex).MascotasIndex(j)) Then
104                 Call FollowAmo(UserList(UserIndex).MascotasIndex(j).ArrayIndex)
                Else
                    Call ClearNpcRef(UserList(UserIndex).MascotasIndex(j))

                End If

            End If

106     Next j

        Exit Sub
AllFollowAmo_Err:
108     Call TraceError(Err.Number, Err.Description, "SistemaCombate.AllFollowAmo", Erl)

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

102         If UserList(UserIndex).MascotasIndex(i).ArrayIndex = NpcIndex Then
104             Call ClearNpcRef(UserList(UserIndex).MascotasIndex(i))
106             UserList(UserIndex).MascotasType(i) = 0
108             UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
                UserList(UserIndex).flags.ModificoMascotas = True
                Exit For

            End If

110     Next i

        Exit Sub
QuitarMascota_Err:
112     Call TraceError(Err.Number, Err.Description, "NPCs.QuitarMascota", Erl)

End Sub

Sub AnimacionIdle(ByVal NpcIndex As Integer, ByVal Show As Boolean)

        On Error GoTo Handler

100     With NpcList(NpcIndex)

102         If .Char.BodyIdle = 0 Then Exit Sub
104         If .flags.NPCIdle = Show Then Exit Sub
106         .flags.NPCIdle = Show
108         Call ChangeNPCChar(NpcIndex, .Char.body, .Char.head, .Char.Heading)

        End With

        Exit Sub
Handler:
110     Call TraceError(Err.Number, Err.Description, "NPCs.AnimacionIdle", Erl)

End Sub

Sub WarpNpcChar(ByVal NpcIndex As Integer, _
                ByVal Map As Byte, _
                ByVal x As Integer, _
                ByVal y As Integer, _
                Optional ByVal FX As Boolean = False)

        Dim NuevaPos  As t_WorldPos
        Dim FuturePos As t_WorldPos
100     Call EraseNPCChar(NpcIndex)
102     FuturePos.Map = Map
104     FuturePos.x = x
106     FuturePos.y = y
108     Call ClosestLegalPos(FuturePos, NuevaPos, True, True)

110     If NuevaPos.Map = 0 Or NuevaPos.x = 0 Or NuevaPos.y = 0 Then
112         Debug.Print "Error al tepear NPC"
114         Call QuitarNPC(NpcIndex, eFailedToWarp)
        Else
116         NpcList(NpcIndex).pos = NuevaPos
118         Call MakeNPCChar(True, 0, NpcIndex, NuevaPos.Map, NuevaPos.x, NuevaPos.y)

120         If FX Then                                    'FX
122             Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessagePlayWave(SND_WARP, NuevaPos.x, NuevaPos.y))
124             Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageCreateFX(NpcList(NpcIndex).Char.charindex, e_FXIDs.FXWARP, 0))

            End If

        End If

End Sub

' Autor: WyroX - 20/01/2021
' Intenta moverlo hacia un "costado" según el heading indicado. Se usa para mover NPCs del camino de otro char.
' Si no hay un lugar válido a los lados, lo mueve a la posición válida más cercana.
Sub MoveNpcToSide(ByVal NpcIndex As Integer, ByVal Heading As e_Heading)

        On Error GoTo Handler

100     With NpcList(NpcIndex)

            ' Elegimos un lado al azar
            Dim r As Integer
102         r = RandomNumber(0, 1) * 2 - 1 ' -1 o 1
            ' Roto el heading original hacia ese lado
104         Heading = Rotate_Heading(Heading, r)

            ' Intento moverlo para ese lado
106         If MoveNPCChar(NpcIndex, Heading) Then Exit Sub
            ' Si falló, intento moverlo para el lado opuesto
108         Heading = InvertHeading(Heading)

110         If MoveNPCChar(NpcIndex, Heading) Then Exit Sub

            ' Si ambos fallan, entonces lo dejo en la posición válida más cercana
            Dim NuevaPos As t_WorldPos
112         Call ClosestLegalPos(.pos, NuevaPos, .flags.AguaValida, .flags.TierraInvalida = 0)
114         Call WarpNpcChar(NpcIndex, NuevaPos.Map, NuevaPos.x, NuevaPos.y)

        End With

        Exit Sub
Handler:
116     Call TraceError(Err.Number, Err.Description, "NPCs.MoveNpcToSide", Erl)

End Sub

Public Sub DummyTargetAttacked(ByVal NpcIndex As Integer)

100     With NpcList(NpcIndex)
102         .Contadores.UltimoAtaque = 30

104         If RandomNumber(1, 5) = 1 Then
106             If UBound(.Char.Animation) > 0 Then
108                 Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageDoAnimation(.Char.charindex, .Char.Animation(1)))

                End If

            End If

        End With

End Sub

Public Sub KillRandomNpc()

    Dim validNpc As Boolean: validNpc = False
    Dim NpcIndex As Integer: NpcIndex = 0

    If GetAvailableNpcIndex > 8000 Or GetAvailableNpcIndex = 0 Then
        Exit Sub

    End If

    Do While Not validNpc
        NpcIndex = RandomNumber(1, 10000)

        If NpcList(NpcIndex).flags.NPCActive And NpcList(NpcIndex).Hostile > 0 Then
            validNpc = True

        End If

    Loop
    Call MuereNpc(NpcIndex, 0)

End Sub

Public Function CanMove(ByRef counter As t_NpcCounters, _
                        ByRef flags As t_NPCFlags) As Boolean
    CanMove = flags.Inmovilizado + flags.Paralizado = 0 And counter.StunEndTime < GetTickCount() And Not flags.TranslationActive

End Function

Public Function CanAttack(ByRef counter As t_NpcCounters, _
                          ByRef flags As t_NPCFlags) As Boolean
    CanAttack = flags.Paralizado = 0 And counter.StunEndTime < GetTickCount()

End Function

Public Sub StunNPc(ByRef Counters As t_NpcCounters)
    Counters.StunEndTime = GetTickCount() + NpcStunTime

End Sub

Public Function ModifyHealth(ByVal NpcIndex As Integer, _
                             ByVal amount As Long, _
                             Optional ByVal MinValue = 0) As Boolean

    With NpcList(NpcIndex)
        ModifyHealth = False
        .Stats.MinHp = .Stats.MinHp + amount

        If .Stats.MinHp > .Stats.MaxHp Then
            .Stats.MinHp = .Stats.MaxHp

        End If

        If .Stats.MinHp <= MinValue Then
            .Stats.MinHp = MinValue
            ModifyHealth = True

        End If

        Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageNpcUpdateHP(NpcIndex))

    End With

End Function

Public Function DoDamageOrHeal(ByVal NpcIndex As Integer, _
                               ByVal SourceIndex As Integer, _
                               ByVal SourceType As e_ReferenceType, _
                               ByVal amount As Long, _
                               ByVal DamageSourceType As e_DamageSourceType, _
                               ByVal DamageSourceIndex As Integer, _
                               Optional ByVal DamageColor As Long = vbRed) As e_DamageResult

        On Error GoTo DoDamageOrHeal_Err

        Dim DamageStr As String
        Dim Color     As Long

        If amount > 0 Then
            Color = vbGreen
        Else
            Color = DamageColor

        End If

        If amount < 0 Then
            If SourceType = eUser Then
                DamageStr = PonerPuntos(Math.Abs(amount))

                If UserList(SourceIndex).ChatCombate = 1 Then
                    Call WriteLocaleMsg(SourceIndex, 382, e_FontTypeNames.FONTTYPE_FIGHT, DamageStr)

                End If

            End If

            amount = EffectsOverTime.TargetApplyDamageReduction(NpcList(NpcIndex).EffectOverTime, amount, SourceIndex, SourceType, DamageSourceType)
            Call EffectsOverTime.TargetWasDamaged(NpcList(NpcIndex).EffectOverTime, SourceIndex, SourceType, DamageSourceType)

        End If

        DamageStr = PonerPuntos(Math.Abs(amount))

        With NpcList(NpcIndex)
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageTextOverChar(DamageStr, .Char.charindex, Color))

            ' Mascotas dan experiencia al amo
            If SourceType = eNpc And amount < 0 Then
138             If IsValidUserRef(NpcList(SourceIndex).MaestroUser) Then
140                 Call CalcularDarExp(NpcList(SourceIndex).MaestroUser.ArrayIndex, NpcIndex, -amount)

                    ' NPC de invasión
142                 If .flags.InvasionIndex Then
144                     Call SumarScoreInvasion(NpcList(SourceIndex).flags.InvasionIndex, NpcList(SourceIndex).MaestroUser.ArrayIndex, -amount)

                    End If

                End If

            ElseIf SourceType = eUser And amount < 0 Then

                ' NPC de invasión
172             If NpcList(NpcIndex).flags.InvasionIndex Then
174                 Call SumarScoreInvasion(NpcList(NpcIndex).flags.InvasionIndex, SourceIndex, -amount)

                End If

186             Call CalcularDarExp(SourceIndex, NpcIndex, -amount)

            End If

100         If NPCs.ModifyHealth(NpcIndex, amount) Then
                DoDamageOrHeal = eDead
                CustomScenarios.NpcDie (NpcIndex)

102             If SourceType = eUser Then
244                 Call CustomScenarios.PlayerKillNpc(.pos.Map, NpcIndex, SourceIndex, DamageSourceType, DamageSourceIndex)
                    Call MuereNpc(NpcIndex, SourceIndex)
                Else

                    If IsValidUserRef(NpcList(SourceIndex).MaestroUser) Then
                        Call PlayerKillNpc(NpcList(NpcIndex).pos.Map, NpcIndex, NpcList(SourceIndex).MaestroUser.ArrayIndex, e_pet, DamageSourceIndex)
                        Call FollowAmo(SourceIndex)
                        Call MuereNpc(NpcIndex, NpcList(SourceIndex).MaestroUser.ArrayIndex)
                    Else
                        Call MuereNpc(NpcIndex, -1)

                    End If

                End If

                Exit Function

            End If

        End With

        DoDamageOrHeal = eStillAlive
        Exit Function
DoDamageOrHeal_Err:
134     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.DoDamageOrHeal_Err", Erl)

End Function

Public Function UserCanAttackNpc(ByVal UserIndex As Integer, _
                                 ByVal NpcIndex As Integer) As t_AttackInteractionResult

        On Error GoTo UserCanAttackNpc_Err

        UserCanAttackNpc.CanAttack = False
        UserCanAttackNpc.TurnPK = False

        'Estas muerto?
100     If UserList(UserIndex).flags.Muerto = 1 Then
104         UserCanAttackNpc.Result = eDeathAttacker
            Exit Function

        End If

106     If UserList(UserIndex).flags.Montado = 1 Then
110         UserCanAttackNpc.Result = eMounted
            Exit Function

        End If

112     If UserList(UserIndex).flags.Inmunidad = 1 Then
116         UserCanAttackNpc.Result = eCreatureInmunity
            Exit Function

        End If

        'Solo administradores, dioses y usuarios pueden atacar a NPC's (PARA TESTING)
118     If (UserList(UserIndex).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Dios Or e_PlayerType.Admin)) = 0 And NpcList(NpcIndex).npcType <> e_NPCType.DummyTarget Then
120         UserCanAttackNpc.Result = eInvalidPrivilege
            Exit Function

        End If

        ' No podes atacar si estas en consulta
122     If UserList(UserIndex).flags.EnConsulta Then
126         UserCanAttackNpc.Result = eTalkWithMaster
            Exit Function

        End If

        'Es una criatura atacable?
128     If NpcList(NpcIndex).Attackable = 0 Then
132         UserCanAttackNpc.Result = eInmuneNpc
            Exit Function

        End If

        'Es valida la distancia a la cual estamos atacando?
        If Abs(UserList(UserIndex).pos.x - NpcList(NpcIndex).pos.x) > RANGO_VISION_X Or Abs(UserList(UserIndex).pos.y - NpcList(NpcIndex).pos.y) > RANGO_VISION_Y Then
138         UserCanAttackNpc.Result = eOutOfRange
            Exit Function

        End If

        Dim IsPet As Boolean
        IsPet = IsValidUserRef(NpcList(NpcIndex).MaestroUser)

        'Si el usuario pertenece a una faccion
140     If esArmada(UserIndex) Or esCaos(UserIndex) Then

            ' Y el NPC pertenece a la misma faccion
142         If NpcList(NpcIndex).flags.Faccion = UserList(UserIndex).Faccion.Status Then
146             UserCanAttackNpc.Result = eSameFaction
                Exit Function

            End If

            ' Si es una mascota, checkeamos en el Maestro
148         If IsPet Then
150             If UserList(NpcList(NpcIndex).MaestroUser.ArrayIndex).Faccion.Status = UserList(UserIndex).Faccion.Status Then
154                 UserCanAttackNpc.Result = eSameFaction
                    Exit Function

                End If

            End If

        End If

        If IsPet Then
            If UserList(NpcList(NpcIndex).MaestroUser.ArrayIndex).Grupo.Id = UserList(UserIndex).Grupo.Id And UserList(UserIndex).Grupo.Id > 0 Then
                UserCanAttackNpc.Result = eSameGroup
                Exit Function

            End If

            If esArmada(UserIndex) Then
                If UserList(NpcList(NpcIndex).MaestroUser.ArrayIndex).Faccion.Status = Ciudadano Or UserList(NpcList(NpcIndex).MaestroUser.ArrayIndex).Faccion.Status = Armada Then
                    UserCanAttackNpc.Result = eSameFaction
                    Exit Function

                End If

            End If

        End If

156     If Status(UserIndex) = Ciudadano Then
158         If IsPet And NpcList(NpcIndex).MaestroUser.ArrayIndex = UserIndex Then
162             UserCanAttackNpc.Result = eOwnPet
                Exit Function

            End If

        End If

        If UserList(UserIndex).flags.CurrentTeam <> 0 Then
            If UserList(UserIndex).flags.CurrentTeam = NpcList(NpcIndex).flags.team Then
                UserCanAttackNpc.Result = eSameTeam
                Exit Function

            End If

        End If

        If IsPet And MapInfo(UserList(UserIndex).pos.Map).Seguro Then
            UserCanAttackNpc.Result = eSafeArea
            Exit Function

        End If

        ' El seguro es SOLO para ciudadanos. La armada debe desenlistarse antes de querer atacar y se checkea arriba.
        ' Los criminales o Caos, ya estan mas alla del seguro.
164     If Status(UserIndex) = Ciudadano Then
166         If NpcList(NpcIndex).flags.Faccion = Armada Or NpcList(NpcIndex).flags.Faccion = consejo Then
168             If UserList(UserIndex).flags.Seguro Then
172                 UserCanAttackNpc.Result = eRemoveSafe
                    Exit Function
                Else
                    UserCanAttackNpc.Result = eAttackSameFaction
                    UserCanAttackNpc.TurnPK = True
                    UserCanAttackNpc.CanAttack = True
                    Exit Function

                End If

            End If

        End If

        If Status(UserIndex) = Ciudadano Or Status(UserIndex) = Armada Or Status(UserIndex) = consejo Then

            'Es el NPC mascota de alguien?
180         If IsPet Then

182             Select Case UserList(NpcList(NpcIndex).MaestroUser.ArrayIndex).Faccion.Status

                    Case e_Facciones.Armada

184                     If UserList(UserIndex).flags.Seguro Then
188                         UserCanAttackNpc.Result = eRemoveSafe
                            Exit Function
                        Else
                            UserCanAttackNpc.TurnPK = True
192                         UserCanAttackNpc.CanAttack = True
194                         UserCanAttackNpc.Result = eAttackPetSameFaction
                            Exit Function

                        End If

196                 Case e_Facciones.Ciudadano

198                     If UserList(UserIndex).flags.Seguro Then
202                         UserCanAttackNpc.Result = eRemoveSafe
                            Exit Function
                        Else
206                         UserCanAttackNpc.TurnPK = True
                            UserCanAttackNpc.CanAttack = True
                            UserCanAttackNpc.Result = eAttackPetSameFaction
                            Exit Function

                        End If

210                 Case Else
212                     UserCanAttackNpc.Result = eCanAttack
                        UserCanAttackNpc.CanAttack = True
                        Exit Function

                End Select

            End If

            If NpcList(NpcIndex).flags.team = 0 Then

                Dim CurrentOwnerIndex As Integer: CurrentOwnerIndex = GetOwnedBy(NpcIndex)

                If CurrentOwnerIndex <> 0 Then
                    If CurrentOwnerIndex <> UserIndex And IsValidNpcRef(UserList(CurrentOwnerIndex).flags.NPCAtacado) Then
                        If UserList(CurrentOwnerIndex).flags.NPCAtacado.ArrayIndex = NpcIndex And UserList(CurrentOwnerIndex).flags.Muerto = 0 And (Status(CurrentOwnerIndex) = Ciudadano Or Status(CurrentOwnerIndex) = Armada Or Status(CurrentOwnerIndex) = consejo) And (UserList(UserIndex).GuildIndex = 0 Or UserList(UserIndex).GuildIndex <> UserList(CurrentOwnerIndex).GuildIndex) And (UserList(UserIndex).Grupo.EnGrupo = False Or UserList(UserIndex).Grupo.Id <> UserList(CurrentOwnerIndex).Grupo.Id) Then
                            If UserList(UserIndex).flags.Seguro Then
                                UserCanAttackNpc.Result = eRemoveSafeCitizenNpc
                                Exit Function
                            Else
                                UserCanAttackNpc.TurnPK = True
                                UserCanAttackNpc.CanAttack = True
                                UserCanAttackNpc.Result = eAttackCitizenNpc
                                Exit Function

                            End If

                        End If

                    End If

                End If

            End If

        End If

        UserCanAttackNpc.CanAttack = True
        UserCanAttackNpc.Result = eCanAttack
        Exit Function
UserCanAttackNpc_Err:
222     Call TraceError(Err.Number, Err.Description, "Npcs.UserCanAttackNpc", Erl)

End Function

Public Function Inmovilize(ByVal SourceIndex As Integer, _
                           ByVal TargetIndex As Integer, _
                           ByVal Time As Integer, _
                           ByVal FX As Integer) As Boolean

        With NpcList(TargetIndex)
142         Call NPCAtacado(TargetIndex, SourceIndex)
172         .flags.Inmovilizado = 1
174         .Contadores.Inmovilizado = Time
176         .flags.Paralizado = 0
178         .Contadores.Paralisis = 0
180         Call AnimacionIdle(TargetIndex, True)
184         Call SendData(SendTarget.ToNPCAliveArea, TargetIndex, PrepareMessageCreateFX(.Char.charindex, FX, 0, .pos.x, .pos.y))
            Inmovilize = True

        End With

End Function

Public Function GetPhysicalDamageModifier(ByRef Npc As t_Npc) As Single
    GetPhysicalDamageModifier = max(1 + Npc.Modifiers.PhysicalDamageBonus, 0)

End Function

Public Function GetMagicDamageModifier(ByRef Npc As t_Npc) As Single
    GetMagicDamageModifier = max(1 + Npc.Modifiers.MagicDamageBonus, 0)

End Function

Public Function GetMagicDamageReduction(ByRef Npc As t_Npc) As Single
    GetMagicDamageReduction = max(1 - Npc.Modifiers.MagicDamageReduction, 0)

End Function

Public Function GetPhysicDamageReduction(ByRef Npc As t_Npc) As Single
    GetPhysicDamageReduction = max(1 - Npc.Modifiers.PhysicalDamageReduction, 0)

End Function

Public Function CanAttackUser(ByVal NpcIndex As Integer, _
                              ByVal UserIndex As Integer) As e_AttackInteractionResult

    With NpcList(NpcIndex)

        If Not IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eAttackUsers) Then
            CanAttackUser = eNotEnougthPrivileges
            Exit Function

        End If

        If UserList(UserIndex).Stats.MinHp = 0 Then
            CanAttackUser = eDeathTarget
            Exit Function

        End If

        If EsGM(UserIndex) Then
            If UserList(UserIndex).flags.EnConsulta Then
                CanAttackUser = eNotEnougthPrivileges
                Exit Function

            End If

            If Not UserList(UserIndex).flags.AdminPerseguible Then
                CanAttackUser = eNotEnougthPrivileges
                Exit Function

            End If

            If UserList(UserIndex).flags.invisible Then
                CanAttackUser = eNotEnougthPrivileges
                Exit Function

            End If

        End If

        Dim AttackerFaction As e_Facciones

        If IsValidUserRef(.MaestroUser) Then
            AttackerFaction = UserList(.MaestroUser.ArrayIndex).Faccion.Status
        Else
            AttackerFaction = .flags.Faccion

        End If

        If Not FactionCanAttackFaction(AttackerFaction, UserList(UserIndex).Faccion.Status) Then
            CanAttackUser = eSameFaction
            Exit Function

        End If

        If .flags.team <> 0 Then
            If .flags.team = UserList(UserIndex).flags.CurrentTeam Then
                CanAttackUser = eSameTeam
                Exit Function

            End If

        End If

    End With

    CanAttackUser = eCanAttack

End Function

Public Function CanHelpUser(ByVal NpcIndex As Integer, _
                            ByVal UserIndex As Integer) As e_InteractionResult

    With NpcList(NpcIndex)

        If Not IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eHelpUsers) Then
            CanHelpUser = eCantHelpUsers
            Exit Function

        End If

        Dim SourceFaction As e_Facciones

        If IsValidUserRef(.MaestroUser) Then
            SourceFaction = UserList(.MaestroUser.ArrayIndex).Faccion.Status
        Else
            SourceFaction = .flags.Faccion

        End If

        CanHelpUser = FactionCanHelpFaction(SourceFaction, UserList(UserIndex).Faccion.Status)

        If CanHelpUser <> e_InteractionResult.eInteractionOk Then
            Exit Function

        End If

        If UserList(UserIndex).flags.CurrentTeam > 0 Then

            Dim NpcTeam As Byte
            NpcTeam = .flags.team

            If NpcTeam <> UserList(UserIndex).flags.CurrentTeam Then
                CanHelpUser = eCantHelpUsers
                Exit Function

            End If

        End If

        CanHelpUser = eInteractionOk

    End With

End Function

Public Function CanHelpNpc(ByVal NpcIndex As Integer, _
                           ByVal TargetNPC As Integer) As e_InteractionResult
    CanHelpNpc = eInteractionOk

    If NpcList(NpcIndex).flags.team > 0 And NpcList(NpcIndex).flags.team <> NpcList(TargetNPC).flags.team Then
        CanHelpNpc = eOposingFaction
        Exit Function

    End If

End Function

Public Function CanAttackNpc(ByVal NpcIndex As Integer, _
                             ByVal TargetIndex As Integer) As e_AttackInteractionResult

    If NpcIndex = TargetIndex Then
        CanAttackNpc = eSameFaction
        Exit Function

    End If

    With NpcList(NpcIndex)

        If NpcList(TargetIndex).Attackable = 0 Then
            CanAttackNpc = eNotEnougthPrivileges
            Exit Function

        End If

        If Not NpcList(TargetIndex).flags.NPCActive Then
            CanAttackNpc = eNotEnougthPrivileges
            Exit Function

        End If

        If Not IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eAttackNpc) Then
            CanAttackNpc = eNotEnougthPrivileges
            Exit Function

        End If

        Dim TargetFaction         As e_Facciones
        Dim AttackerFaction       As e_Facciones
        Dim AttackerIsFreeCrature As Boolean
        Dim TargetIsFreeCrature   As Boolean

        If IsValidUserRef(NpcList(TargetIndex).MaestroUser) Then
            TargetFaction = UserList(NpcList(TargetIndex).MaestroUser.ArrayIndex).Faccion.Status
            TargetIsFreeCrature = False
        Else
            TargetFaction = NpcList(TargetIndex).flags.Faccion
            TargetIsFreeCrature = True

        End If

        If IsValidUserRef(.MaestroUser) Then
            AttackerFaction = UserList(.MaestroUser.ArrayIndex).Faccion.Status
            AttackerIsFreeCrature = False
        Else
            AttackerFaction = .flags.Faccion
            AttackerIsFreeCrature = True

        End If

        If Not FactionCanAttackFaction(AttackerFaction, TargetFaction) Then
            CanAttackNpc = eSameFaction
            Exit Function

        End If

        If .flags.team = NpcList(TargetIndex).flags.team Then
            CanAttackNpc = eSameTeam
            Exit Function

        End If

    End With

    CanAttackNpc = eCanAttack

End Function

Public Function GetEvasionBonus(ByRef Npc As t_Npc) As Integer
    GetEvasionBonus = Npc.Modifiers.EvasionBonus

End Function

Public Function GetHitBonus(ByRef Npc As t_Npc) As Integer
    GetHitBonus = Npc.Modifiers.HitBonus

End Function

'Defines the healing bonus when using a potion, a spell or any other healing source
Public Function GetSelfHealingBonus(ByRef Npc As t_Npc) As Single
    GetSelfHealingBonus = max(1 + Npc.Modifiers.SelfHealingBonus, 0)

End Function

'Defines bonus when healing someone with magic
Public Function GetMagicHealingBonus(ByRef Npc As t_Npc) As Single
    GetMagicHealingBonus = max(1 + Npc.Modifiers.MagicHealingBonus, 0)

End Function

Public Function CanSeeUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    CanSeeUser = UserMod.IsVisible(UserList(UserIndex))

End Function

Public Function CanPerformAttackAction(ByVal NpcIndex As Integer, _
                                       ByVal AttackInterval As Long)

    With NpcList(NpcIndex)
        CanPerformAttackAction = GlobalFrameTime - .Contadores.IntervaloLanzarHechizo > AttackInterval And GlobalFrameTime - .Contadores.IntervaloAtaque > AttackInterval

    End With

End Function

Public Function GetLinearDamageBonus(ByVal NpcIndex As Integer) As Integer
    GetLinearDamageBonus = NpcList(NpcIndex).Modifiers.PhysicalDamageLinearBonus

End Function

Public Sub SetBlockTileState(ByVal NpcIndex As Integer, ByVal Block As Boolean)

    Dim CurrentValue As Boolean

    With NpcList(NpcIndex)
        CurrentValue = IsSet(.flags.StatusMask, e_StatusMask.eDontBlockTile)

        If CurrentValue = Block Then Exit Sub
        If Block Then
            Call SetMask(.flags.StatusMask, e_StatusMask.eDontBlockTile)
        Else
            Call UnsetMask(.flags.StatusMask, e_StatusMask.eDontBlockTile)

        End If

        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareUpdateCharValue(.Char.charindex, e_CharValue.eDontBlockTile, Block))

    End With

End Sub

Public Function GetDefenseBonus(ByVal NpcIndex As Integer) As Integer
    GetDefenseBonus = NpcList(NpcIndex).Modifiers.DefenseBonus

End Function

' Retorna el usuario que esta atacando al NPC actualmente (medido con tiempo)
Public Function GetOwnedBy(ByVal NpcIndex As Integer) As Integer
    GetOwnedBy = 0

    With NpcList(NpcIndex).flags

        If .AttackedBy = vbNullString Then Exit Function
        If GlobalFrameTime - .AttackedTime > IntervaloNpcOwner Then Exit Function

        Dim Attacker As t_UserReference: Attacker = NameIndex(.AttackedBy)

        If Not IsValidUserRef(Attacker) Then Exit Function
        GetOwnedBy = Attacker.ArrayIndex

    End With

End Function

' Retorna si un NPC puede atacar un usuario diferente al que lo esta atacando
Public Function CanAttackNotOwner(ByVal NpcIndex As Integer, _
                                  ByVal UserIndex As Integer) As Boolean

    Dim AttackResult As t_AttackInteractionResult
    AttackResult = UserCanAttackNpc(UserIndex, NpcIndex)

    ' Si el usuario puede atacar al NPC
    If AttackResult.CanAttack Then
        ' Lo atacamos solo si puede atacar sin hacerse PK (no lo forzamos a hacerse PK)
        CanAttackNotOwner = Not AttackResult.TurnPK
    Else
        ' En caso que el usuario no pueda atacar al NPC, este debe ignorarlo a el
        ' Excepto que no pueda atacar por los siguientes motivos: esta montado, esta fuera de su campo de vision
        CanAttackNotOwner = AttackResult.Result = eMounted Or AttackResult.Result = eOutOfRange

    End If

End Function
