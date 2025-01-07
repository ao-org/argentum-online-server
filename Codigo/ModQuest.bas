Attribute VB_Name = "ModQuest"
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
Option Explicit
 
'Constantes de las quests
Public Const MAXUSERQUESTS As Integer = 5     'Maxima cantidad de quests que puede tener un usuario al mismo tiempo.

Public Function TieneQuest(ByVal UserIndex As Integer, ByVal QuestNumber As Integer) As Byte
        
        On Error GoTo TieneQuest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Devuelve el slot de UserQuests en que tiene la quest QuestNumber. En caso contrario devuelve 0.
        'Last modified: 27/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i As Integer
 
100     For i = 1 To MAXUSERQUESTS

102         If UserList(UserIndex).QuestStats.Quests(i).QuestIndex = QuestNumber Then
104             TieneQuest = i
                Exit Function

            End If

106     Next i
    
108     TieneQuest = 0

        
        Exit Function

TieneQuest_Err:
110     Call TraceError(Err.Number, Err.Description, "ModQuest.TieneQuest", Erl)

        
End Function
 
Public Function FreeQuestSlot(ByVal UserIndex As Integer) As Byte
        
        On Error GoTo FreeQuestSlot_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Devuelve el proximo slot de quest libre.
        'Last modified: 27/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i As Integer
 
100     For i = 1 To MAXUSERQUESTS

102         If UserList(UserIndex).QuestStats.Quests(i).QuestIndex = 0 Then
104             FreeQuestSlot = i
                Exit Function

            End If

106     Next i
    
108     FreeQuestSlot = 0

        
        Exit Function

FreeQuestSlot_Err:
110     Call TraceError(Err.Number, Err.Description, "ModQuest.FreeQuestSlot", Erl)

        
End Function
 
Public Sub FinishQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte)
        On Error GoTo FinishQuest_Err
        'Maneja el evento de terminar una quest.
        Dim i, j           As Integer
        Dim InvSlotsLibres As Byte
        Dim NpcIndex       As Integer
100     NpcIndex = UserList(UserIndex).flags.TargetNPC.ArrayIndex
102     With QuestList(QuestIndex)
            'Comprobamos que tenga los objetos.
104         If .RequiredOBJs > 0 Then
106             For i = 1 To .RequiredOBJs
108                 If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).amount, UserIndex) = False Then
110                     Call WriteChatOverHead(UserIndex, "No has conseguido todos los objetos que te he pedido.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
                        Exit Sub
                    End If
112             Next i
            End If
            'Comprobamos que haya matado todas las criaturas.
114         If .RequiredNPCs > 0 Then
116             For i = 1 To .RequiredNPCs
118                 If .RequiredNPC(i).amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
120                     Call WriteChatOverHead(UserIndex, "No has matado todas las criaturas que te he pedido.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
                        Exit Sub
                    End If
122             Next i
            End If
            If .RequiredSpellCount > 0 Then
                For i = 1 To .RequiredSpellCount
                    If Not UserHasSpell(UserIndex, .RequiredSpellList(i)) Then
                        Call WriteLocaleChatOverHead(UserIndex, MsgRequiredSpell, Hechizos(.RequiredSpellList(i)).nombre, NpcList(NpcIndex).Char.charindex, vbYellow)
                        Exit Sub
                    End If
                Next i
            End If
            'Comprobamos que haya targeteado todos los npc
124          If .RequiredTargetNPCs > 0 Then
126              For i = 1 To .RequiredTargetNPCs
128                  If .RequiredTargetNPC(i).amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsTarget(i) Then
130                      Call WriteChatOverHead(UserIndex, "No has visitado al npc que te pedi.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
                        Exit Sub
                        End If
132              Next i
            End If
            'Check required skill
            If .RequiredSkill.SkillType > 0 Then
                If UserList(UserIndex).Stats.UserSkills(.RequiredSkill.SkillType) < .RequiredSkill.RequiredValue Then
                    Call WriteLocaleChatOverHead(UserIndex, MsgRequiredSkill, SkillsNames(.RequiredSkill.SkillType), NpcList(NpcIndex).Char.charindex, vbYellow)
                    Exit Sub
                End If
            End If
            'Comprobamos que el usuario tenga espacio para recibir los items.
134         If .RewardOBJs > 0 Then
                'Buscamos la cantidad de slots de inventario libres.
136             For i = 1 To UserList(UserIndex).CurrentInventorySlots
138                 If UserList(UserIndex).Invent.Object(i).ObjIndex = 0 Then InvSlotsLibres = InvSlotsLibres + 1
140             Next i
                'Nos fijamos si entra
142             If InvSlotsLibres < .RewardOBJs Then
144                 Call WriteChatOverHead(UserIndex, "No tienes suficiente espacio en el inventario para recibir la recompensa. Vuelve cuando hayas hecho mas espacio.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
                    Exit Sub
                End If
            End If
            
            Dim KnownSkills As Integer
            If .RewardSpellCount > 0 Then
                For i = 1 To .RewardSpellCount
                    For j = 1 To UBound(UserList(UserIndex).Stats.UserHechizos)
                        If UserList(UserIndex).Stats.UserHechizos(j) = .RewardSpellList(i) Then
                            KnownSkills = KnownSkills + 1
                        End If
                    Next j
                Next i
                If KnownSkills = .RewardSpellCount Then
                    Call WriteLocaleChatOverHead(UserIndex, MsgSkillAlreadyKnown, vbNullString, NpcList(npcIndex).Char.charindex, vbYellow)
                    Exit Sub
                End If
            End If
    
            'A esta altura ya cumplio los objetivos, entonces se le entregan las recompensas.
146         Call WriteChatOverHead(UserIndex, "QUESTFIN*" & QuestIndex, NpcList(NpcIndex).Char.CharIndex, vbYellow)

            'Si la quest pedia objetos, se los saca al personaje.
148         If .RequiredOBJs Then
150             For i = 1 To .RequiredOBJs
152                 Call QuitarObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).amount, UserIndex)
154             Next i
            End If
            
            If .RequiredSpellCount > 0 Then
                For i = 1 To .RequiredSpellCount
                    For j = 1 To UBound(UserList(UserIndex).Stats.UserHechizos)
                        If UserList(UserIndex).Stats.UserHechizos(j) = .RequiredSpellList(i) Then
                            UserList(UserIndex).Stats.UserHechizos(j) = 0
                            Call UpdateUserHechizos(False, UserIndex, CByte(j))
                        End If
                    Next j
                Next i
                UserList(UserIndex).flags.ModificoHechizos = True
            End If
        
            'Se entrega la experiencia.
156         If .RewardEXP Then
158             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
160                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + (.RewardEXP * SvrConfig.GetValue("ExpMult"))
162                 Call WriteUpdateExp(UserIndex)
164                 Call CheckUserLevel(UserIndex)
166                 Call WriteLocaleMsg(UserIndex, "140", e_FontTypeNames.FONTTYPE_EXP, (.RewardEXP * SvrConfig.GetValue("ExpMult")))
                Else
                    'Msg1314= No se te ha dado experiencia porque eres nivel máximo.
                    Call WriteLocaleMsg(UserIndex, "1314", e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        
            'Se entrega el oro.
170         If .RewardGLD > 0 Then
                Dim GiveGLD As Long
                GiveGLD = (.RewardGLD * SvrConfig.GetValue("GoldMult"))
                If GiveGLD < 100000 Then
172                 UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + GiveGLD
                    'Msg1315= Has ganado ¬1 monedas de oro como recompensa.
                    Call WriteLocaleMsg(UserIndex, "1315", e_FontTypeNames.FONTTYPE_INFOIAO, PonerPuntos(GiveGLD))
176                 Call WriteUpdateGold(UserIndex)
                Else
                    UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + GiveGLD
                    'Msg1316= Has ganado ¬1 monedas de oro como recompensa. La recompensa ha sido depositada en su cuenta del Banco Goliath.
                    Call WriteLocaleMsg(UserIndex, "1316", e_FontTypeNames.FONTTYPE_INFOIAO, PonerPuntos(GiveGLD))
                End If
            End If
        
            'Si hay recompensa de objetos, se entregan.
178         If .RewardOBJs > 0 Then
180             For i = 1 To .RewardOBJs
182                 If .RewardOBJ(i).amount Then
184                     Call MeterItemEnInventario(UserIndex, .RewardOBJ(i))
186                     'Msg1318=Has recibido ¬1 como recompensa.
                        Call WriteLocaleMsg(UserIndex, "1318", e_FontTypeNames.FONTTYPE_FIGHT, QuestList(QuestIndex).RewardOBJ(i).amount & " " & ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).name)
                    End If
188             Next i
            End If
            
            If .RewardSpellCount > 0 Then
                For i = 1 To .RewardSpellCount
                    If Not TieneHechizo(.RewardSpellList(i), UserIndex) Then
                        'Buscamos un slot vacio
204                     For j = 1 To MAXUSERHECHIZOS
206                          If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
208                     Next j
210                     If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
                            'Msg1317= No tenes espacio para mas hechizos.
                            Call WriteLocaleMsg(UserIndex, "1317", e_FontTypeNames.FONTTYPE_INFO)
                        Else
214                         UserList(UserIndex).Stats.UserHechizos(j) = .RewardSpellList(i)
216                         Call UpdateUserHechizos(False, UserIndex, CByte(j))
                        End If
                        UserList(UserIndex).flags.ModificoHechizos = True
                    End If
                Next i
            End If
    
            'Actualizamos el personaje
290         Call UpdateUserInv(True, UserIndex, 0)
    
            'Limpiamos el slot de quest.
292         Call CleanQuestSlot(UserIndex, QuestSlot)
        
            'Ordenamos las quests
294         Call ArrangeUserQuests(UserIndex)

            'Se agrega que el usuario ya hizo esta quest. - WyroX: La agrego aunque sea repetible, para llevar el control
298         Call AddDoneQuest(UserIndex, QuestIndex)

300         If .Repetible = 0 Then
                Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 2)
            Else
                Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 1)
            End If
        End With
        Exit Sub

FinishQuest_Err:
    Call TraceError(Err.Number, Err.Description, "ModQuest.FinishQuest", Erl)
End Sub
 
Public Sub AddDoneQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer)
        
        On Error GoTo AddDoneQuest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Agrega la quest QuestIndex a la lista de quests hechas.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
100     With UserList(UserIndex).QuestStats
102         .NumQuestsDone = .NumQuestsDone + 1
104         ReDim Preserve .QuestsDone(1 To .NumQuestsDone)
106         .QuestsDone(.NumQuestsDone) = QuestIndex
            
        End With

        
        Exit Sub

AddDoneQuest_Err:
108     Call TraceError(Err.Number, Err.Description, "ModQuest.AddDoneQuest", Erl)

        
End Sub
 
Public Function UserDoneQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer) As Boolean
        
        On Error GoTo UserDoneQuest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Verifica si el usuario hizo la quest QuestIndex.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i As Integer
100     If QuestIndex = 0 Then
102         UserDoneQuest = True
            Exit Function
        End If
            
            

104     With UserList(UserIndex).QuestStats

106         If .NumQuestsDone Then

108             For i = 1 To .NumQuestsDone

110                 If .QuestsDone(i) = QuestIndex Then
112                     UserDoneQuest = True
                        Exit Function

                    End If

114             Next i

            End If

        End With
    
116     UserDoneQuest = False
        
        
        Exit Function

UserDoneQuest_Err:
118     Call TraceError(Err.Number, Err.Description, "ModQuest.UserDoneQuest", Erl)

        
End Function
 
Public Sub CleanQuestSlot(ByVal UserIndex As Integer, ByVal QuestSlot As Integer)
        
        On Error GoTo CleanQuestSlot_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Limpia un slot de quest de un usuario.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i As Integer
 
100     With UserList(UserIndex).QuestStats.Quests(QuestSlot)

102         If .QuestIndex Then

116             If QuestList(.QuestIndex).RequiredNPCs Then

118                 For i = 1 To QuestList(.QuestIndex).RequiredNPCs
120                     .NPCsKilled(i) = 0
122                 Next i

                End If
                
124           If QuestList(.QuestIndex).RequiredTargetNPCs Then

126              For i = 1 To QuestList(.QuestIndex).RequiredTargetNPCs
128                  .NPCsTarget(i) = 0
130              Next i

                End If

            End If

132         .QuestIndex = 0
            
            UserList(UserIndex).flags.ModificoQuests = True
        End With

        
        Exit Sub

CleanQuestSlot_Err:
134     Call TraceError(Err.Number, Err.Description, "ModQuest.CleanQuestSlot", Erl)

        
End Sub
 
Public Sub ResetQuestStats(ByVal UserIndex As Integer)
        
        On Error GoTo ResetQuestStats_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Limpia todos los QuestStats de un usuario
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i As Integer
 
100     For i = 1 To MAXUSERQUESTS
102         Call CleanQuestSlot(UserIndex, i)
104     Next i
    
106     With UserList(UserIndex).QuestStats
108         .NumQuestsDone = 0
110         Erase .QuestsDone

        End With

        
        Exit Sub

ResetQuestStats_Err:
112     Call TraceError(Err.Number, Err.Description, "ModQuest.ResetQuestStats", Erl)

        
End Sub
 
Public Sub LoadQuests()

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Carga el archivo QUESTS.DAT en el array QuestList.
        'Last modified: 27/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        On Error GoTo ErrorHandler

        Dim Reader    As clsIniManager

        Dim NumQuests As Integer

        Dim tmpStr    As String

        Dim i         As Integer

        Dim j         As Integer
    
        'Cargamos el clsIniManager en memoria
100     Set Reader = New clsIniManager
    
        'Lo inicializamos para el archivo Quests.DAT
102     Call Reader.Initialize(DatPath & "Quests.DAT")
    
        'Redimensionamos el array
104     NumQuests = Reader.GetValue("INIT", "NumQuests")
106     ReDim QuestList(1 To NumQuests)
    
        'Cargamos los datos
108     For i = 1 To NumQuests

110         With QuestList(i)
112             .nombre = Reader.GetValue("QUEST" & i, "Nombre")
114             .Desc = Reader.GetValue("QUEST" & i, "Desc")
116             .RequiredLevel = val(Reader.GetValue("QUEST" & i, "RequiredLevel"))
                .RequiredClass = val(Reader.GetValue("QUEST" & i, "RequiredClass"))
118             .RequiredQuest = val(Reader.GetValue("QUEST" & i, "RequiredQuest"))
            
120             .DescFinal = Reader.GetValue("QUEST" & i, "DescFinal")
            
122             .NextQuest = Reader.GetValue("QUEST" & i, "NextQuest")
            
                'CARGAMOS OBJETOS REQUERIDOS
124             .RequiredOBJs = val(Reader.GetValue("QUEST" & i, "RequiredOBJs"))
125             .Trabajador = IIf(val(Reader.GetValue("QUEST" & i, "Trabajador")) = 1, True, False)
123             .TalkTo = val(Reader.GetValue("QUEST" & i, "TalkTo"))

126             If .RequiredOBJs > 0 Then
128                 ReDim .RequiredOBJ(1 To .RequiredOBJs)

130                 For j = 1 To .RequiredOBJs
132                     tmpStr = Reader.GetValue("QUEST" & i, "RequiredOBJ" & j)
                    
134                     .RequiredOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
136                     .RequiredOBJ(j).amount = val(ReadField(2, tmpStr, 45))
138                 Next j

                End If
                .RequiredSpellCount = val(Reader.GetValue("QUEST" & i, "RequiredSpellCount"))
                If .RequiredSpellCount > 0 Then
                    ReDim .RequiredSpellList(1 To .RequiredSpellCount) As Integer
                    For j = 1 To .RequiredSpellCount
                        .RequiredSpellList(j) = val(Reader.GetValue("QUEST" & i, "RequiredSpell" & j))
                    Next j
                End If
    
            
                'CARGAMOS NPCS REQUERIDOS
140             .RequiredNPCs = val(Reader.GetValue("QUEST" & i, "RequiredNPCs"))

142             If .RequiredNPCs > 0 Then
144                 ReDim .RequiredNPC(1 To .RequiredNPCs)

146                 For j = 1 To .RequiredNPCs
148                     tmpStr = Reader.GetValue("QUEST" & i, "RequiredNPC" & j)
                    
150                     .RequiredNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
152                     .RequiredNPC(j).amount = val(ReadField(2, tmpStr, 45))
154                 Next j

                End If
                .RequiredSkill.SkillType = val(Reader.GetValue("QUEST" & i, "RequiredSkill"))
                .RequiredSkill.RequiredValue = val(Reader.GetValue("QUEST" & i, "RequiredValue"))
                
            
                'CARGAMOS NPCS TARGET REQUERIDOS
156             .RequiredTargetNPCs = val(Reader.GetValue("QUEST" & i, "RequiredTargetNPCs"))

158             If .RequiredTargetNPCs > 0 Then
160                 ReDim .RequiredTargetNPC(1 To .RequiredTargetNPCs)

162                 For j = 1 To .RequiredTargetNPCs
164                     tmpStr = Reader.GetValue("QUEST" & i, "RequiredTargetNPC" & j)
                    
166                     .RequiredTargetNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
168                     .RequiredTargetNPC(j).amount = 1
170                 Next j

                End If

172             .RewardGLD = val(Reader.GetValue("QUEST" & i, "RewardGLD"))
174             .RewardEXP = val(Reader.GetValue("QUEST" & i, "RewardEXP"))
176             .Repetible = val(Reader.GetValue("QUEST" & i, "Repetible"))
            
                'CARGAMOS OBJETOS DE RECOMPENSA
178             .RewardOBJs = val(Reader.GetValue("QUEST" & i, "RewardOBJs"))
180             If .RewardOBJs > 0 Then
182                 ReDim .RewardOBJ(1 To .RewardOBJs)
184                 For j = 1 To .RewardOBJs
186                     tmpStr = Reader.GetValue("QUEST" & i, "RewardOBJ" & j)
188                     .RewardOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
190                     .RewardOBJ(j).amount = val(ReadField(2, tmpStr, 45))
192                 Next j
                End If
                .RewardSpellCount = val(Reader.GetValue("QUEST" & i, "RewardSkills"))
200             If .RewardSpellCount > 0 Then
202                 ReDim .RewardSpellList(1 To .RewardSpellCount)
204                 For j = 1 To .RewardSpellCount
206                     .RewardSpellList(j) = val(Reader.GetValue("QUEST" & i, "RewardSkill" & j))
208                 Next j
                End If
            End With
220     Next i
    
        'Eliminamos la clase
230     Set Reader = Nothing
        Exit Sub
                    
ErrorHandler:
    MsgBox "Error cargando el archivo QUESTS.DAT.", vbOKOnly + vbCritical
End Sub
 
Public Sub ArrangeUserQuests(ByVal UserIndex As Integer)
        
        On Error GoTo ArrangeUserQuests_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Ordena las quests del usuario de manera que queden todas al principio del arreglo.
        'Last modified: 30/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i As Integer

        Dim j As Integer
 
100     With UserList(UserIndex).QuestStats

102         For i = 1 To MAXUSERQUESTS - 1

104             If .Quests(i).QuestIndex = 0 Then

106                 For j = i + 1 To MAXUSERQUESTS

108                     If .Quests(j).QuestIndex Then
110                         .Quests(i) = .Quests(j)
112                         Call CleanQuestSlot(UserIndex, j)
                            Exit For

                        End If

114                 Next j

                End If

116         Next i

        End With

        
        Exit Sub

ArrangeUserQuests_Err:
118     Call TraceError(Err.Number, Err.Description, "ModQuest.ArrangeUserQuests", Erl)

        
End Sub
 
Public Sub EnviarQuest(ByVal UserIndex As Integer)
        On Error GoTo EnviarQuest_Err
        Dim NpcIndex As Integer
        Dim tmpByte  As Byte
        
100     If Not IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then Exit Sub
102     NpcIndex = UserList(UserIndex).flags.TargetNPC.ArrayIndex
        'Esta el personaje en la distancia correcta?
104     If Distancia(UserList(UserIndex).Pos, NpcList(NpcIndex).Pos) > 5 Then
106         ' Msg8=Estas demasiado lejos.
            Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        'El NPC hace quests?
108     If NpcList(NpcIndex).NumQuest = 0 Then
110         Call WriteChatOverHead(UserIndex, "No tengo ninguna misión para ti.", NpcList(NpcIndex).Char.charindex, vbYellow)
            Exit Sub
        End If
        
        'Hago un for para chequear si alguna de las misiones que da el NPC ya se completo.
        Dim q As Byte
        Dim i As Long, j As Long
        For i = 1 To UBound(QuestList)
            If QuestList(i).TalkTo > 0 And QuestList(i).TalkTo = NpcList(NpcIndex).Numero Then
                tmpByte = TieneQuest(UserIndex, i)
                If tmpByte > 0 Then
                    For j = 1 To MAXUSERQUESTS
                         If FinishQuestCheck(UserIndex, i, tmpByte) Then
111                         Call FinishQuest(UserIndex, i, tmpByte)
                            Exit Sub
                        End If
                    Next j
                End If
            End If
        Next i
112     For q = 1 To NpcList(NpcIndex).NumQuest
114         tmpByte = TieneQuest(UserIndex, NpcList(NpcIndex).QuestNumber(q))
116         If tmpByte Then
                'El usuario esta haciendo la quest, entonces va a hablar con el NPC para recibir la recompensa.
118             If FinishQuestCheck(UserIndex, NpcList(NpcIndex).QuestNumber(q), tmpByte) Then
120                 Call FinishQuest(UserIndex, NpcList(NpcIndex).QuestNumber(q), tmpByte)
                    Exit Sub
                End If
            End If
122     Next q
      
124   Call WriteNpcQuestListSend(UserIndex, NpcIndex)
      Exit Sub

EnviarQuest_Err:
126     Call TraceError(Err.Number, Err.Description, "ModQuest.EnviarQuest", Erl)

        
End Sub

Public Function FinishQuestCheck(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte) As Boolean
        On Error GoTo FinishQuestCheck_Err
        Dim i              As Integer
        Dim InvSlotsLibres As Byte
        Dim NpcIndex       As Integer
100     NpcIndex = UserList(UserIndex).flags.TargetNPC.ArrayIndex
102     With QuestList(QuestIndex)
            'Comprobamos que tenga los objetos.
104         If .RequiredOBJs > 0 Then
106             For i = 1 To .RequiredOBJs
108                 If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).amount, UserIndex) = False Then
110                     FinishQuestCheck = False
                        Exit Function
                    End If
112             Next i
            End If
            'Comprobamos que haya matado todas las criaturas.
114         If .RequiredNPCs > 0 Then
116             For i = 1 To .RequiredNPCs
118                 If .RequiredNPC(i).amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
120                     FinishQuestCheck = False
                        Exit Function
                    End If
122             Next i
            End If
            'Check required spells
            If .RequiredSpellCount > 0 Then
                For i = 1 To .RequiredSpellCount
                    If Not UserHasSpell(UserIndex, .RequiredSpellList(i)) Then
                        FinishQuestCheck = False
                        Exit Function
                    End If
                Next i
            End If
            'Check required skill
            If .RequiredSkill.SkillType > 0 Then
                If UserList(UserIndex).Stats.UserSkills(.RequiredSkill.SkillType) < .RequiredSkill.RequiredValue Then
                    FinishQuestCheck = False
                    Exit Function
                End If
            End If
            'Comprobamos que haya targeteado todas las criaturas.
124         If .RequiredTargetNPCs > 0 Then
126             For i = 1 To .RequiredTargetNPCs
128                 If .RequiredTargetNPC(i).amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsTarget(i) Then
130                     FinishQuestCheck = False
                        Exit Function
                    End If
132             Next i
            End If
        End With
134     FinishQuestCheck = True
        Exit Function
FinishQuestCheck_Err:
136     Call TraceError(Err.Number, Err.Description, "ModQuest.FinishQuestCheck", Erl)
End Function

Function FaltanItemsQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal ObjIndex As Integer) As Boolean

        On Error GoTo Handler

100     With QuestList(QuestIndex)

            ' Por las dudas...
102         If .RequiredOBJs > 0 Then
        
                Dim i As Integer
        
104             For i = 1 To .RequiredOBJs
            
                    ' Encontramos el objeto
106                 If ObjIndex = .RequiredOBJ(i).ObjIndex Then

                        ' Devolvemos si ya tiene todos los que la quest pide
108                     FaltanItemsQuest = Not TieneObjetos(ObjIndex, .RequiredOBJ(i).amount, UserIndex)
                        Exit Function

                    End If
            
110             Next i
        
            End If

        End With
        
        Exit Function
            
Handler:
112     Call TraceError(Err.Number, Err.Description, "ModQuest.FaltanItemsQuest", Erl)


End Function
