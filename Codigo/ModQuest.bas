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
Public Function TieneQuest(ByVal UserIndex As Integer, ByVal QuestNumber As Integer) As Byte
    On Error GoTo TieneQuest_Err
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Devuelve el slot de UserQuests en que tiene la quest QuestNumber. En caso contrario devuelve 0.
    'Last modified: 27/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
    For i = 1 To MAXUSERQUESTS
        If UserList(UserIndex).QuestStats.Quests(i).QuestIndex = QuestNumber Then
            TieneQuest = i
            Exit Function
        End If
    Next i
    TieneQuest = 0
    Exit Function
TieneQuest_Err:
    Call TraceError(Err.Number, Err.Description, "ModQuest.TieneQuest", Erl)
End Function
 
Public Function FreeQuestSlot(ByVal UserIndex As Integer) As Byte
    On Error GoTo FreeQuestSlot_Err
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Devuelve el proximo slot de quest libre.
    'Last modified: 27/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
    For i = 1 To MAXUSERQUESTS
        If UserList(UserIndex).QuestStats.Quests(i).QuestIndex = 0 Then
            FreeQuestSlot = i
            Exit Function
        End If
    Next i
    FreeQuestSlot = 0
    Exit Function
FreeQuestSlot_Err:
    Call TraceError(Err.Number, Err.Description, "ModQuest.FreeQuestSlot", Erl)
End Function
 
Public Sub FinishQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte)
    On Error GoTo FinishQuest_Err
    'Maneja el evento de terminar una quest.
    Dim i, j           As Integer
    Dim InvSlotsLibres As Byte
    Dim NpcIndex       As Integer
    NpcIndex = UserList(UserIndex).flags.TargetNPC.ArrayIndex
    With QuestList(QuestIndex)
        'Comprobamos que tenga los objetos.
        If .RequiredOBJs > 0 Then
            For i = 1 To .RequiredOBJs
                If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).amount, UserIndex) = False Then
                    Call WriteLocaleChatOverHead(UserIndex, "1336", "", NpcList(NpcIndex).Char.charindex, vbYellow) ' Msg1336=No has conseguido todos los objetos que te he pedido.
                    Exit Sub
                End If
            Next i
        End If
        'Comprobamos que haya matado todas las criaturas.
        If .RequiredNPCs > 0 Then
            For i = 1 To .RequiredNPCs
                If .RequiredNPC(i).amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
                    Call WriteLocaleChatOverHead(UserIndex, "1337", "", NpcList(NpcIndex).Char.charindex, vbYellow) ' Msg1337=No has matado todas las criaturas que te he pedido.
                    Exit Sub
                End If
            Next i
        End If
        If .RequiredSpellCount > 0 Then
            For i = 1 To .RequiredSpellCount
                If Not UserHasSpell(UserIndex, .RequiredSpellList(i)) Then
                    Call WriteLocaleChatOverHead(UserIndex, "1338", Hechizos(.RequiredSpellList(i)).nombre, NpcList(NpcIndex).Char.charindex, vbYellow) ' Msg1338=Necesitas aprender el hechizo: {0}.
                    Exit Sub
                End If
            Next i
        End If
        'Comprobamos que haya targeteado todos los npc
        If .RequiredTargetNPCs > 0 Then
            For i = 1 To .RequiredTargetNPCs
                If .RequiredTargetNPC(i).amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsTarget(i) Then
                    Call WriteLocaleChatOverHead(UserIndex, "1339", "", NpcList(NpcIndex).Char.charindex, vbYellow) ' Msg1339=No has visitado al npc que te pedi.
                    Exit Sub
                End If
            Next i
        End If
        'Check required skill
        If .RequiredSkill.SkillType > 0 Then
            If UserList(UserIndex).Stats.UserSkills(.RequiredSkill.SkillType) < .RequiredSkill.RequiredValue Then
                Call WriteLocaleChatOverHead(UserIndex, MsgRequiredSkill, SkillsNames(.RequiredSkill.SkillType), NpcList(NpcIndex).Char.charindex, vbYellow)
                Exit Sub
            End If
        End If
        'Comprobamos que el usuario tenga espacio para recibir los items.
        If .RewardOBJs > 0 Then
            'Buscamos la cantidad de slots de inventario libres.
            For i = 1 To UserList(UserIndex).CurrentInventorySlots
                If UserList(UserIndex).invent.Object(i).ObjIndex = 0 Then InvSlotsLibres = InvSlotsLibres + 1
            Next i
            'Nos fijamos si entra
            If InvSlotsLibres < .RewardOBJs Then
                Call WriteLocaleChatOverHead(UserIndex, "1340", "", NpcList(NpcIndex).Char.charindex, vbYellow) ' Msg1340=No tienes suficiente espacio en el inventario para recibir la recompensa. Vuelve cuando hayas hecho mas espacio.
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
                Call WriteLocaleChatOverHead(UserIndex, MsgSkillAlreadyKnown, vbNullString, NpcList(NpcIndex).Char.charindex, vbYellow)
                Exit Sub
            End If
        End If
        'A esta altura ya cumplio los objetivos, entonces se le entregan las recompensas.
        Call WriteChatOverHead(UserIndex, "QUESTFIN*" & QuestIndex, NpcList(NpcIndex).Char.charindex, vbYellow)
        'Si la quest pedia objetos, se los saca al personaje.
        If .RequiredOBJs Then
            For i = 1 To .RequiredOBJs
                Call QuitarObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).amount, UserIndex)
            Next i
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
        If .RewardEXP Then
            If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + (.RewardEXP * SvrConfig.GetValue("ExpMult"))
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)
                Call WriteLocaleMsg(UserIndex, "140", e_FontTypeNames.FONTTYPE_EXP, (.RewardEXP * SvrConfig.GetValue("ExpMult")))
            Else
                'Msg1314= No se te ha dado experiencia porque eres nivel máximo.
                Call WriteLocaleMsg(UserIndex, "1314", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End If
        'Se entrega el oro.
        If .RewardGLD > 0 Then
            Dim GiveGLD As Long
            GiveGLD = (.RewardGLD * SvrConfig.GetValue("GoldMult"))
            If GiveGLD < 100000 Then
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + GiveGLD
                'Msg1315= Has ganado ¬1 monedas de oro como recompensa.
                Call WriteLocaleMsg(UserIndex, "1315", e_FontTypeNames.FONTTYPE_INFOIAO, PonerPuntos(GiveGLD))
                Call WriteUpdateGold(UserIndex)
            Else
                UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + GiveGLD
                'Msg1316= Has ganado ¬1 monedas de oro como recompensa. La recompensa ha sido depositada en su cuenta del Banco Goliath.
                Call WriteLocaleMsg(UserIndex, "1316", e_FontTypeNames.FONTTYPE_INFOIAO, PonerPuntos(GiveGLD))
            End If
        End If
        'Si hay recompensa de objetos, se entregan.
        If .RewardOBJs > 0 Then
            For i = 1 To .RewardOBJs
                If .RewardOBJ(i).amount Then
                    Call MeterItemEnInventario(UserIndex, .RewardOBJ(i))
                    'Msg1318=Has recibido ¬1 como recompensa.
                    Call WriteLocaleMsg(UserIndex, "1318", e_FontTypeNames.FONTTYPE_FIGHT, QuestList(QuestIndex).RewardOBJ(i).amount & " " & ObjData(QuestList( _
                            QuestIndex).RewardOBJ(i).ObjIndex).name)
                End If
            Next i
        End If
        If .RewardSpellCount > 0 Then
            For i = 1 To .RewardSpellCount
                If Not TieneHechizo(.RewardSpellList(i), UserIndex) Then
                    'Buscamos un slot vacio
                    For j = 1 To MAXUSERHECHIZOS
                        If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
                    Next j
                    If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
                        'Msg1317= No tenes espacio para mas hechizos.
                        Call WriteLocaleMsg(UserIndex, "1317", e_FontTypeNames.FONTTYPE_INFO)
                    Else
                        UserList(UserIndex).Stats.UserHechizos(j) = .RewardSpellList(i)
                        Call UpdateUserHechizos(False, UserIndex, CByte(j))
                    End If
                    UserList(UserIndex).flags.ModificoHechizos = True
                End If
            Next i
        End If
        'Actualizamos el personaje
        Call UpdateUserInv(True, UserIndex, 0)
        'Limpiamos el slot de quest.
        Call CleanQuestSlot(UserIndex, QuestSlot)
        'Ordenamos las quests
        Call ArrangeUserQuests(UserIndex)
        'Se agrega que el usuario ya hizo esta quest. -  La agrego aunque sea repetible, para llevar el control
        Call AddDoneQuest(UserIndex, QuestIndex)
        If .Repetible = 0 Then
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
    With UserList(UserIndex).QuestStats
        .NumQuestsDone = .NumQuestsDone + 1
        ReDim Preserve .QuestsDone(1 To .NumQuestsDone)
        .QuestsDone(.NumQuestsDone) = QuestIndex
    End With
    Exit Sub
AddDoneQuest_Err:
    Call TraceError(Err.Number, Err.Description, "ModQuest.AddDoneQuest", Erl)
End Sub
 
Public Function UserDoneQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer) As Boolean
    On Error GoTo UserDoneQuest_Err
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Verifica si el usuario hizo la quest QuestIndex.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
    If QuestIndex = 0 Then
        UserDoneQuest = True
        Exit Function
    End If
    With UserList(UserIndex).QuestStats
        If .NumQuestsDone Then
            For i = 1 To .NumQuestsDone
                If .QuestsDone(i) = QuestIndex Then
                    UserDoneQuest = True
                    Exit Function
                End If
            Next i
        End If
    End With
    UserDoneQuest = False
    Exit Function
UserDoneQuest_Err:
    Call TraceError(Err.Number, Err.Description, "ModQuest.UserDoneQuest", Erl)
End Function
 
Public Sub CleanQuestSlot(ByVal UserIndex As Integer, ByVal QuestSlot As Integer)
    On Error GoTo CleanQuestSlot_Err
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Limpia un slot de quest de un usuario.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
    With UserList(UserIndex).QuestStats.Quests(QuestSlot)
        If .QuestIndex Then
            If QuestList(.QuestIndex).RequiredNPCs Then
                For i = 1 To QuestList(.QuestIndex).RequiredNPCs
                    .NPCsKilled(i) = 0
                Next i
            End If
            If QuestList(.QuestIndex).RequiredTargetNPCs Then
                For i = 1 To QuestList(.QuestIndex).RequiredTargetNPCs
                    .NPCsTarget(i) = 0
                Next i
            End If
        End If
        .QuestIndex = 0
        UserList(UserIndex).flags.ModificoQuests = True
    End With
    Exit Sub
CleanQuestSlot_Err:
    Call TraceError(Err.Number, Err.Description, "ModQuest.CleanQuestSlot", Erl)
End Sub
 
Public Sub ResetQuestStats(ByVal UserIndex As Integer)
    On Error GoTo ResetQuestStats_Err
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Limpia todos los QuestStats de un usuario
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer
    For i = 1 To MAXUSERQUESTS
        Call CleanQuestSlot(UserIndex, i)
    Next i
    With UserList(UserIndex).QuestStats
        .NumQuestsDone = 0
        Erase .QuestsDone
    End With
    Exit Sub
ResetQuestStats_Err:
    Call TraceError(Err.Number, Err.Description, "ModQuest.ResetQuestStats", Erl)
End Sub
 
Public Sub LoadQuests()
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Carga el archivo QUESTS.DAT en el array QuestList.
    'Last modified: 27/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    On Error GoTo ErrorHandler
    Dim reader    As clsIniManager
    Dim NumQuests As Integer
    Dim tmpStr    As String
    Dim i         As Integer
    Dim j         As Integer
    'Cargamos el clsIniManager en memoria
    Set reader = New clsIniManager
    'Lo inicializamos para el archivo Quests.DAT
    Call reader.Initialize(DatPath & "Quests.DAT")
    'Redimensionamos el array
    NumQuests = reader.GetValue("INIT", "NumQuests")
    ReDim QuestList(1 To NumQuests)
    'Cargamos los datos
    For i = 1 To NumQuests
        With QuestList(i)
            .nombre = reader.GetValue("QUEST" & i, "Nombre")
            .Desc = reader.GetValue("QUEST" & i, "Desc")
            .RequiredLevel = val(reader.GetValue("QUEST" & i, "RequiredLevel"))
            .RequiredClass = val(reader.GetValue("QUEST" & i, "RequiredClass"))
            .RequiredQuest = val(reader.GetValue("QUEST" & i, "RequiredQuest"))
            .LimitLevel = val(reader.GetValue("QUEST" & i, "LimitLevel"))
            .DescFinal = reader.GetValue("QUEST" & i, "DescFinal")
            .NextQuest = reader.GetValue("QUEST" & i, "NextQuest")
            'CARGAMOS OBJETOS REQUERIDOS
            .RequiredOBJs = val(reader.GetValue("QUEST" & i, "RequiredOBJs"))
            .Trabajador = IIf(val(reader.GetValue("QUEST" & i, "Trabajador")) = 1, True, False)
            .TalkTo = val(reader.GetValue("QUEST" & i, "TalkTo"))
            If .RequiredOBJs > 0 Then
                ReDim .RequiredOBJ(1 To .RequiredOBJs)
                For j = 1 To .RequiredOBJs
                    tmpStr = reader.GetValue("QUEST" & i, "RequiredOBJ" & j)
                    .RequiredOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredOBJ(j).amount = val(ReadField(2, tmpStr, 45))
                Next j
            End If
            .RequiredSpellCount = val(reader.GetValue("QUEST" & i, "RequiredSpellCount"))
            If .RequiredSpellCount > 0 Then
                ReDim .RequiredSpellList(1 To .RequiredSpellCount) As Integer
                For j = 1 To .RequiredSpellCount
                    .RequiredSpellList(j) = val(reader.GetValue("QUEST" & i, "RequiredSpell" & j))
                Next j
            End If
            'CARGAMOS NPCS REQUERIDOS
            .RequiredNPCs = val(reader.GetValue("QUEST" & i, "RequiredNPCs"))
            If .RequiredNPCs > 0 Then
                ReDim .RequiredNPC(1 To .RequiredNPCs)
                For j = 1 To .RequiredNPCs
                    tmpStr = reader.GetValue("QUEST" & i, "RequiredNPC" & j)
                    .RequiredNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredNPC(j).amount = val(ReadField(2, tmpStr, 45))
                Next j
            End If
            .RequiredSkill.SkillType = val(reader.GetValue("QUEST" & i, "RequiredSkill"))
            .RequiredSkill.RequiredValue = val(reader.GetValue("QUEST" & i, "RequiredValue"))
            'CARGAMOS NPCS TARGET REQUERIDOS
            .RequiredTargetNPCs = val(reader.GetValue("QUEST" & i, "RequiredTargetNPCs"))
            If .RequiredTargetNPCs > 0 Then
                ReDim .RequiredTargetNPC(1 To .RequiredTargetNPCs)
                For j = 1 To .RequiredTargetNPCs
                    tmpStr = reader.GetValue("QUEST" & i, "RequiredTargetNPC" & j)
                    .RequiredTargetNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredTargetNPC(j).amount = 1
                Next j
            End If
            .RewardGLD = val(reader.GetValue("QUEST" & i, "RewardGLD"))
            .RewardEXP = val(reader.GetValue("QUEST" & i, "RewardEXP"))
            .Repetible = val(reader.GetValue("QUEST" & i, "Repetible"))
            'CARGAMOS OBJETOS DE RECOMPENSA
            .RewardOBJs = val(reader.GetValue("QUEST" & i, "RewardOBJs"))
            If .RewardOBJs > 0 Then
                ReDim .RewardOBJ(1 To .RewardOBJs)
                For j = 1 To .RewardOBJs
                    tmpStr = reader.GetValue("QUEST" & i, "RewardOBJ" & j)
                    .RewardOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RewardOBJ(j).amount = val(ReadField(2, tmpStr, 45))
                Next j
            End If
            .RewardSpellCount = val(reader.GetValue("QUEST" & i, "RewardSkills"))
            If .RewardSpellCount > 0 Then
                ReDim .RewardSpellList(1 To .RewardSpellCount)
                For j = 1 To .RewardSpellCount
                    .RewardSpellList(j) = val(reader.GetValue("QUEST" & i, "RewardSkill" & j))
                Next j
            End If
        End With
    Next i
    'Eliminamos la clase
    Set reader = Nothing
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
    With UserList(UserIndex).QuestStats
        For i = 1 To MAXUSERQUESTS - 1
            If .Quests(i).QuestIndex = 0 Then
                For j = i + 1 To MAXUSERQUESTS
                    If .Quests(j).QuestIndex Then
                        .Quests(i) = .Quests(j)
                        Call CleanQuestSlot(UserIndex, j)
                        Exit For
                    End If
                Next j
            End If
        Next i
    End With
    Exit Sub
ArrangeUserQuests_Err:
    Call TraceError(Err.Number, Err.Description, "ModQuest.ArrangeUserQuests", Erl)
End Sub
 
Public Sub EnviarQuest(ByVal UserIndex As Integer)
    On Error GoTo EnviarQuest_Err
    Dim NpcIndex As Integer
    Dim tmpByte  As Byte
    If Not IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then Exit Sub
    NpcIndex = UserList(UserIndex).flags.TargetNPC.ArrayIndex
    'Esta el personaje en la distancia correcta?
    If Distancia(UserList(UserIndex).pos, NpcList(NpcIndex).pos) > 5 Then
        ' Msg8=Estas demasiado lejos.
        Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    'El NPC hace quests?
    If NpcList(NpcIndex).NumQuest = 0 Then
        Call WriteLocaleChatOverHead(UserIndex, "1341", "", NpcList(NpcIndex).Char.charindex, vbYellow) ' Msg1341=No tengo ninguna misión para ti.
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
                        Call FinishQuest(UserIndex, i, tmpByte)
                        Exit Sub
                    End If
                Next j
            End If
        End If
    Next i
    For q = 1 To NpcList(NpcIndex).NumQuest
        tmpByte = TieneQuest(UserIndex, NpcList(NpcIndex).QuestNumber(q))
        If tmpByte Then
            'El usuario esta haciendo la quest, entonces va a hablar con el NPC para recibir la recompensa.
            If FinishQuestCheck(UserIndex, NpcList(NpcIndex).QuestNumber(q), tmpByte) Then
                Call FinishQuest(UserIndex, NpcList(NpcIndex).QuestNumber(q), tmpByte)
                Exit Sub
            End If
        End If
    Next q
    Call WriteNpcQuestListSend(UserIndex, NpcIndex)
    Exit Sub
EnviarQuest_Err:
    Call TraceError(Err.Number, Err.Description, "ModQuest.EnviarQuest", Erl)
End Sub

Public Function FinishQuestCheck(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte) As Boolean
    On Error GoTo FinishQuestCheck_Err
    Dim i        As Integer
    Dim NpcIndex As Integer
    NpcIndex = UserList(UserIndex).flags.TargetNPC.ArrayIndex
    With QuestList(QuestIndex)
        'Comprobamos que tenga los objetos.
        If .RequiredOBJs > 0 Then
            For i = 1 To .RequiredOBJs
                If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).amount, UserIndex) = False Then
                    FinishQuestCheck = False
                    Exit Function
                End If
            Next i
        End If
        'Comprobamos que haya matado todas las criaturas.
        If .RequiredNPCs > 0 Then
            For i = 1 To .RequiredNPCs
                If .RequiredNPC(i).amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
                    FinishQuestCheck = False
                    Exit Function
                End If
            Next i
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
        If .RequiredTargetNPCs > 0 Then
            For i = 1 To .RequiredTargetNPCs
                If .RequiredTargetNPC(i).amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsTarget(i) Then
                    FinishQuestCheck = False
                    Exit Function
                End If
            Next i
        End If
    End With
    FinishQuestCheck = True
    Exit Function
FinishQuestCheck_Err:
    Call TraceError(Err.Number, Err.Description, "ModQuest.FinishQuestCheck", Erl)
End Function

Function FaltanItemsQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    On Error GoTo Handler
    With QuestList(QuestIndex)
        ' Por las dudas...
        If .RequiredOBJs > 0 Then
            Dim i As Integer
            For i = 1 To .RequiredOBJs
                ' Encontramos el objeto
                If ObjIndex = .RequiredOBJ(i).ObjIndex Then
                    ' Devolvemos si ya tiene todos los que la quest pide
                    FaltanItemsQuest = Not TieneObjetos(ObjIndex, .RequiredOBJ(i).amount, UserIndex)
                    Exit Function
                End If
            Next i
        End If
    End With
    Exit Function
Handler:
    Call TraceError(Err.Number, Err.Description, "ModQuest.FaltanItemsQuest", Erl)
End Function

Public Function CanUserAcceptQuest(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal QuestIndex As Integer, ByRef tmpQuest As t_Quest) As Boolean
    On Error GoTo ErrHandler
    CanUserAcceptQuest = False
    If tmpQuest.Trabajador And UserList(UserIndex).clase <> e_Class.Trabajador Then
        Call WriteLocaleMsg(UserIndex, 1262, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If NpcIndex > 0 Then
        If Distancia(UserList(UserIndex).pos, NpcList(NpcIndex).pos) > 5 Then
            ' Msg8=Estas demasiado lejos.
            Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    End If
    If TieneQuest(UserIndex, QuestIndex) Then
        Call WriteLocaleMsg(UserIndex, 1263, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If tmpQuest.RequiredQuest > 0 Then
        If Not UserDoneQuest(UserIndex, tmpQuest.RequiredQuest) Then
            Call WriteLocaleMsg(UserIndex, 1424, e_FontTypeNames.FONTTYPE_INFO, QuestList(tmpQuest.RequiredQuest).nombre)
            Exit Function
        End If
    End If
    'El personaje tiene suficiente nivel?
    If UserList(UserIndex).Stats.ELV < tmpQuest.RequiredLevel Then
        Call WriteLocaleMsg(UserIndex, 1425, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    'El personaje es nivel muy alto?
    If tmpQuest.LimitLevel > 0 Then 'Si el nivel limite es mayor a 0, por si no esta asignada la propiedad en quest.dat
        If UserList(UserIndex).Stats.ELV > tmpQuest.LimitLevel Then
            Call WriteLocaleMsg(UserIndex, 1416, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    End If
    If tmpQuest.RequiredSkill.SkillType > 0 Then
        If UserList(UserIndex).Stats.UserSkills(tmpQuest.RequiredSkill.SkillType) < tmpQuest.RequiredSkill.RequiredValue Then
            Call WriteLocaleMsg(UserIndex, 473, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    End If
    If UserList(UserIndex).clase <> tmpQuest.RequiredClass And tmpQuest.RequiredClass > 0 Then
        Call WriteLocaleMsg(UserIndex, 1426, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    If tmpQuest.Repetible = 0 Then
        If UserDoneQuest(UserIndex, QuestIndex) Then
            Call WriteLocaleMsg(UserIndex, "QUESTNEXT*", e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
    End If
    CanUserAcceptQuest = True
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "ModQuest.CanUserAcceptQuest", Erl)
End Function
