Attribute VB_Name = "ModQuest"
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at [url=http://www.affero.org/oagpl.html]http://www.affero.org/oagpl.html[/url]
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at [email=aaron@baronsoft.com]aaron@baronsoft.com[/email]
'for more information about ORE please visit [url=http://www.baronsoft.com/]http://www.baronsoft.com/[/url]
Option Explicit
 
'Constantes de las quests
Public Const MAXUSERQUESTS As Integer = 5     'Maxima cantidad de quests que puede tener un usuario al mismo tiempo.

Public Function TieneQuest(ByVal UserIndex As Integer, ByVal QuestNumber As Integer) As Byte

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

End Function
 
Public Function FreeQuestSlot(ByVal UserIndex As Integer) As Byte

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

End Function
 
Public Sub FinishQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el evento de terminar una quest.
    'Last modified: 29/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i              As Integer

    Dim InvSlotsLibres As Byte

    Dim NpcIndex       As Integer
 
    NpcIndex = UserList(UserIndex).flags.TargetNPC
    
    With QuestList(QuestIndex)

        'Comprobamos que tenga los objetos.
        If .RequiredOBJs > 0 Then

            For i = 1 To .RequiredOBJs

                If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, UserIndex) = False Then
                    Call WriteChatOverHead(UserIndex, "No has conseguido todos los objetos que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
                    
                    Exit Sub

                End If

            Next i

        End If
        
        'Comprobamos que haya matado todas las criaturas.
        If .RequiredNPCs > 0 Then

            For i = 1 To .RequiredNPCs

                If .RequiredNPC(i).Amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
                    Call WriteChatOverHead(UserIndex, "No has matado todas las criaturas que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
                    Exit Sub

                End If

            Next i

        End If
    
        'Comprobamos que el usuario tenga espacio para recibir los items.
        If .RewardOBJs > 0 Then

            'Buscamos la cantidad de slots de inventario libres.
            For i = 1 To UserList(UserIndex).CurrentInventorySlots

                If UserList(UserIndex).Invent.Object(i).ObjIndex = 0 Then InvSlotsLibres = InvSlotsLibres + 1
            Next i
            
            'Nos fijamos si entra
            If InvSlotsLibres < .RewardOBJs Then
                Call WriteChatOverHead(UserIndex, "No tienes suficiente espacio en el inventario para recibir la recompensa. Vuelve cuando hayas hecho mas espacio.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
                Exit Sub

            End If

        End If
    
        'A esta altura ya cumplio los objetivos, entonces se le entregan las recompensas.
        'Call WriteConsoleMsg(UserIndex, "Has completado la mision " & Chr(34) & QuestList(QuestIndex).Nombre & Chr(34) & "!", FontTypeNames.FONTTYPE_New_Celeste)
        
        Call WriteChatOverHead(UserIndex, "QUESTFIN*" & Npclist(NpcIndex).QuestNumber, Npclist(NpcIndex).Char.CharIndex, vbYellow)
        
        'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(QuestList(Npclist(NpcIndex).QuestNumber).DescFinal, Npclist(NpcIndex).Char.CharIndex, vbYellow))

        'Si la quest pedia objetos, se los saca al personaje.
        If .RequiredOBJs Then

            For i = 1 To .RequiredOBJs
                Call QuitarObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, UserIndex)
            Next i

        End If
        
        'Se entrega la experiencia.
        If .RewardEXP Then
            If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + .RewardEXP
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)
                Call WriteLocaleMsg(UserIndex, "140", FontTypeNames.FONTTYPE_EXP, .RewardEXP)
            Else
                Call WriteConsoleMsg(UserIndex, "No se te ha dado experiencia porque eres nivel máximo.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'Se entrega el oro.
        If .RewardGLD Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + .RewardGLD
            Call WriteConsoleMsg(UserIndex, "Has ganado " & .RewardGLD & " monedas de oro como recompensa.", FontTypeNames.FONTTYPE_INFOIAO)

        End If
        
        'Si hay recompensa de objetos, se entregan.
        If .RewardOBJs > 0 Then

            For i = 1 To .RewardOBJs

                If .RewardOBJ(i).Amount Then
                    Call MeterItemEnInventario(UserIndex, .RewardOBJ(i))
                    Call WriteConsoleMsg(UserIndex, "Has recibido " & QuestList(QuestIndex).RewardOBJ(i).Amount & " " & ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).name & " como recompensa.", FontTypeNames.FONTTYPE_INFOIAO)

                End If

            Next i

        End If
        
        Call WriteUpdateGold(UserIndex)
    
        'Actualizamos el personaje
        Call UpdateUserInv(True, UserIndex, 0)
    
        'Limpiamos el slot de quest.
        Call CleanQuestSlot(UserIndex, QuestSlot)
        
        'Ordenamos las quests
        Call ArrangeUserQuests(UserIndex)
        
        If .Repetible = 0 Then
            'Se agrega que el usuario ya hizo esta quest.
            Call AddDoneQuest(UserIndex, QuestIndex)
            Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 2)

        End If
        
    End With

End Sub
 
Public Sub AddDoneQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Agrega la quest QuestIndex a la lista de quests hechas.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    With UserList(UserIndex).QuestStats
        .NumQuestsDone = .NumQuestsDone + 1
        ReDim Preserve .QuestsDone(1 To .NumQuestsDone)
        .QuestsDone(.NumQuestsDone) = QuestIndex

    End With

End Sub
 
Public Function UserDoneQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer) As Boolean

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Verifica si el usuario hizo la quest QuestIndex.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i As Integer

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
        
End Function
 
Public Sub CleanQuestSlot(ByVal UserIndex As Integer, ByVal QuestSlot As Integer)

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

        End If

        .QuestIndex = 0

    End With

End Sub
 
Public Sub ResetQuestStats(ByVal UserIndex As Integer)

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

End Sub
 
Public Sub LoadQuests()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Carga el archivo QUESTS.DAT en el array QuestList.
    'Last modified: 27/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    On Error GoTo ErrorHandler

    Dim Reader    As clsIniReader

    Dim NumQuests As Integer

    Dim tmpStr    As String

    Dim i         As Integer

    Dim j         As Integer
    
    'Cargamos el clsIniManager en memoria
    Set Reader = New clsIniReader
    
    'Lo inicializamos para el archivo Quests.DAT
    Call Reader.Initialize(DatPath & "Quests.DAT")
    
    'Redimensionamos el array
    NumQuests = Reader.GetValue("INIT", "NumQuests")
    ReDim QuestList(1 To NumQuests)
    
    'Cargamos los datos
    For i = 1 To NumQuests

        With QuestList(i)
            .nombre = Reader.GetValue("QUEST" & i, "Nombre")
            .Desc = Reader.GetValue("QUEST" & i, "Desc")
            .RequiredLevel = val(Reader.GetValue("QUEST" & i, "RequiredLevel"))
            
            .DescFinal = Reader.GetValue("QUEST" & i, "DescFinal")
            
            .NextQuest = Reader.GetValue("QUEST" & i, "NextQuest")
            
            'CARGAMOS OBJETOS REQUERIDOS
            .RequiredOBJs = val(Reader.GetValue("QUEST" & i, "RequiredOBJs"))

            If .RequiredOBJs > 0 Then
                ReDim .RequiredOBJ(1 To .RequiredOBJs)

                For j = 1 To .RequiredOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredOBJ" & j)
                    
                    .RequiredOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredOBJ(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If
            
            'CARGAMOS NPCS REQUERIDOS
            .RequiredNPCs = val(Reader.GetValue("QUEST" & i, "RequiredNPCs"))

            If .RequiredNPCs > 0 Then
                ReDim .RequiredNPC(1 To .RequiredNPCs)

                For j = 1 To .RequiredNPCs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredNPC" & j)
                    
                    .RequiredNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredNPC(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If
            
            .RewardGLD = val(Reader.GetValue("QUEST" & i, "RewardGLD"))
            .RewardEXP = val(Reader.GetValue("QUEST" & i, "RewardEXP"))
            .Repetible = val(Reader.GetValue("QUEST" & i, "Repetible"))
            
            'CARGAMOS OBJETOS DE RECOMPENSA
            .RewardOBJs = val(Reader.GetValue("QUEST" & i, "RewardOBJs"))

            If .RewardOBJs > 0 Then
                ReDim .RewardOBJ(1 To .RewardOBJs)

                For j = 1 To .RewardOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RewardOBJ" & j)
                    
                    .RewardOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RewardOBJ(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If

        End With

    Next i
    
    'Eliminamos la clase
    Set Reader = Nothing
    Exit Sub
                    
ErrorHandler:
    MsgBox "Error cargando el archivo QUESTS.DAT.", vbOKOnly + vbCritical

End Sub
 
Public Sub LoadQuestStats(ByVal UserIndex As Integer, ByRef UserFile As clsIniReader)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Carga las QuestStats del usuario.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i        As Integer

    Dim j        As Integer

    Dim tmpStr   As String

    Dim Fields() As String
 
    For i = 1 To MAXUSERQUESTS

        With UserList(UserIndex).QuestStats.Quests(i)
            tmpStr = UserFile.GetValue("QUESTS", "Q" & i)
            
            ' Para evitar modificar TODOS los charfiles
            If tmpStr = vbNullString Then
                .QuestIndex = 0

            Else
                Fields = Split(tmpStr, "-")

                .QuestIndex = val(Fields(0))

                If .QuestIndex Then
                    If QuestList(.QuestIndex).RequiredNPCs Then
                        ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)

                        For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                            .NPCsKilled(j) = val(Fields(j))
                        Next j

                    End If

                End If

            End If

        End With

    Next i
    
    With UserList(UserIndex).QuestStats
        tmpStr = UserFile.GetValue("QUESTS", "QuestsDone")
        
        If tmpStr = vbNullString Then
            .NumQuestsDone = 0
        
        Else
            Fields = Split(tmpStr, "-")

            .NumQuestsDone = val(Fields(0))

            If .NumQuestsDone Then
                ReDim .QuestsDone(1 To .NumQuestsDone)

                For i = 1 To .NumQuestsDone
                    .QuestsDone(i) = val(Fields(i))
                Next i

            End If

        End If

    End With
                   
End Sub
 
Public Sub SaveQuestStats(ByVal UserIndex As Integer, ByRef UserFile As String)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Guarda las QuestStats del usuario.
    'Last modified: 29/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim i      As Integer

    Dim j      As Integer

    Dim tmpStr As String
 
    For i = 1 To MAXUSERQUESTS

        With UserList(UserIndex).QuestStats.Quests(i)
            tmpStr = .QuestIndex
            
            If .QuestIndex Then
                If QuestList(.QuestIndex).RequiredNPCs Then

                    For j = 1 To QuestList(.QuestIndex).RequiredNPCs
                        tmpStr = tmpStr & "-" & .NPCsKilled(j)
                    Next j

                End If

            End If
        
            Call WriteVar(UserFile, "QUESTS", "Q" & i, tmpStr)

        End With

    Next i
    
    With UserList(UserIndex).QuestStats
        tmpStr = .NumQuestsDone
        
        If .NumQuestsDone Then

            For i = 1 To .NumQuestsDone
                tmpStr = tmpStr & "-" & .QuestsDone(i)
            Next i

        End If
        
        Call WriteVar(UserFile, "QUESTS", "QuestsDone", tmpStr)

    End With

End Sub
  
Public Sub ArrangeUserQuests(ByVal UserIndex As Integer)

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

End Sub
 
Public Sub EnviarQuest(ByVal UserIndex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el paquete Quest.
    'Last modified: 28/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim NpcIndex As Integer

    Dim tmpByte  As Byte
 
    NpcIndex = UserList(UserIndex).flags.TargetNPC
    
    If NpcIndex = 0 Then Exit Sub
    
    'Esta el personaje en la distancia correcta?
    If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > 5 Then
        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    'El NPC hace quests?
    If Npclist(NpcIndex).QuestNumber = 0 Then
        Call WriteChatOverHead(UserIndex, "No tengo ninguna mision para ti.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
        Exit Sub

    End If
    
    'El personaje ya hizo la quest?
    If UserDoneQuest(UserIndex, Npclist(NpcIndex).QuestNumber) Then
        ' Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(QuestList(Npclist(NpcIndex).QuestNumber).NextQuest, Npclist(NpcIndex).Char.CharIndex, vbYellow))
        
        Call WriteChatOverHead(UserIndex, "QUESTNEXT*" & Npclist(NpcIndex).QuestNumber, Npclist(NpcIndex).Char.CharIndex, vbYellow)
        Exit Sub

    End If
 
    'El personaje tiene suficiente nivel?
    If UserList(UserIndex).Stats.ELV < QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel Then
        Call WriteChatOverHead(UserIndex, "Debes ser por lo menos nivel " & QuestList(Npclist(NpcIndex).QuestNumber).RequiredLevel & " para emprender esta mision.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
        
        Exit Sub

    End If
    
    'A esta altura ya analizo todas las restricciones y esta preparado para el handle propiamente dicho
 
    tmpByte = TieneQuest(UserIndex, Npclist(NpcIndex).QuestNumber)
    
    If tmpByte Then
        'El usuario esta haciendo la quest, entonces va a hablar con el NPC para recibir la recompensa.
        Call FinishQuest(UserIndex, Npclist(NpcIndex).QuestNumber, tmpByte)
    Else
        'El usuario no esta haciendo la quest, entonces primero recibe un informe con los detalles de la mision.
        tmpByte = FreeQuestSlot(UserIndex)
        
        'El personaje tiene algun slot de quest para la nueva quest?
        If tmpByte = 0 Then
            Call WriteChatOverHead(UserIndex, "Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Exit Sub

        End If
        
        'Enviamos los detalles de la quest
        Call WriteQuestDetails(UserIndex, Npclist(NpcIndex).QuestNumber)

    End If

End Sub
