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

Public Function TieneQuest(ByVal Userindex As Integer, ByVal QuestNumber As Integer) As Byte
        
        On Error GoTo TieneQuest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Devuelve el slot de UserQuests en que tiene la quest QuestNumber. En caso contrario devuelve 0.
        'Last modified: 27/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i As Integer
 
100     For i = 1 To MAXUSERQUESTS

102         If UserList(Userindex).QuestStats.Quests(i).QuestIndex = QuestNumber Then
104             TieneQuest = i
                Exit Function

            End If

106     Next i
    
108     TieneQuest = 0

        
        Exit Function

TieneQuest_Err:
        Call RegistrarError(Err.Number, Err.description, "ModQuest.TieneQuest", Erl)
        Resume Next
        
End Function
 
Public Function FreeQuestSlot(ByVal Userindex As Integer) As Byte
        
        On Error GoTo FreeQuestSlot_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Devuelve el proximo slot de quest libre.
        'Last modified: 27/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i As Integer
 
100     For i = 1 To MAXUSERQUESTS

102         If UserList(Userindex).QuestStats.Quests(i).QuestIndex = 0 Then
104             FreeQuestSlot = i
                Exit Function

            End If

106     Next i
    
108     FreeQuestSlot = 0

        
        Exit Function

FreeQuestSlot_Err:
        Call RegistrarError(Err.Number, Err.description, "ModQuest.FreeQuestSlot", Erl)
        Resume Next
        
End Function
 
Public Sub FinishQuest(ByVal Userindex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte)
        
        On Error GoTo FinishQuest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el evento de terminar una quest.
        'Last modified: 29/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i              As Integer

        Dim InvSlotsLibres As Byte

        Dim NpcIndex       As Integer
        

 
100     NpcIndex = UserList(Userindex).flags.TargetNPC
    
102     With QuestList(QuestIndex)

            'Comprobamos que tenga los objetos.
104         If .RequiredOBJs > 0 Then

106             For i = 1 To .RequiredOBJs

108                 If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, Userindex) = False Then
110                     Call WriteChatOverHead(Userindex, "No has conseguido todos los objetos que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
                    
                        Exit Sub

                    End If

112             Next i

            End If
        
            'Comprobamos que haya matado todas las criaturas.
114         If .RequiredNPCs > 0 Then

116             For i = 1 To .RequiredNPCs

118                 If .RequiredNPC(i).Amount > UserList(Userindex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
120                     Call WriteChatOverHead(Userindex, "No has matado todas las criaturas que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
                        Exit Sub

                    End If

122             Next i

            End If
            
            'Comprobamos que haya targeteado todos los npc
             If .RequiredTargetNPCs > 0 Then

                 For i = 1 To .RequiredTargetNPCs
    
                     If .RequiredTargetNPC(i).Amount > UserList(Userindex).QuestStats.Quests(QuestSlot).NPCsTarget(i) Then
                         Call WriteChatOverHead(Userindex, "No has visitado al npc que te pedi.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
                        Exit Sub
    
                        End If
    
                 Next i

            End If
    
            'Comprobamos que el usuario tenga espacio para recibir los items.
124         If .RewardOBJs > 0 Then

                'Buscamos la cantidad de slots de inventario libres.
126             For i = 1 To UserList(Userindex).CurrentInventorySlots

128                 If UserList(Userindex).Invent.Object(i).ObjIndex = 0 Then InvSlotsLibres = InvSlotsLibres + 1
130             Next i
            
                'Nos fijamos si entra
132             If InvSlotsLibres < .RewardOBJs Then
134                 Call WriteChatOverHead(Userindex, "No tienes suficiente espacio en el inventario para recibir la recompensa. Vuelve cuando hayas hecho mas espacio.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
                    Exit Sub

                End If

            End If
    
            'A esta altura ya cumplio los objetivos, entonces se le entregan las recompensas.
136         Call WriteChatOverHead(Userindex, "QUESTFIN*" & QuestIndex, Npclist(NpcIndex).Char.CharIndex, vbYellow)
        

            'Si la quest pedia objetos, se los saca al personaje.
138         If .RequiredOBJs Then

140             For i = 1 To .RequiredOBJs
142                 Call QuitarObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, Userindex)
144             Next i

            End If
        
            'Se entrega la experiencia.
146         If .RewardEXP Then
148             If UserList(Userindex).Stats.ELV < STAT_MAXELV Then
150                 UserList(Userindex).Stats.Exp = UserList(Userindex).Stats.Exp + .RewardEXP
152                 Call WriteUpdateExp(Userindex)
154                 Call CheckUserLevel(Userindex)
156                 Call WriteLocaleMsg(Userindex, "140", FontTypeNames.FONTTYPE_EXP, .RewardEXP)
                Else
158                 Call WriteConsoleMsg(Userindex, "No se te ha dado experiencia porque eres nivel máximo.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If
        
            'Se entrega el oro.
160         If .RewardGLD Then
162             UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD + .RewardGLD
164             Call WriteConsoleMsg(Userindex, "Has ganado " & .RewardGLD & " monedas de oro como recompensa.", FontTypeNames.FONTTYPE_INFOIAO)

            End If
        
            'Si hay recompensa de objetos, se entregan.
166         If .RewardOBJs > 0 Then

168             For i = 1 To .RewardOBJs

170                 If .RewardOBJ(i).Amount Then
172                     Call MeterItemEnInventario(Userindex, .RewardOBJ(i))
174                     Call WriteConsoleMsg(Userindex, "Has recibido " & QuestList(QuestIndex).RewardOBJ(i).Amount & " " & ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).name & " como recompensa.", FontTypeNames.FONTTYPE_INFOIAO)

                    End If

176             Next i

            End If
        
178         Call WriteUpdateGold(Userindex)
    
            'Actualizamos el personaje
180         Call UpdateUserInv(True, Userindex, 0)
    
            'Limpiamos el slot de quest.
182         Call CleanQuestSlot(Userindex, QuestSlot)
        
            'Ordenamos las quests
184         Call ArrangeUserQuests(Userindex)
        
186         If .Repetible = 0 Then
                'Se agrega que el usuario ya hizo esta quest.
188             Call AddDoneQuest(Userindex, QuestIndex)
190             Call WriteUpdateNPCSimbolo(Userindex, NpcIndex, 2)
            End If
        
        End With

        
        Exit Sub

FinishQuest_Err:
        Call RegistrarError(Err.Number, Err.description, "ModQuest.FinishQuest", Erl)
        Resume Next
        
End Sub
 
Public Sub AddDoneQuest(ByVal Userindex As Integer, ByVal QuestIndex As Integer)
        
        On Error GoTo AddDoneQuest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Agrega la quest QuestIndex a la lista de quests hechas.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
100     With UserList(Userindex).QuestStats
102         .NumQuestsDone = .NumQuestsDone + 1
104         ReDim Preserve .QuestsDone(1 To .NumQuestsDone)
106         .QuestsDone(.NumQuestsDone) = QuestIndex

        End With

        
        Exit Sub

AddDoneQuest_Err:
        Call RegistrarError(Err.Number, Err.description, "ModQuest.AddDoneQuest", Erl)
        Resume Next
        
End Sub
 
Public Function UserDoneQuest(ByVal Userindex As Integer, ByVal QuestIndex As Integer) As Boolean
        
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
            
            

100     With UserList(Userindex).QuestStats

102         If .NumQuestsDone Then

104             For i = 1 To .NumQuestsDone

106                 If .QuestsDone(i) = QuestIndex Then
108                     UserDoneQuest = True
                        Exit Function

                    End If

110             Next i

            End If

        End With
    
112     UserDoneQuest = False
        
        
        Exit Function

UserDoneQuest_Err:
        Call RegistrarError(Err.Number, Err.description, "ModQuest.UserDoneQuest", Erl)
        Resume Next
        
End Function
 
Public Sub CleanQuestSlot(ByVal Userindex As Integer, ByVal QuestSlot As Integer)
        
        On Error GoTo CleanQuestSlot_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Limpia un slot de quest de un usuario.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i As Integer
 
100     With UserList(Userindex).QuestStats.Quests(QuestSlot)

102         If .QuestIndex Then
104             If QuestList(.QuestIndex).RequiredNPCs Then

106                 For i = 1 To QuestList(.QuestIndex).RequiredNPCs
108                     .NPCsKilled(i) = 0
110                 Next i

                End If
                
              If QuestList(.QuestIndex).RequiredTargetNPCs Then

                 For i = 1 To QuestList(.QuestIndex).RequiredTargetNPCs
                     .NPCsTarget(i) = 0
                 Next i

                End If

            End If

112         .QuestIndex = 0

        End With

        
        Exit Sub

CleanQuestSlot_Err:
        Call RegistrarError(Err.Number, Err.description, "ModQuest.CleanQuestSlot", Erl)
        Resume Next
        
End Sub
 
Public Sub ResetQuestStats(ByVal Userindex As Integer)
        
        On Error GoTo ResetQuestStats_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Limpia todos los QuestStats de un usuario
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i As Integer
 
100     For i = 1 To MAXUSERQUESTS
102         Call CleanQuestSlot(Userindex, i)
104     Next i
    
106     With UserList(Userindex).QuestStats
108         .NumQuestsDone = 0
110         Erase .QuestsDone

        End With

        
        Exit Sub

ResetQuestStats_Err:
        Call RegistrarError(Err.Number, Err.description, "ModQuest.ResetQuestStats", Erl)
        Resume Next
        
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
            
            .RequiredQuest = val(Reader.GetValue("QUEST" & i, "RequiredQuest"))
            
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
            
            
            
            'CARGAMOS NPCS TARGET REQUERIDOS
            .RequiredTargetNPCs = val(Reader.GetValue("QUEST" & i, "RequiredTargetNPCs"))

            If .RequiredTargetNPCs > 0 Then
                ReDim .RequiredTargetNPC(1 To .RequiredTargetNPCs)

                For j = 1 To .RequiredTargetNPCs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredTargetNPC" & j)
                    
                    .RequiredTargetNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredTargetNPC(j).Amount = 1
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
 
Public Sub LoadQuestStats(ByVal Userindex As Integer, ByRef UserFile As clsIniReader)
        
        On Error GoTo LoadQuestStats_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Carga las QuestStats del usuario.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i        As Integer

        Dim j        As Integer

        Dim tmpStr   As String

        Dim Fields() As String
 
100     For i = 1 To MAXUSERQUESTS

102         With UserList(Userindex).QuestStats.Quests(i)
104             tmpStr = UserFile.GetValue("QUESTS", "Q" & i)
            
                ' Para evitar modificar TODOS los charfiles
106             If tmpStr = vbNullString Then
108                 .QuestIndex = 0

                Else
110                 Fields = Split(tmpStr, "-")

112                 .QuestIndex = val(Fields(0))

114                 If .QuestIndex Then
116                     If QuestList(.QuestIndex).RequiredNPCs Then
118                         ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)

120                         For j = 1 To QuestList(.QuestIndex).RequiredNPCs
122                             .NPCsKilled(j) = val(Fields(j))
124                         Next j

                        End If

                    End If

                End If

            End With

126     Next i
    
128     With UserList(Userindex).QuestStats
130         tmpStr = UserFile.GetValue("QUESTS", "QuestsDone")
        
132         If tmpStr = vbNullString Then
134             .NumQuestsDone = 0
        
            Else
136             Fields = Split(tmpStr, "-")

138             .NumQuestsDone = val(Fields(0))

140             If .NumQuestsDone Then
142                 ReDim .QuestsDone(1 To .NumQuestsDone)

144                 For i = 1 To .NumQuestsDone
146                     .QuestsDone(i) = val(Fields(i))
148                 Next i

                End If

            End If

        End With
                   
        
        Exit Sub

LoadQuestStats_Err:
        Call RegistrarError(Err.Number, Err.description, "ModQuest.LoadQuestStats", Erl)
        Resume Next
        
End Sub
 
Public Sub SaveQuestStats(ByVal Userindex As Integer, ByRef UserFile As String)
        
        On Error GoTo SaveQuestStats_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Guarda las QuestStats del usuario.
        'Last modified: 29/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i      As Integer

        Dim j      As Integer

        Dim tmpStr As String
 
100     For i = 1 To MAXUSERQUESTS

102         With UserList(Userindex).QuestStats.Quests(i)
104             tmpStr = .QuestIndex
            
106             If .QuestIndex Then
108                 If QuestList(.QuestIndex).RequiredNPCs Then

110                     For j = 1 To QuestList(.QuestIndex).RequiredNPCs
112                         tmpStr = tmpStr & "-" & .NPCsKilled(j)
114                     Next j

                    End If
                
                End If
        
116             Call WriteVar(UserFile, "QUESTS", "Q" & i, tmpStr)


           End With
           
           
           
        

118     Next i
    
120     With UserList(Userindex).QuestStats
122         tmpStr = .NumQuestsDone
        
124         If .NumQuestsDone Then

126             For i = 1 To .NumQuestsDone
128                 tmpStr = tmpStr & "-" & .QuestsDone(i)
130             Next i

            End If
        
132         Call WriteVar(UserFile, "QUESTS", "QuestsDone", tmpStr)

        End With

        
        Exit Sub

SaveQuestStats_Err:
        Call RegistrarError(Err.Number, Err.description, "ModQuest.SaveQuestStats", Erl)
        Resume Next
        
End Sub
  
Public Sub ArrangeUserQuests(ByVal Userindex As Integer)
        
        On Error GoTo ArrangeUserQuests_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Ordena las quests del usuario de manera que queden todas al principio del arreglo.
        'Last modified: 30/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i As Integer

        Dim j As Integer
 
100     With UserList(Userindex).QuestStats

102         For i = 1 To MAXUSERQUESTS - 1

104             If .Quests(i).QuestIndex = 0 Then

106                 For j = i + 1 To MAXUSERQUESTS

108                     If .Quests(j).QuestIndex Then
110                         .Quests(i) = .Quests(j)
112                         Call CleanQuestSlot(Userindex, j)
                            Exit For

                        End If

114                 Next j

                End If

116         Next i

        End With

        
        Exit Sub

ArrangeUserQuests_Err:
        Call RegistrarError(Err.Number, Err.description, "ModQuest.ArrangeUserQuests", Erl)
        Resume Next
        
End Sub
 
Public Sub EnviarQuest(ByVal Userindex As Integer)
        
        On Error GoTo EnviarQuest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete Quest.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim NpcIndex As Integer

        Dim tmpByte  As Byte
 
100     NpcIndex = UserList(Userindex).flags.TargetNPC
    
102     If NpcIndex = 0 Then Exit Sub
    
        'Esta el personaje en la distancia correcta?
104     If Distancia(UserList(Userindex).Pos, Npclist(NpcIndex).Pos) > 5 Then
106         Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        'El NPC hace quests?
108     If Npclist(NpcIndex).NumQuest = 0 Then
110         Call WriteChatOverHead(Userindex, "No tengo ninguna mision para ti.", Npclist(NpcIndex).Char.CharIndex, vbYellow)
            Exit Sub

        End If
            
        
        'Hago un for para chequear si alguna de las misiones que da el NPC ya se completo.
        Dim q As Byte
        
        
        For q = 1 To Npclist(NpcIndex).NumQuest
120         tmpByte = TieneQuest(Userindex, Npclist(NpcIndex).QuestNumber(q))
        
122         If tmpByte Then
                'El usuario esta haciendo la quest, entonces va a hablar con el NPC para recibir la recompensa.
                If FinishQuestCheck(Userindex, Npclist(NpcIndex).QuestNumber(q), tmpByte) Then
124                 Call FinishQuest(Userindex, Npclist(NpcIndex).QuestNumber(q), tmpByte)
                    Exit Sub
                End If

            End If
        
        Next q
       ' Else
            'El usuario no esta haciendo la quest, entonces primero recibe un informe con los detalles de la mision.
126         'tmpByte = FreeQuestSlot(UserIndex)
        
            'El personaje tiene algun slot de quest para la nueva quest?
128         'If tmpByte = 0 Then
130           '  Call WriteChatOverHead(UserIndex, "Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
              '  Exit Sub

           ' End If
        
            'Enviamos los detalles de la quest
132         'Call WriteQuestDetails(UserIndex, Npclist(NpcIndex).QuestNumber(1))

      '  End If
      
      Call WriteNpcQuestListSend(Userindex, NpcIndex)

        
        Exit Sub

EnviarQuest_Err:
        Call RegistrarError(Err.Number, Err.description, "ModQuest.EnviarQuest", Erl)
        Resume Next
        
End Sub



Public Function FinishQuestCheck(ByVal Userindex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte) As Boolean
        '<EhHeader>
        On Error GoTo FinishQuestCheck_Err
        '</EhHeader>
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Funcion para chequear si finalizo una quest
        'Ladder
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i              As Integer

        Dim InvSlotsLibres As Byte

        Dim NpcIndex       As Integer
 
100     NpcIndex = UserList(Userindex).flags.TargetNPC
    
102     With QuestList(QuestIndex)

            'Comprobamos que tenga los objetos.
104         If .RequiredOBJs > 0 Then

106             For i = 1 To .RequiredOBJs

108                 If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, Userindex) = False Then
110                     FinishQuestCheck = False
                    
                        Exit Function

                    End If

112             Next i

            End If
        
            'Comprobamos que haya matado todas las criaturas.
114         If .RequiredNPCs > 0 Then

116             For i = 1 To .RequiredNPCs

118                 If .RequiredNPC(i).Amount > UserList(Userindex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
120                     FinishQuestCheck = False
                        Exit Function

                    End If

122             Next i

            End If
            
            'Comprobamos que haya targeteado todas las criaturas.
124      If .RequiredTargetNPCs > 0 Then

126          For i = 1 To .RequiredTargetNPCs

128              If .RequiredTargetNPC(i).Amount > UserList(Userindex).QuestStats.Quests(QuestSlot).NPCsTarget(i) Then
130                  FinishQuestCheck = False
                        Exit Function

                    End If

132          Next i

            End If
            
        End With
        
        
134     FinishQuestCheck = True
        
        Exit Function

FinishQuestCheck_Err:
        Call RegistrarError(Err.description & vbCrLf & _
               "in Argentum20Server.ModQuest.FinishQuestCheck " & _
           "at line " & Erl, _
           vbExclamation + vbOKOnly, "Application Error")
    Resume Next
    '</EhFooter>
End Function
