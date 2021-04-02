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
110     Call RegistrarError(Err.Number, Err.Description, "ModQuest.TieneQuest", Erl)
112     Resume Next
        
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
110     Call RegistrarError(Err.Number, Err.Description, "ModQuest.FreeQuestSlot", Erl)
112     Resume Next
        
End Function
 
Public Sub FinishQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte)
        
        On Error GoTo FinishQuest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el evento de terminar una quest.
        'Last modified: 29/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i              As Integer

        Dim InvSlotsLibres As Byte

        Dim NpcIndex       As Integer
        

 
100     NpcIndex = UserList(UserIndex).flags.TargetNPC
    
102     With QuestList(QuestIndex)

            'Comprobamos que tenga los objetos.
104         If .RequiredOBJs > 0 Then

106             For i = 1 To .RequiredOBJs

108                 If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, UserIndex) = False Then
110                     Call WriteChatOverHead(UserIndex, "No has conseguido todos los objetos que te he pedido.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
                    
                        Exit Sub

                    End If

112             Next i

            End If
        
            'Comprobamos que haya matado todas las criaturas.
114         If .RequiredNPCs > 0 Then

116             For i = 1 To .RequiredNPCs

118                 If .RequiredNPC(i).Amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
120                     Call WriteChatOverHead(UserIndex, "No has matado todas las criaturas que te he pedido.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
                        Exit Sub

                    End If

122             Next i

            End If
            
            'Comprobamos que haya targeteado todos los npc
124          If .RequiredTargetNPCs > 0 Then

126              For i = 1 To .RequiredTargetNPCs
    
128                  If .RequiredTargetNPC(i).Amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsTarget(i) Then
130                      Call WriteChatOverHead(UserIndex, "No has visitado al npc que te pedi.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
                        Exit Sub
    
                        End If
    
132              Next i

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
    
            'A esta altura ya cumplio los objetivos, entonces se le entregan las recompensas.
146         Call WriteChatOverHead(UserIndex, "QUESTFIN*" & QuestIndex, NpcList(NpcIndex).Char.CharIndex, vbYellow)
        

            'Si la quest pedia objetos, se los saca al personaje.
148         If .RequiredOBJs Then

150             For i = 1 To .RequiredOBJs
152                 Call QuitarObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, UserIndex)
154             Next i

            End If
        
            'Se entrega la experiencia.
156         If .RewardEXP Then
158             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
160                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + (.RewardEXP * ExpMult)
162                 Call WriteUpdateExp(UserIndex)
164                 Call CheckUserLevel(UserIndex)
166                 Call WriteLocaleMsg(UserIndex, "140", FontTypeNames.FONTTYPE_EXP, (.RewardEXP * ExpMult))
                Else
168                 Call WriteConsoleMsg(UserIndex, "No se te ha dado experiencia porque eres nivel máximo.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If
        
            'Se entrega el oro.
170         If .RewardGLD Then
172             UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + (.RewardGLD * OroMult)
174             Call WriteConsoleMsg(UserIndex, "Has ganado " & PonerPuntos((.RewardGLD * OroMult)) & " monedas de oro como recompensa.", FontTypeNames.FONTTYPE_INFOIAO)
176             Call WriteUpdateGold(UserIndex)
            End If
        
            'Si hay recompensa de objetos, se entregan.
178         If .RewardOBJs > 0 Then

180             For i = 1 To .RewardOBJs

182                 If .RewardOBJ(i).Amount Then
184                     Call MeterItemEnInventario(UserIndex, .RewardOBJ(i))
186                     Call WriteConsoleMsg(UserIndex, "Has recibido " & QuestList(QuestIndex).RewardOBJ(i).Amount & " " & ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).Name & " como recompensa.", FontTypeNames.FONTTYPE_INFOIAO)

                    End If

188             Next i

            End If
        
    
            'Actualizamos el personaje
190         Call UpdateUserInv(True, UserIndex, 0)
    
            'Limpiamos el slot de quest.
192         Call CleanQuestSlot(UserIndex, QuestSlot)
        
            'Ordenamos las quests
194         Call ArrangeUserQuests(UserIndex)
        
196         If .Repetible = 0 Then
                'Se agrega que el usuario ya hizo esta quest.
198             Call AddDoneQuest(UserIndex, QuestIndex)
200             Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 2)
            End If
        
        End With

        
        Exit Sub

FinishQuest_Err:
202     Call RegistrarError(Err.Number, Err.Description, "ModQuest.FinishQuest", Erl)
204     Resume Next
        
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
108     Call RegistrarError(Err.Number, Err.Description, "ModQuest.AddDoneQuest", Erl)
110     Resume Next
        
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
118     Call RegistrarError(Err.Number, Err.Description, "ModQuest.UserDoneQuest", Erl)
120     Resume Next
        
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
104             If QuestList(.QuestIndex).RequiredNPCs Then

106                 For i = 1 To QuestList(.QuestIndex).RequiredNPCs
108                     .NPCsKilled(i) = 0
110                 Next i

                End If
                
112           If QuestList(.QuestIndex).RequiredTargetNPCs Then

114              For i = 1 To QuestList(.QuestIndex).RequiredTargetNPCs
116                  .NPCsTarget(i) = 0
118              Next i

                End If

            End If

120         .QuestIndex = 0

        End With

        
        Exit Sub

CleanQuestSlot_Err:
122     Call RegistrarError(Err.Number, Err.Description, "ModQuest.CleanQuestSlot", Erl)
124     Resume Next
        
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
112     Call RegistrarError(Err.Number, Err.Description, "ModQuest.ResetQuestStats", Erl)
114     Resume Next
        
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
            
118             .RequiredQuest = val(Reader.GetValue("QUEST" & i, "RequiredQuest"))
            
120             .DescFinal = Reader.GetValue("QUEST" & i, "DescFinal")
            
122             .NextQuest = Reader.GetValue("QUEST" & i, "NextQuest")
            
                'CARGAMOS OBJETOS REQUERIDOS
124             .RequiredOBJs = val(Reader.GetValue("QUEST" & i, "RequiredOBJs"))

126             If .RequiredOBJs > 0 Then
128                 ReDim .RequiredOBJ(1 To .RequiredOBJs)

130                 For j = 1 To .RequiredOBJs
132                     tmpStr = Reader.GetValue("QUEST" & i, "RequiredOBJ" & j)
                    
134                     .RequiredOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
136                     .RequiredOBJ(j).Amount = val(ReadField(2, tmpStr, 45))
138                 Next j

                End If
            
                'CARGAMOS NPCS REQUERIDOS
140             .RequiredNPCs = val(Reader.GetValue("QUEST" & i, "RequiredNPCs"))

142             If .RequiredNPCs > 0 Then
144                 ReDim .RequiredNPC(1 To .RequiredNPCs)

146                 For j = 1 To .RequiredNPCs
148                     tmpStr = Reader.GetValue("QUEST" & i, "RequiredNPC" & j)
                    
150                     .RequiredNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
152                     .RequiredNPC(j).Amount = val(ReadField(2, tmpStr, 45))
154                 Next j

                End If
            
            
            
                'CARGAMOS NPCS TARGET REQUERIDOS
156             .RequiredTargetNPCs = val(Reader.GetValue("QUEST" & i, "RequiredTargetNPCs"))

158             If .RequiredTargetNPCs > 0 Then
160                 ReDim .RequiredTargetNPC(1 To .RequiredTargetNPCs)

162                 For j = 1 To .RequiredTargetNPCs
164                     tmpStr = Reader.GetValue("QUEST" & i, "RequiredTargetNPC" & j)
                    
166                     .RequiredTargetNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
168                     .RequiredTargetNPC(j).Amount = 1
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
190                     .RewardOBJ(j).Amount = val(ReadField(2, tmpStr, 45))
192                 Next j

                End If

            End With

194     Next i
    
        'Eliminamos la clase
196     Set Reader = Nothing
        Exit Sub
                    
ErrorHandler:
198     MsgBox "Error cargando el archivo QUESTS.DAT.", vbOKOnly + vbCritical

End Sub
 
Public Sub LoadQuestStats(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)
        
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

102         With UserList(UserIndex).QuestStats.Quests(i)
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
    
128     With UserList(UserIndex).QuestStats
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
150     Call RegistrarError(Err.Number, Err.Description, "ModQuest.LoadQuestStats", Erl)
152     Resume Next
        
End Sub
 
Public Sub SaveQuestStats(ByVal UserIndex As Integer, ByRef UserFile As String)
        
        On Error GoTo SaveQuestStats_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Guarda las QuestStats del usuario.
        'Last modified: 29/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim i      As Integer

        Dim j      As Integer

        Dim tmpStr As String
 
100     For i = 1 To MAXUSERQUESTS

102         With UserList(UserIndex).QuestStats.Quests(i)
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
    
120     With UserList(UserIndex).QuestStats
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
134     Call RegistrarError(Err.Number, Err.Description, "ModQuest.SaveQuestStats", Erl)
136     Resume Next
        
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
118     Call RegistrarError(Err.Number, Err.Description, "ModQuest.ArrangeUserQuests", Erl)
120     Resume Next
        
End Sub
 
Public Sub EnviarQuest(ByVal UserIndex As Integer)
        
        On Error GoTo EnviarQuest_Err
        

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete Quest.
        'Last modified: 28/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim NpcIndex As Integer

        Dim tmpByte  As Byte
 
100     NpcIndex = UserList(UserIndex).flags.TargetNPC
    
102     If NpcIndex = 0 Then Exit Sub
    
        'Esta el personaje en la distancia correcta?
104     If Distancia(UserList(UserIndex).Pos, NpcList(NpcIndex).Pos) > 5 Then
106         Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        'El NPC hace quests?
108     If NpcList(NpcIndex).NumQuest = 0 Then
110         Call WriteChatOverHead(UserIndex, "No tengo ninguna mision para ti.", NpcList(NpcIndex).Char.CharIndex, vbYellow)
            Exit Sub

        End If
            
        
        'Hago un for para chequear si alguna de las misiones que da el NPC ya se completo.
        Dim q As Byte
        
        
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
       ' Else
            'El usuario no esta haciendo la quest, entonces primero recibe un informe con los detalles de la mision.
            'tmpByte = FreeQuestSlot(UserIndex)
        
            'El personaje tiene algun slot de quest para la nueva quest?
            'If tmpByte = 0 Then
              '  Call WriteChatOverHead(UserIndex, "Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", NpcList(NpcIndex).Char.CharIndex, vbWhite)
              '  Exit Sub

           ' End If
        
            'Enviamos los detalles de la quest
            'Call WriteQuestDetails(UserIndex, NpcList(NpcIndex).QuestNumber(1))

      '  End If
      
124   Call WriteNpcQuestListSend(UserIndex, NpcIndex)

        
        Exit Sub

EnviarQuest_Err:
126     Call RegistrarError(Err.Number, Err.Description, "ModQuest.EnviarQuest", Erl)
128     Resume Next
        
End Sub

Public Function FinishQuestCheck(ByVal UserIndex As Integer, ByVal QuestIndex As Integer, ByVal QuestSlot As Byte) As Boolean
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Funcion para chequear si finalizo una quest
        'Ladder
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

        On Error GoTo FinishQuestCheck_Err
        
        Dim i              As Integer

        Dim InvSlotsLibres As Byte

        Dim NpcIndex       As Integer
 
100     NpcIndex = UserList(UserIndex).flags.TargetNPC
    
102     With QuestList(QuestIndex)

            'Comprobamos que tenga los objetos.
104         If .RequiredOBJs > 0 Then

106             For i = 1 To .RequiredOBJs

108                 If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, UserIndex) = False Then
110                     FinishQuestCheck = False
                    
                        Exit Function

                    End If

112             Next i

            End If
        
            'Comprobamos que haya matado todas las criaturas.
114         If .RequiredNPCs > 0 Then

116             For i = 1 To .RequiredNPCs

118                 If .RequiredNPC(i).Amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
120                     FinishQuestCheck = False
                        Exit Function

                    End If

122             Next i

            End If
            
            'Comprobamos que haya targeteado todas las criaturas.
124      If .RequiredTargetNPCs > 0 Then

126          For i = 1 To .RequiredTargetNPCs

128              If .RequiredTargetNPC(i).Amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsTarget(i) Then
130                  FinishQuestCheck = False
                        Exit Function

                    End If

132          Next i

            End If
            
        End With
        
        
134     FinishQuestCheck = True

        

        Exit Function

FinishQuestCheck_Err:
136     Call RegistrarError(Err.Number, Err.Description, "ModQuest.FinishQuestCheck", Erl)

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
108                     FaltanItemsQuest = Not TieneObjetos(ObjIndex, .RequiredOBJ(i).Amount, UserIndex)
                        Exit Function

                    End If
            
110             Next i
        
            End If

        End With
            
Handler:
112     Call RegistrarError(Err.Number, Err.Description, "ModQuest.FaltanItemsQuest", Erl)
114     Resume Next

End Function
