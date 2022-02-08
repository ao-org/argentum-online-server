Attribute VB_Name = "InvNpc"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
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
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Inv & Obj
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Modulo para controlar los objetos y los inventarios.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
Public Function TirarItemAlPiso(Pos As t_WorldPos, obj As t_Obj, Optional PuedeAgua As Boolean = True) As t_WorldPos

        On Error GoTo ErrHandler

        Dim NuevaPos As t_WorldPos

100     NuevaPos.X = 0
102     NuevaPos.Y = 0
    
104     Tilelibre Pos, NuevaPos, obj, PuedeAgua, True

106     If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
108         Call MakeObj(obj, Pos.Map, NuevaPos.X, NuevaPos.Y)
        End If

110     TirarItemAlPiso = NuevaPos

        Exit Function
ErrHandler:

End Function

Public Sub NPC_TIRAR_ITEMS(ByRef npc As t_Npc)
        
        On Error GoTo NPC_TIRAR_ITEMS_Err
    
        

        'TIRA TODOS LOS ITEMS DEL NPC
        

100     If npc.Invent.NroItems > 0 Then
    
            Dim i     As Byte

            Dim MiObj As t_Obj
    
102         For i = 1 To MAX_INVENTORY_SLOTS
    
104             If npc.Invent.Object(i).ObjIndex > 0 Then
106                 MiObj.amount = npc.Invent.Object(i).amount
108                 MiObj.ObjIndex = npc.Invent.Object(i).ObjIndex
110                 Call TirarItemAlPiso(npc.Pos, MiObj, npc.flags.AguaValida = 1)

                End If
      
112         Next i

        End If

        
        Exit Sub

NPC_TIRAR_ITEMS_Err:
114     Call TraceError(Err.Number, Err.Description, "InvNpc.NPC_TIRAR_ITEMS", Erl)

        
End Sub

Function QuedanItems(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Boolean
        
        On Error GoTo QuedanItems_Err

        Dim i As Integer

100     If NpcList(NpcIndex).Invent.NroItems > 0 Then

102         For i = 1 To MAX_INVENTORY_SLOTS

104             If NpcList(NpcIndex).Invent.Object(i).ObjIndex = ObjIndex Then
106                 QuedanItems = True
                    Exit Function

                End If

            Next

        End If

108     QuedanItems = False

        
        Exit Function

QuedanItems_Err:
110     Call TraceError(Err.Number, Err.Description, "InvNpc.QuedanItems", Erl)

        
End Function

Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer
        
        On Error GoTo EncontrarCant_Err
        
        'Devuelve la cantidad original del obj de un npc

        Dim ln As String, npcfile As String

        Dim i  As Integer

        'If NpcList(NpcIndex).Numero > 499 Then
        '    npcfile = DatPath & "NPCs-HOSTILES.dat"
        'Else
100     npcfile = DatPath & "NPCs.dat"
        'End If
 
102     For i = 1 To MAX_INVENTORY_SLOTS
104         ln = GetVar(npcfile, "NPC" & NpcList(NpcIndex).Numero, "Obj" & i)

106         If ObjIndex = val(ReadField(1, ln, 45)) Then
108             EncontrarCant = val(ReadField(2, ln, 45))
                Exit Function

            End If

        Next
                   
110     EncontrarCant = 0
        
        Exit Function

EncontrarCant_Err:
112     Call TraceError(Err.Number, Err.Description, "InvNpc.EncontrarCant", Erl)

        
End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)
        
        On Error GoTo ResetNpcInv_Err
    
        Dim i As Integer

100     NpcList(NpcIndex).Invent.NroItems = 0

102     For i = 1 To MAX_INVENTORY_SLOTS
104         NpcList(NpcIndex).Invent.Object(i).ObjIndex = 0
106         NpcList(NpcIndex).Invent.Object(i).amount = 0
108     Next i

110     NpcList(NpcIndex).InvReSpawn = 0
        
        Exit Sub

ResetNpcInv_Err:
112     Call TraceError(Err.Number, Err.Description, "InvNpc.ResetNpcInv", Erl)

        
End Sub

Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarNpcInvItem_Err
        

        Dim ObjIndex As Integer

100     ObjIndex = NpcList(NpcIndex).Invent.Object(Slot).ObjIndex

        'Quita un Obj
102     If ObjData(NpcList(NpcIndex).Invent.Object(Slot).ObjIndex).Crucial = 0 Then
104         NpcList(NpcIndex).Invent.Object(Slot).amount = NpcList(NpcIndex).Invent.Object(Slot).amount - Cantidad
        
106         If NpcList(NpcIndex).Invent.Object(Slot).amount <= 0 Then
108             NpcList(NpcIndex).Invent.NroItems = NpcList(NpcIndex).Invent.NroItems - 1
110             NpcList(NpcIndex).Invent.Object(Slot).ObjIndex = 0
112             NpcList(NpcIndex).Invent.Object(Slot).amount = 0

114             If NpcList(NpcIndex).Invent.NroItems = 0 And NpcList(NpcIndex).InvReSpawn <> 1 Then
116                 Call CargarInvent(NpcIndex) 'Reponemos el inventario

                End If

            End If

        Else
118         NpcList(NpcIndex).Invent.Object(Slot).amount = NpcList(NpcIndex).Invent.Object(Slot).amount - Cantidad
        
120         If NpcList(NpcIndex).Invent.Object(Slot).amount <= 0 Then
122             NpcList(NpcIndex).Invent.NroItems = NpcList(NpcIndex).Invent.NroItems - 1
124             NpcList(NpcIndex).Invent.Object(Slot).ObjIndex = 0
126             NpcList(NpcIndex).Invent.Object(Slot).amount = 0
            
128             If Not QuedanItems(NpcIndex, ObjIndex) Then

                    Dim NoEsdeAca As Integer

130                 NoEsdeAca = EncontrarCant(NpcIndex, ObjIndex)

132                 If NoEsdeAca <> 0 Then
134                     NpcList(NpcIndex).Invent.Object(Slot).ObjIndex = ObjIndex
136                     NpcList(NpcIndex).Invent.Object(Slot).amount = EncontrarCant(NpcIndex, ObjIndex)
138                     NpcList(NpcIndex).Invent.NroItems = NpcList(NpcIndex).Invent.NroItems + 1

                    End If

                End If
            
140             If NpcList(NpcIndex).Invent.NroItems = 0 And NpcList(NpcIndex).InvReSpawn <> 1 Then
142                 Call CargarInvent(NpcIndex) 'Reponemos el inventario

                End If

            End If
    
        End If

        
        Exit Sub

QuitarNpcInvItem_Err:
144     Call TraceError(Err.Number, Err.Description, "InvNpc.QuitarNpcInvItem", Erl)

        
End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)
        
        On Error GoTo CargarInvent_Err
        

        'Vuelve a cargar el inventario del npc NpcIndex
        Dim LoopC   As Integer

        Dim ln      As String

        Dim npcfile As String

        'If NpcList(NpcIndex).Numero > 499 Then
        '    npcfile = DatPath & "NPCs-HOSTILES.dat"
        'Else
100     npcfile = DatPath & "NPCs.dat"
        'End If

102     NpcList(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcList(NpcIndex).Numero, "NROITEMS"))
        
        If NpcList(NpcIndex).Invent.NroItems > 0 Then
104         For LoopC = 1 To NpcList(NpcIndex).Invent.NroItems
106             ln = GetVar(npcfile, "NPC" & NpcList(NpcIndex).Numero, "Obj" & LoopC)
108             NpcList(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
110             NpcList(NpcIndex).Invent.Object(LoopC).amount = val(ReadField(2, ln, 45))
        
112         Next LoopC
        End If

        
        Exit Sub

CargarInvent_Err:
114     Call TraceError(Err.Number, Err.Description, "InvNpc.CargarInvent", Erl)

        
End Sub

Public Sub NpcDropeo(ByRef npc As t_Npc, ByRef UserIndex As Integer)

        On Error GoTo ErrHandler

100     If npc.NumQuiza = 0 Then Exit Sub
102     If DropActive = 0 Then Exit Sub 'Esta el Dropeo activado?

        Dim Dropeo       As t_Obj

        Dim Probabilidad As Long

        Dim objRandom    As Byte

        Dim ProbTiro     As Byte

        Dim nfile        As Integer

104     If npc.QuizaProb = 0 Then
106         Probabilidad = RandomNumber(1, DropMult) 'Tiro Item?
        Else
108         Probabilidad = RandomNumber(1, npc.QuizaProb) 'Tiro Item?

        End If

110     If UserList(UserIndex).Invent.MagicoObjIndex = 383 Then
112         If npc.QuizaProb = 0 Then
114             Probabilidad = RandomNumber(1, DropMult / 2) 'Tiro Item?
            Else
116             Probabilidad = RandomNumber(1, npc.QuizaProb / 2) 'Tiro Item?

            End If

        End If

118     If Probabilidad <> 1 Then Exit Sub

120     objRandom = RandomNumber(1, npc.NumQuiza) 'Que item puede ser que tire?

        Dim obj      As Integer

        Dim Cantidad As Integer

122     obj = val(ReadField(1, npc.QuizaDropea(objRandom), Asc("-")))
124     Cantidad = val(ReadField(2, npc.QuizaDropea(objRandom), Asc("-")))

126     Dropeo.amount = Cantidad 'Cantidad
128     Dropeo.ObjIndex = obj 'NUMERO DEL ITEM EN EL OBJ.DAT
130     Call TirarItemAlPiso(npc.Pos, Dropeo, npc.flags.AguaValida = 1)
132     Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave(e_FXSound.Dropeo_Sound, npc.Pos.X, npc.Pos.Y))
    
        Exit Sub

ErrHandler:
134     Call LogError("Error al dropear el item " & ObjData(npc.QuizaDropea(objRandom)).Name & ", al usuario " & UserList(UserIndex).Name & ". " & Err.Description & ".")

End Sub


Public Sub DropObjQuest(ByRef npc As t_Npc, ByRef UserIndex As Integer)
    'Dropeo por Quest
    'Ladder
    '3/12/2020
        On Error GoTo ErrHandler

100     If npc.NumDropQuest = 0 Then Exit Sub
    
        Dim Dropeo As t_Obj
        Dim Probabilidad As Long
        
        Dim i As Byte
    
102     For i = 1 To npc.NumDropQuest

104         With npc.DropQuest(i)

106             If .QuestIndex > 0 <> 0 Then
                    ' Tiene la quest?
108                 If TieneQuest(UserIndex, .QuestIndex) <> 0 Then
                        ' Si aún me faltan más de estos items de esta quest
110                     If FaltanItemsQuest(UserIndex, .QuestIndex, .ObjIndex) Then

112                         Probabilidad = RandomNumber(1, .Probabilidad) 'Tiro Item?
    
114                         If Probabilidad = 1 Then
116                             Dropeo.amount = .amount
118                             Dropeo.ObjIndex = .ObjIndex

                                'Call TirarItemAlPiso(npc.Pos, Dropeo)
                                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_FXSound.Dropeo_Sound, npc.Pos.X, npc.Pos.Y))

                                ' WyroX: Ahora te lo da en el inventario, si hay espacio, y el sonido lo escuchas vos solo
120                             Call MeterItemEnInventario(UserIndex, Dropeo)
122                             Call SendData(ToIndex, UserIndex, PrepareMessagePlayWave(e_FXSound.Dropeo_Sound, npc.Pos.X, npc.Pos.Y))
                            End If
                        End If
                    End If
                End If

            End With
124     Next i

        Exit Sub

ErrHandler:
126     Call LogError("Error DropObjQuest al dropear el item " & ObjData(npc.DropQuest(i).ObjIndex).Name & ", al usuario " & UserList(UserIndex).Name & ". " & Err.Description & ".")

End Sub

