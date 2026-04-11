Attribute VB_Name = "InvNpc"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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

Public Const MaxDropCount As Byte = 12

Public Function TirarItemAlPiso(pos As t_WorldPos, obj As t_Obj, Optional PuedeAgua As Boolean = True) As t_WorldPos
    On Error GoTo ErrHandler
    Dim NuevaPos As t_WorldPos
    NuevaPos.x = 0
    NuevaPos.y = 0
    Tilelibre pos, NuevaPos, obj, PuedeAgua, True
    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then
        Call MakeObj(obj, pos.Map, NuevaPos.x, NuevaPos.y)
    End If
    TirarItemAlPiso = NuevaPos
    Exit Function
ErrHandler:
End Function

Function QuedanItems(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    On Error GoTo QuedanItems_Err
    Dim i As Integer
    If NpcList(NpcIndex).invent.NroItems > 0 Then
        For i = 1 To MAX_INVENTORY_SLOTS
            If NpcList(NpcIndex).invent.Object(i).ObjIndex = ObjIndex Then
                QuedanItems = True
                Exit Function
            End If
        Next
    End If
    QuedanItems = False
    Exit Function
QuedanItems_Err:
    Call TraceError(Err.Number, Err.Description, "InvNpc.QuedanItems", Erl)
End Function

Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer
    On Error GoTo EncontrarCant_Err
    'Devuelve la cantidad original del obj de un npc
    Dim ln As String, npcfile As String
    Dim i  As Integer
    npcfile = DatPath & "NPCs.dat"
    'End If
    For i = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & NpcList(NpcIndex).Numero, "Obj" & i)
        If ObjIndex = val(ReadField(1, ln, 45)) Then
            EncontrarCant = val(ReadField(2, ln, 45))
            Exit Function
        End If
    Next
    EncontrarCant = 0
    Exit Function
EncontrarCant_Err:
    Call TraceError(Err.Number, Err.Description, "InvNpc.EncontrarCant", Erl)
End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)
    On Error GoTo ResetNpcInv_Err
    Dim i As Integer
    NpcList(NpcIndex).invent.NroItems = 0
    For i = 1 To MAX_INVENTORY_SLOTS
        NpcList(NpcIndex).invent.Object(i).ObjIndex = 0
        NpcList(NpcIndex).invent.Object(i).amount = 0
    Next i
    NpcList(NpcIndex).InvReSpawn = 0
    Exit Sub
ResetNpcInv_Err:
    Call TraceError(Err.Number, Err.Description, "InvNpc.ResetNpcInv", Erl)
End Sub

Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
    On Error GoTo QuitarNpcInvItem_Err
    Dim ObjIndex As Integer
    ObjIndex = NpcList(NpcIndex).invent.Object(Slot).ObjIndex
    'Quita un Obj
    If ObjData(NpcList(NpcIndex).invent.Object(Slot).ObjIndex).Crucial = 0 Then
        NpcList(NpcIndex).invent.Object(Slot).amount = NpcList(NpcIndex).invent.Object(Slot).amount - Cantidad
        If NpcList(NpcIndex).invent.Object(Slot).amount <= 0 Then
            NpcList(NpcIndex).invent.NroItems = NpcList(NpcIndex).invent.NroItems - 1
            NpcList(NpcIndex).invent.Object(Slot).ObjIndex = 0
            NpcList(NpcIndex).invent.Object(Slot).amount = 0
            If NpcList(NpcIndex).invent.NroItems = 0 And NpcList(NpcIndex).InvReSpawn <> 1 Then
                Call CargarInvent(NpcIndex) 'Reponemos el inventario
            End If
        End If
    Else
        NpcList(NpcIndex).invent.Object(Slot).amount = NpcList(NpcIndex).invent.Object(Slot).amount - Cantidad
        If NpcList(NpcIndex).invent.Object(Slot).amount <= 0 Then
            NpcList(NpcIndex).invent.NroItems = NpcList(NpcIndex).invent.NroItems - 1
            NpcList(NpcIndex).invent.Object(Slot).ObjIndex = 0
            NpcList(NpcIndex).invent.Object(Slot).amount = 0
            If Not QuedanItems(NpcIndex, ObjIndex) Then
                Dim NoEsdeAca As Integer
                NoEsdeAca = EncontrarCant(NpcIndex, ObjIndex)
                If NoEsdeAca <> 0 Then
                    NpcList(NpcIndex).invent.Object(Slot).ObjIndex = ObjIndex
                    NpcList(NpcIndex).invent.Object(Slot).amount = EncontrarCant(NpcIndex, ObjIndex)
                    NpcList(NpcIndex).invent.NroItems = NpcList(NpcIndex).invent.NroItems + 1
                End If
            End If
            If NpcList(NpcIndex).invent.NroItems = 0 And NpcList(NpcIndex).InvReSpawn <> 1 Then
                Call CargarInvent(NpcIndex) 'Reponemos el inventario
            End If
        End If
    End If
    Exit Sub
QuitarNpcInvItem_Err:
    Call TraceError(Err.Number, Err.Description, "InvNpc.QuitarNpcInvItem", Erl)
End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)
    On Error GoTo CargarInvent_Err
    'Vuelve a cargar el inventario del npc NpcIndex
    Dim LoopC   As Integer
    Dim ln      As String
    Dim npcfile As String
    npcfile = DatPath & "NPCs.dat"
    'End If
    NpcList(NpcIndex).invent.NroItems = val(GetVar(npcfile, "NPC" & NpcList(NpcIndex).Numero, "NROITEMS"))
    If NpcList(NpcIndex).invent.NroItems > 0 Then
        For LoopC = 1 To NpcList(NpcIndex).invent.NroItems
            ln = GetVar(npcfile, "NPC" & NpcList(NpcIndex).Numero, "Obj" & LoopC)
            NpcList(NpcIndex).invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
            NpcList(NpcIndex).invent.Object(LoopC).amount = val(ReadField(2, ln, 45))
        Next LoopC
    End If
    Exit Sub
CargarInvent_Err:
    Call TraceError(Err.Number, Err.Description, "InvNpc.CargarInvent", Erl)
End Sub

Public Sub NpcDropObj(ByRef Npc As t_Npc, ByRef UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim Dropeo       As t_Obj
    Dim Probabilidad As Long
    Dim DropMultiplier As Integer
    If Npc.DropCount = 0 Then Exit Sub
    DropMultiplier = SvrConfig.GetValue("DropMult")
    'out of bounds and data type check
    If DropMultiplier <= 0 Then
        DropMultiplier = 1
    End If
    Dim i As Byte
    For i = 1 To Npc.DropCount
        'long cast on division truncates decimals 5/2 = 2
        Probabilidad = Npc.Drop(i).DropChance / DropMultiplier
        If RandomNumber(1, Probabilidad) = 1 Then
            Dropeo.Amount = RandomNumber(Npc.Drop(i).LowQuantityBound, Npc.Drop(i).HighQuantityBound)
            Dropeo.ObjIndex = Npc.Drop(i).ItemIndex
            Call SendData(ToIndex, UserIndex, PrepareMessagePlayWave(e_SoundEffects.Dropeo_Sound, Npc.pos.x, Npc.pos.y))
            Call TirarItemAlPiso(Npc.pos, Dropeo, Npc.flags.AguaValida = 1)
        End If
    Next i
Exit Sub
ErrHandler:
    Call LogError("Error al dropear el item del npc numero:" & Npc.Numero & ", al usuario " & UserList(UserIndex).Name & ". " & Err.Description & ".")
End Sub

Public Sub NpcDropQuestObj(ByRef Npc As t_Npc, ByRef UserIndex As Integer)
    On Error GoTo ErrHandler
    If Npc.NumDropQuest = 0 Then Exit Sub
    Dim Dropeo       As t_Obj
    Dim Probabilidad As Long
    Dim i            As Byte
    For i = 1 To Npc.NumDropQuest
        With Npc.DropQuest(i)
        
            If .QuestIndex = 0 Then
                'next i  or continue not working
                GoTo NextQuest
            End If
            
            If TieneQuest(UserIndex, .QuestIndex) = 0 Then
                'next i  or continue not working
                GoTo NextQuest
            End If
            
            If Not FaltanItemsQuest(UserIndex, .QuestIndex, .ObjIndex) Then
                'next i  or continue not working
                GoTo NextQuest
            End If
            
            Probabilidad = RandomNumber(1, .Probabilidad) 'Tiro Item?
            If Probabilidad = 1 Then
                Dropeo.Amount = .Amount
                Dropeo.ObjIndex = .ObjIndex
                Call MeterItemEnInventario(UserIndex, Dropeo)
                Call SendData(ToIndex, UserIndex, PrepareMessagePlayWave(e_SoundEffects.Dropeo_Sound, Npc.pos.x, Npc.pos.y))
            End If
            
        End With
NextQuest:
    Next i
    Exit Sub
ErrHandler:
    Call LogError("Error DropObjQuest al dropear el item " & ObjData(Npc.DropQuest(i).ObjIndex).name & ", al usuario " & UserList(UserIndex).name & ". " & Err.Description & ".")
End Sub
