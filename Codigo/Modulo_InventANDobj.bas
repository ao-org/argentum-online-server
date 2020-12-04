Attribute VB_Name = "InvNpc"
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
Public Function TirarItemAlPiso(Pos As WorldPos, obj As obj, Optional NotPirata As Boolean = True) As WorldPos

    On Error GoTo ErrHandler

    Dim NuevaPos As WorldPos

    NuevaPos.X = 0
    NuevaPos.Y = 0
    
    Tilelibre Pos, NuevaPos, obj, NotPirata, True

    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
        Call MakeObj(obj, Pos.Map, NuevaPos.X, NuevaPos.Y)
    End If

    TirarItemAlPiso = NuevaPos

    Exit Function
ErrHandler:

End Function

Public Sub NPC_TIRAR_ITEMS(ByRef npc As npc)

    'TIRA TODOS LOS ITEMS DEL NPC
    On Error Resume Next

    If npc.Invent.NroItems > 0 Then
    
        Dim i     As Byte

        Dim MiObj As obj
    
        For i = 1 To MAX_INVENTORY_SLOTS
    
            If npc.Invent.Object(i).ObjIndex > 0 Then
                MiObj.Amount = npc.Invent.Object(i).Amount
                MiObj.ObjIndex = npc.Invent.Object(i).ObjIndex
                Call TirarItemAlPiso(npc.Pos, MiObj)

            End If
      
        Next i

    End If

End Sub

Function QuedanItems(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Boolean

    On Error Resume Next

    'Call LogTarea("Function QuedanItems npcindex:" & NpcIndex & " objindex:" & ObjIndex)

    Dim i As Integer

    If Npclist(NpcIndex).Invent.NroItems > 0 Then

        For i = 1 To MAX_INVENTORY_SLOTS

            If Npclist(NpcIndex).Invent.Object(i).ObjIndex = ObjIndex Then
                QuedanItems = True
                Exit Function

            End If

        Next

    End If

    QuedanItems = False

End Function

Function EncontrarCant(ByVal NpcIndex As Integer, ByVal ObjIndex As Integer) As Integer

    On Error Resume Next

    'Devuelve la cantidad original del obj de un npc

    Dim ln As String, npcfile As String

    Dim i  As Integer

    'If Npclist(NpcIndex).Numero > 499 Then
    '    npcfile = DatPath & "NPCs-HOSTILES.dat"
    'Else
    npcfile = DatPath & "NPCs.dat"
    'End If
 
    For i = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & i)

        If ObjIndex = val(ReadField(1, ln, 45)) Then
            EncontrarCant = val(ReadField(2, ln, 45))
            Exit Function

        End If

    Next
                   
    EncontrarCant = 0

End Function

Sub ResetNpcInv(ByVal NpcIndex As Integer)

    On Error Resume Next

    Dim i As Integer

    Npclist(NpcIndex).Invent.NroItems = 0

    For i = 1 To MAX_INVENTORY_SLOTS
        Npclist(NpcIndex).Invent.Object(i).ObjIndex = 0
        Npclist(NpcIndex).Invent.Object(i).Amount = 0
    Next i

    Npclist(NpcIndex).InvReSpawn = 0

End Sub

Sub QuitarNpcInvItem(ByVal NpcIndex As Integer, ByVal slot As Byte, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarNpcInvItem_Err
        

        Dim ObjIndex As Integer

100     ObjIndex = Npclist(NpcIndex).Invent.Object(slot).ObjIndex

        'Quita un Obj
102     If ObjData(Npclist(NpcIndex).Invent.Object(slot).ObjIndex).Crucial = 0 Then
104         Npclist(NpcIndex).Invent.Object(slot).Amount = Npclist(NpcIndex).Invent.Object(slot).Amount - Cantidad
        
106         If Npclist(NpcIndex).Invent.Object(slot).Amount <= 0 Then
108             Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
110             Npclist(NpcIndex).Invent.Object(slot).ObjIndex = 0
112             Npclist(NpcIndex).Invent.Object(slot).Amount = 0

114             If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
116                 Call CargarInvent(NpcIndex) 'Reponemos el inventario

                End If

            End If

        Else
118         Npclist(NpcIndex).Invent.Object(slot).Amount = Npclist(NpcIndex).Invent.Object(slot).Amount - Cantidad
        
120         If Npclist(NpcIndex).Invent.Object(slot).Amount <= 0 Then
122             Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
124             Npclist(NpcIndex).Invent.Object(slot).ObjIndex = 0
126             Npclist(NpcIndex).Invent.Object(slot).Amount = 0
            
128             If Not QuedanItems(NpcIndex, ObjIndex) Then

                    Dim NoEsdeAca As Integer

130                 NoEsdeAca = EncontrarCant(NpcIndex, ObjIndex)

132                 If NoEsdeAca <> 0 Then
134                     Npclist(NpcIndex).Invent.Object(slot).ObjIndex = ObjIndex
136                     Npclist(NpcIndex).Invent.Object(slot).Amount = EncontrarCant(NpcIndex, ObjIndex)
138                     Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1

                    End If

                End If
            
140             If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
142                 Call CargarInvent(NpcIndex) 'Reponemos el inventario

                End If

            End If
    
        End If

        
        Exit Sub

QuitarNpcInvItem_Err:
        Call RegistrarError(Err.Number, Err.description, "InvNpc.QuitarNpcInvItem", Erl)
        Resume Next
        
End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)
        
        On Error GoTo CargarInvent_Err
        

        'Vuelve a cargar el inventario del npc NpcIndex
        Dim LoopC   As Integer

        Dim ln      As String

        Dim npcfile As String

        'If Npclist(NpcIndex).Numero > 499 Then
        '    npcfile = DatPath & "NPCs-HOSTILES.dat"
        'Else
100     npcfile = DatPath & "NPCs.dat"
        'End If

102     Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "NROITEMS"))

104     For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
106         ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & LoopC)
108         Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
110         Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
    
112     Next LoopC

        
        Exit Sub

CargarInvent_Err:
        Call RegistrarError(Err.Number, Err.description, "InvNpc.CargarInvent", Erl)
        Resume Next
        
End Sub

Public Sub NpcDropeo(ByRef npc As npc, ByRef UserIndex As Integer)

    On Error GoTo ErrHandler

    If npc.NumQuiza = 0 Then Exit Sub
    If DropActive = 0 Then Exit Sub 'Esta el Dropeo activado?

    Dim Dropeo       As obj

    Dim Probabilidad As Long

    Dim objRandom    As Byte

    Dim ProbTiro     As Byte

    Dim nfile        As Integer

    If npc.QuizaProb = 0 Then
        Probabilidad = RandomNumber(1, DropMult) 'Tiro Item?
    Else
        Probabilidad = RandomNumber(1, npc.QuizaProb) 'Tiro Item?

    End If

    If UserList(UserIndex).Invent.MagicoObjIndex = 383 Then
        If npc.QuizaProb = 0 Then
            Probabilidad = RandomNumber(1, DropMult / 2) 'Tiro Item?
        Else
            Probabilidad = RandomNumber(1, npc.QuizaProb / 2) 'Tiro Item?

        End If

    End If

    If Probabilidad <> 1 Then Exit Sub

    objRandom = RandomNumber(1, npc.NumQuiza) 'Que item puede ser que tire?

    Dim obj      As Integer

    Dim Cantidad As Integer

    obj = val(ReadField(1, npc.QuizaDropea(objRandom), Asc("-")))
    Cantidad = val(ReadField(2, npc.QuizaDropea(objRandom), Asc("-")))

    Dropeo.Amount = Cantidad 'Cantidad
    Dropeo.ObjIndex = obj 'NUMERO DEL ITEM EN EL OBJ.DAT
    Call TirarItemAlPiso(npc.Pos, Dropeo)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(FXSound.Dropeo_Sound, npc.Pos.X, npc.Pos.Y))
        
    'nfile = FreeFile ' obtenemos un canal
    'Open App.Path & "\logs\Dropeo de items.log" For Append Shared As #nfile
    'Print #nfile, "El dia " & Date & " a las " & Time & " al usuario " & UserList(UserIndex).Name & " se le a dropiado el objeto " & ObjData(obj).Name & "."
    ' Close #nfile
    
    Exit Sub

ErrHandler:
    Call LogError("Error al dropear el item " & ObjData(npc.QuizaDropea(objRandom)).name & ", al usuario " & UserList(UserIndex).name & ". " & Err.description & ".")

End Sub


Public Sub DropObjQuest(ByRef npc As npc, ByRef UserIndex As Integer)
'Dropeo por Quest
'Ladder
'3/12/2020
    On Error GoTo ErrHandler

    If npc.DropQuest = "" Then Exit Sub
        
    
    Dim Dropeo       As obj
    Dim QuestIndex As Integer
    Dim ObjIndex As Integer
    Dim Amount As Integer
    Dim Probabilidad As Byte

    QuestIndex = val(ReadField(1, npc.DropQuest, Asc("-")))
    ObjIndex = val(ReadField(2, npc.DropQuest, Asc("-")))
    Amount = val(ReadField(3, npc.DropQuest, Asc("-")))
    Probabilidad = val(ReadField(4, npc.DropQuest, Asc("-")))
    
    If QuestIndex = 0 Then Exit Sub
    
    If TieneQuest(UserIndex, QuestIndex) = 0 Then Exit Sub
    

    Probabilidad = RandomNumber(1, Probabilidad) 'Tiro Item?


    If Probabilidad <> 1 Then Exit Sub

    Dropeo.Amount = Amount
    Dropeo.ObjIndex = ObjIndex
    Call TirarItemAlPiso(npc.Pos, Dropeo)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(FXSound.Dropeo_Sound, npc.Pos.X, npc.Pos.Y))

    Exit Sub

ErrHandler:
    Call LogError("Error DropObjQuest al dropear el item " & ObjData(ObjIndex).name & ", al usuario " & UserList(UserIndex).name & ". " & Err.description & ".")

End Sub

