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

    On Error GoTo Errhandler

    Dim NuevaPos As WorldPos

    NuevaPos.x = 0
    NuevaPos.Y = 0
    
    Tilelibre Pos, NuevaPos, obj, NotPirata, True

    If NuevaPos.x <> 0 And NuevaPos.Y <> 0 Then
        Call MakeObj(obj, Pos.Map, NuevaPos.x, NuevaPos.Y)

    End If

    TirarItemAlPiso = NuevaPos

    Exit Function
Errhandler:

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

    Dim ObjIndex As Integer

    ObjIndex = Npclist(NpcIndex).Invent.Object(slot).ObjIndex

    'Quita un Obj
    If ObjData(Npclist(NpcIndex).Invent.Object(slot).ObjIndex).Crucial = 0 Then
        Npclist(NpcIndex).Invent.Object(slot).Amount = Npclist(NpcIndex).Invent.Object(slot).Amount - Cantidad
        
        If Npclist(NpcIndex).Invent.Object(slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(slot).Amount = 0

            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
                Call CargarInvent(NpcIndex) 'Reponemos el inventario

            End If

        End If

    Else
        Npclist(NpcIndex).Invent.Object(slot).Amount = Npclist(NpcIndex).Invent.Object(slot).Amount - Cantidad
        
        If Npclist(NpcIndex).Invent.Object(slot).Amount <= 0 Then
            Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems - 1
            Npclist(NpcIndex).Invent.Object(slot).ObjIndex = 0
            Npclist(NpcIndex).Invent.Object(slot).Amount = 0
            
            If Not QuedanItems(NpcIndex, ObjIndex) Then

                Dim NoEsdeAca As Integer

                NoEsdeAca = EncontrarCant(NpcIndex, ObjIndex)

                If NoEsdeAca <> 0 Then
                    Npclist(NpcIndex).Invent.Object(slot).ObjIndex = ObjIndex
                    Npclist(NpcIndex).Invent.Object(slot).Amount = EncontrarCant(NpcIndex, ObjIndex)
                    Npclist(NpcIndex).Invent.NroItems = Npclist(NpcIndex).Invent.NroItems + 1

                End If

            End If
            
            If Npclist(NpcIndex).Invent.NroItems = 0 And Npclist(NpcIndex).InvReSpawn <> 1 Then
                Call CargarInvent(NpcIndex) 'Reponemos el inventario

            End If

        End If
    
    End If

End Sub

Sub CargarInvent(ByVal NpcIndex As Integer)

    'Vuelve a cargar el inventario del npc NpcIndex
    Dim LoopC   As Integer

    Dim ln      As String

    Dim npcfile As String

    'If Npclist(NpcIndex).Numero > 499 Then
    '    npcfile = DatPath & "NPCs-HOSTILES.dat"
    'Else
    npcfile = DatPath & "NPCs.dat"
    'End If

    Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "NROITEMS"))

    For LoopC = 1 To Npclist(NpcIndex).Invent.NroItems
        ln = GetVar(npcfile, "NPC" & Npclist(NpcIndex).Numero, "Obj" & LoopC)
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
    
    Next LoopC

End Sub

Public Sub NpcDropeo(ByRef npc As npc, ByRef UserIndex As Integer)

    On Error GoTo Errhandler

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
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(FXSound.Dropeo_Sound, npc.Pos.x, npc.Pos.Y))
        
    'nfile = FreeFile ' obtenemos un canal
    'Open App.Path & "\logs\Dropeo de items.log" For Append Shared As #nfile
    'Print #nfile, "El dia " & Date & " a las " & Time & " al usuario " & UserList(UserIndex).Name & " se le a dropiado el objeto " & ObjData(obj).Name & "."
    ' Close #nfile
    
    Exit Sub

Errhandler:
    Call LogError("Error al dropear el item " & ObjData(npc.QuizaDropea(objRandom)).name & ", al usuario " & UserList(UserIndex).name & ". " & Err.description & ".")

End Sub

