Attribute VB_Name = "modSistemaComercio"
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
Enum eModoComercio
    Compra = 1
    Venta = 2
End Enum

Public Const REDUCTOR_PRECIOVENTA As Byte = 3

Public Sub Comercio(ByVal Modo As eModoComercio, ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Slot As Integer, ByVal Cantidad As Integer)
    On Error GoTo Comercio_Err
    Dim precio           As Long
    Dim Objeto           As t_Obj
    Dim objquedo         As t_Obj
    Dim precioenvio      As Single
    Dim NpcSlot          As Integer
    Dim Objeto_A_Comprar As t_UserOBJ
    If Cantidad < 1 Or Slot < 1 Then Exit Sub
    If Modo = eModoComercio.Compra Then
        If Slot > MAX_INVENTORY_SLOTS Then
            Exit Sub
        ElseIf Cantidad > GetMaxInvOBJ() Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1746, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)) 'Msg1746=¬1 ha sido baneado por el sistema anti-cheats.
            Call Ban(UserList(UserIndex).name, "Sistema Anti Cheats", "Intentar hackear el sistema de comercio. Quiso comprar demasiados items:" & Cantidad)
            UserList(UserIndex).flags.Ban = 1
            Call WriteShowMessageBox(UserIndex, 1757, vbNullString) 'Msg1751=Has sido baneado por el Sistema AntiCheat.
            Call CloseSocket(UserIndex)
            Exit Sub
        ElseIf Not NpcList(NpcIndex).invent.Object(Slot).amount > 0 Then
            Exit Sub
        End If
        Objeto_A_Comprar = NpcList(NpcIndex).invent.Object(Slot)
        If Objeto_A_Comprar.ObjIndex = 0 Then Exit Sub
        If ObjData(Objeto_A_Comprar.ObjIndex).Crucial = 0 Then
            If Cantidad > Objeto_A_Comprar.amount Then
                Cantidad = Objeto_A_Comprar.amount
            End If
        ElseIf ObjData(Objeto_A_Comprar.ObjIndex).Crucial = 1 Then
            'si es un item que vende el NPC le dejo comprar todo lo que quiera
            If Not NpcSellsItem(NpcList(NpcIndex).Numero, Objeto_A_Comprar.ObjIndex) Then
                If Cantidad > Objeto_A_Comprar.amount Then
                    Cantidad = Objeto_A_Comprar.amount
                End If
            End If
        End If
        'NpcSellsItem
        Objeto.amount = Cantidad
        Objeto.ObjIndex = Objeto_A_Comprar.ObjIndex
        'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
        'Es decir, 1.1 = 2, por lo cual se hace de la siguiente forma Precio = Clng(PrecioFinal + 0.5) Siempre va a darte el proximo numero. O el "Techo" (MarKoxX)
        precio = Ceil(ObjData(Objeto_A_Comprar.ObjIndex).Valor / Descuento(UserIndex) * Cantidad)
        If UserList(UserIndex).Stats.GLD < precio Then
            'Msg1082= No tienes suficiente dinero.
            Call WriteLocaleMsg(UserIndex, 1082, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not MeterItemEnInventario(UserIndex, Objeto) Then Exit Sub
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - precio
        Call WriteUpdateGold(UserIndex)
        Call QuitarNpcInvItem(NpcIndex, Slot, Cantidad)
        Call UpdateNpcInvToAll(False, NpcIndex, Slot)
        'Agregado para que no se vuelvan a vender las llaves si se recargan los .dat.
        If ObjData(Objeto.ObjIndex).OBJType = otKeys Then
            Call WriteVar(DatPath & "NPCs.dat", "NPC" & NpcList(NpcIndex).Numero, "obj" & Slot, Objeto.ObjIndex & "-0")
            Call logVentaCasa(UserList(UserIndex).name & " compro " & ObjData(Objeto.ObjIndex).name)
        End If
    ElseIf Modo = eModoComercio.Venta Then
        If Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
        If Cantidad > UserList(UserIndex).invent.Object(Slot).amount Then Cantidad = UserList(UserIndex).invent.Object(Slot).amount
        Objeto.amount = Cantidad
        Objeto.ObjIndex = UserList(UserIndex).invent.Object(Slot).ObjIndex
        If Objeto.ObjIndex = 0 Then
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Newbie = 1 Then
            'Msg1083= Lo siento, no comercio objetos para newbies.
            Call WriteLocaleMsg(UserIndex, 1083, e_FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Destruye = 1 Then
            'Msg1084= Lo siento, no puedo comprarte ese item.
            Call WriteLocaleMsg(UserIndex, 1084, e_FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        ElseIf ObjData(Objeto.ObjIndex).Instransferible = 1 Then
            'Msg1085= Lo siento, no puedo comprarte ese item.
            Call WriteLocaleMsg(UserIndex, 1085, e_FontTypeNames.FONTTYPE_TALK)
            Exit Sub
        ElseIf ((NpcList(NpcIndex).TipoItems <> ObjData(Objeto.ObjIndex).OBJType And NpcList(NpcIndex).TipoItems <> e_OBJType.otElse) Or Objeto.ObjIndex = iORO) Then
            'Agrego que si vende el item, lo compre tambien.
            Dim LoVende As Boolean
            Dim i       As Integer
            For i = 1 To NpcList(NpcIndex).invent.NroItems
                If NpcList(NpcIndex).invent.Object(i).ObjIndex = Objeto.ObjIndex Then
                    LoVende = True
                End If
            Next i
            If Not LoVende Then
                'Msg1086= Lo siento, no estoy interesado en este tipo de objetos.
                Call WriteLocaleMsg(UserIndex, 1086, e_FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
        ElseIf UserList(UserIndex).invent.Object(Slot).amount < 0 Or Cantidad = 0 Then
            Exit Sub
        ElseIf Slot < LBound(UserList(UserIndex).invent.Object()) Or Slot > UBound(UserList(UserIndex).invent.Object()) Then
            Exit Sub
        ElseIf UserList(UserIndex).flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios) Then
            ' Msg767=No podés vender items.
            Call WriteLocaleMsg(UserIndex, 767, e_FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        Call QuitarUserInvItem(UserIndex, Slot, Cantidad)
        Call UpdateUserInv(False, UserIndex, Slot)
        'Precio = Round(ObjData(Objeto.ObjIndex).valor / REDUCTOR_PRECIOVENTA * Cantidad, 0)
        precio = Fix(SalePrice(Objeto.ObjIndex, UserIndex) * Cantidad)
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + precio
        If UserList(UserIndex).Stats.GLD > MAXORO Then UserList(UserIndex).Stats.GLD = MAXORO
        Call WriteUpdateGold(UserIndex)
        If Not IsFeatureEnabled("destroy_npc_bought_items") Then
            NpcSlot = SlotEnNPCInv(NpcIndex, Objeto.ObjIndex, Objeto.Amount)
            If NpcSlot > 0 And NpcSlot <= MAX_INVENTORY_SLOTS Then 'Slot valido
                ' Saque este incremento de SlotEnNPCInv porque me parece mejor manejarlo junto con el resto de las asignaciones
                If NpcList(NpcIndex).invent.Object(NpcSlot).ObjIndex = 0 Then
                    NpcList(NpcIndex).invent.NroItems = NpcList(NpcIndex).invent.NroItems + 1
                End If
                'Mete el obj en el slot
                NpcList(NpcIndex).invent.Object(NpcSlot).ObjIndex = Objeto.ObjIndex
                NpcList(NpcIndex).invent.Object(NpcSlot).Amount = NpcList(NpcIndex).invent.Object(NpcSlot).Amount + Objeto.Amount
                If NpcList(NpcIndex).invent.Object(NpcSlot).Amount > GetMaxInvOBJ() Then
                    NpcList(NpcIndex).invent.Object(NpcSlot).Amount = GetMaxInvOBJ()
                End If
                Call UpdateNpcInvToAll(False, NpcIndex, NpcSlot)
            End If
        End If
    End If
    Call SubirSkill(UserIndex, e_Skill.Comerciar)
    Exit Sub
Comercio_Err:
    Call TraceError(Err.Number, Err.Description, "modSistemaComercio.Comercio", Erl)
End Sub

Public Sub IniciarComercioNPC(ByVal UserIndex As Integer)
    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 2/8/06
    '*************************************************
    On Error GoTo IniciarComercioNPC_Err
    If Not IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
        ' Msg770=El comerciante no está disponible.
        Call WriteLocaleMsg(UserIndex, 770, e_FontTypeNames.FONTTYPE_WARNING)
        Exit Sub
    End If
    Call UpdateNpcInv(True, UserIndex, UserList(UserIndex).flags.TargetNPC.ArrayIndex, 0)
    UserList(UserIndex).flags.Comerciando = True
    Call WriteCommerceInit(UserIndex)
    Exit Sub
IniciarComercioNPC_Err:
    Call TraceError(Err.Number, Err.Description, "modSistemaComercio.IniciarComercioNPC", Erl)
End Sub

Private Function SlotEnNPCInv(ByVal NpcIndex As Integer, ByVal Objeto As Integer, ByVal Cantidad As Integer) As Integer
    '*************************************************
    'Devuelve el slot en el cual se debe agregar el nuevo objeto, o 0 si no se debe asignar en ningun lado
    '*************************************************
    On Error GoTo SlotEnNPCInv_Err
    With NpcList(NpcIndex).invent
        Dim Slot            As Byte
        Dim matchingSlots   As New Collection
        Dim firstEmptySpace As Integer
        ' Recorro el inventario buscando el objeto a agregar y espacios vacios
        firstEmptySpace = 0
        For Slot = 1 To MAX_INVENTORY_SLOTS
            If .Object(Slot).ObjIndex = Objeto Then
                matchingSlots.Add (Slot)
            ElseIf .Object(Slot).ObjIndex = 0 And firstEmptySpace = 0 Then
                firstEmptySpace = Slot
            End If
        Next Slot
        ' Recorro los slots donde hay objetos que matcheen con el objeto a agregar y si alguno tiene espacio, lo agrego ahi. Si no, se descarta
        If matchingSlots.count <> 0 Then
            Dim i As Variant
            For Each i In matchingSlots
                If .Object(i).amount < GetMaxInvOBJ() Then
                    SlotEnNPCInv = i
                    Exit Function
                End If
            Next i
            SlotEnNPCInv = 0
            Exit Function
        End If
        SlotEnNPCInv = firstEmptySpace
        Exit Function
    End With
SlotEnNPCInv_Err:
    Call TraceError(Err.Number, Err.Description, "modSistemaComercio.SlotEnNPCInv", Erl)
End Function

Private Function Descuento(ByVal UserIndex As Integer) As Single
    '*************************************************
    'Author: Nacho (Integer)
    'Last modified: 2/8/06
    '*************************************************
    On Error GoTo Descuento_Err
    Descuento = 1 + UserList(UserIndex).Stats.UserSkills(e_Skill.Comerciar) / 100
    Exit Function
Descuento_Err:
    Call TraceError(Err.Number, Err.Description, "modSistemaComercio.Descuento", Erl)
End Function

''
' Update the inventory of the Npc to the user
'
' @param updateAll if is needed to update all
' @param npcIndex The index of the NPC
Private Sub UpdateNpcInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal Slot As Byte)
    On Error GoTo EnviarNpcInv_Err
    Dim obj   As t_Obj
    Dim LoopC As Long
    Dim Desc  As Single
    Dim val   As Single
    Desc = Descuento(UserIndex)
    'Actualiza un solo slot
    If Not UpdateAll Then
        With NpcList(NpcIndex).invent.Object(Slot)
            obj.ObjIndex = .ObjIndex
            obj.amount = .amount
            If .ObjIndex > 0 Then
                val = Ceil(ObjData(.ObjIndex).Valor / Desc)
            End If
            Call WriteChangeNPCInventorySlot(UserIndex, Slot, obj, val)
        End With
    Else
        'Actualiza todos los slots
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            With NpcList(NpcIndex).invent.Object(LoopC)
                obj.ObjIndex = .ObjIndex
                obj.amount = .amount
                If .ObjIndex > 0 Then
                    val = Ceil(ObjData(.ObjIndex).Valor / Desc)
                End If
                Call WriteChangeNPCInventorySlot(UserIndex, LoopC, obj, val)
            End With
        Next LoopC
    End If
    Exit Sub
EnviarNpcInv_Err:
    Call TraceError(Err.Number, Err.Description, "modSistemaComercio.UpdateNpcInv", Erl)
End Sub


Public Sub UpdateNpcInvToAll(ByVal UpdateAll As Boolean, ByVal NpcIndex As Integer, ByVal Slot As Byte)
    On Error GoTo ErrHandler:
    Dim LoopC As Long
    ' Recorremos todos los usuarios
    For LoopC = 1 To LastUser
        With UserList(LoopC)
            ' Si esta comerciando
            If .flags.Comerciando Then
                ' Si el ultimo NPC que cliqueo es el que hay que actualizar
                If .flags.TargetNPC.ArrayIndex = NpcIndex Then
                    ' Actualizamos el inventario del NPC
                    Call UpdateNpcInv(UpdateAll, LoopC, NpcIndex, Slot)
                End If
            End If
        End With
    Next
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "modSistemaComercio.UpdateNpcInvToAll")
End Sub

Public Function SalePrice(ByVal ObjIndex As Integer, Optional ByVal UserIndex As Integer = 0) As Single
    On Error GoTo SalePrice_Err
    If ObjIndex < 1 Or ObjIndex > UBound(ObjData) Then Exit Function
    If ItemNewbie(ObjIndex) Then Exit Function
    Dim denom As Single
    denom = REDUCTOR_PRECIOVENTA
    If UserIndex > 0 Then
        If UserList(UserIndex).clase = e_Class.Trabajador Then
            denom = denom - (UserList(UserIndex).Stats.ELV * 0.025) '0.25/10 = 0.025
            If denom < 2 Then denom = 2 'clamp: evita div0 y negativos
        End If
    End If
    SalePrice = ObjData(ObjIndex).Valor / denom
    Exit Function
SalePrice_Err:
    Call TraceError(Err.Number, Err.Description, "modSistemaComercio.SalePrice", Erl)
End Function
