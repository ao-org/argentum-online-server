Attribute VB_Name = "InvUsuario"
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
Private Const PATREON_HEAD = 900

Public Function GetMaxInvOBJ() As Integer
GetMaxInvOBJ = CInt(SvrConfig.GetValue("MaxInventoryObjs"))
End Function


Public Function IsObjecIndextInInventory(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    On Error GoTo IsObjecIndextInInventory_Err
    Debug.Assert UserIndex >= LBound(UserList) And UserIndex <= UBound(UserList)
    ' If no match is found, return False
    IsObjecIndextInInventory = False
    Dim i                 As Integer
    Dim maxItemsInventory As Integer
    Dim currentObjIndex   As Integer
    With UserList(UserIndex)
        maxItemsInventory = get_num_inv_slots_from_tier(.Stats.tipoUsuario)
        ' Search inventory for the object
        For i = 1 To maxItemsInventory
            currentObjIndex = .invent.Object(i).ObjIndex
            If currentObjIndex = ObjIndex Then
                IsObjecIndextInInventory = True
                Exit Function
            End If
        Next i
    End With
    Exit Function
IsObjecIndextInInventory_Err:
    Call TraceError(Err.Number, Err.Description, "IsObjecIndextInInventory", Erl)
End Function

Public Function get_object_amount_from_inventory(ByVal user_index, ByVal obj_index As Integer) As Integer
    On Error GoTo get_object_amount_from_inventory_Err
    Debug.Assert user_index >= LBound(UserList) And user_index <= UBound(UserList)
    ' If no match is found, return 0
    get_object_amount_from_inventory = 0
    Dim i                 As Integer
    Dim maxItemsInventory As Integer
    With UserList(user_index)
        maxItemsInventory = get_num_inv_slots_from_tier(.Stats.tipoUsuario)
        ' Search inventory for the object
        For i = 1 To maxItemsInventory
            If .invent.Object(i).ObjIndex = obj_index Then
                get_object_amount_from_inventory = .invent.Object(i).amount
                Exit Function
            End If
        Next i
    End With
    Exit Function
get_object_amount_from_inventory_Err:
    Call TraceError(Err.Number, Err.Description, "get_object_amount_from_inventory", Erl)
End Function

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean
    On Error GoTo TieneObjetosRobables_Err
    Dim i        As Integer
    Dim ObjIndex As Integer
    If UserList(UserIndex).CurrentInventorySlots > 0 Then
        For i = 1 To UserList(UserIndex).CurrentInventorySlots
            ObjIndex = UserList(UserIndex).invent.Object(i).ObjIndex
            If ObjIndex > 0 Then
                If (ObjData(ObjIndex).OBJType <> e_OBJType.otKeys And ObjData(ObjIndex).OBJType <> e_OBJType.otShips And ObjData(ObjIndex).OBJType <> e_OBJType.otSaddles And _
                        ObjData(ObjIndex).OBJType <> e_OBJType.otDonator And ObjData(ObjIndex).OBJType <> e_OBJType.otRecallStones) Then
                    TieneObjetosRobables = True
                    Exit Function
                End If
            End If
        Next i
    End If
    Exit Function
TieneObjetosRobables_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.TieneObjetosRobables", Erl)
End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional Slot As Byte) As Boolean
    On Error GoTo manejador
    Dim Flag As Boolean
    If Slot <> 0 Then
        If UserList(UserIndex).invent.Object(Slot).Equipped Then
            ClasePuedeUsarItem = True
            Exit Function
        End If
    End If
    If EsGM(UserIndex) Then
        ClasePuedeUsarItem = True
        Exit Function
    End If
    Dim i As Integer
    For i = 1 To NUMCLASES
        If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
            ClasePuedeUsarItem = False
            Exit Function
        End If
    Next i
    ClasePuedeUsarItem = True
    Exit Function
manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Function RazaPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional Slot As Byte) As Boolean
    On Error GoTo RazaPuedeUsarItem_Err
    Dim Objeto As t_ObjData, i As Long
    Objeto = ObjData(ObjIndex)
    If EsGM(UserIndex) Then
        RazaPuedeUsarItem = True
        Exit Function
    End If
    For i = 1 To NUMRAZAS
        If Objeto.RazaProhibida(i) = UserList(UserIndex).raza Then
            RazaPuedeUsarItem = False
            Exit Function
        End If
    Next i
    ' Si el objeto no define una raza en particular
    If Objeto.RazaDrow + Objeto.RazaElfa + Objeto.RazaEnana + Objeto.RazaGnoma + Objeto.RazaHumana + Objeto.RazaOrca = 0 Then
        RazaPuedeUsarItem = True
    Else ' El objeto esta definido para alguna raza en especial
        Select Case UserList(UserIndex).raza
            Case e_Raza.Humano
                RazaPuedeUsarItem = Objeto.RazaHumana > 0
            Case e_Raza.Elfo
                RazaPuedeUsarItem = Objeto.RazaElfa > 0
            Case e_Raza.Drow
                RazaPuedeUsarItem = Objeto.RazaDrow > 0
            Case e_Raza.Orco
                RazaPuedeUsarItem = Objeto.RazaOrca > 0
            Case e_Raza.Gnomo
                RazaPuedeUsarItem = Objeto.RazaGnoma > 0
            Case e_Raza.Enano
                RazaPuedeUsarItem = Objeto.RazaEnana > 0
        End Select
    End If
    If RazaPuedeUsarItem And Objeto.OBJType = e_OBJType.otArmor Then
        RazaPuedeUsarItem = ObtenerRopaje(UserIndex, Objeto) <> 0
    End If
    Exit Function
RazaPuedeUsarItem_Err:
    LogError ("Error en RazaPuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
    On Error GoTo QuitarNewbieObj_Err
    Dim j As Integer
    If UserList(UserIndex).CurrentInventorySlots > 0 Then
        For j = 1 To UserList(UserIndex).CurrentInventorySlots
            If UserList(UserIndex).invent.Object(j).ObjIndex > 0 Then
                If ObjData(UserList(UserIndex).invent.Object(j).ObjIndex).Newbie = 1 Then
                    Call QuitarUserInvItem(UserIndex, j, GetMaxInvOBJ())
                    Call UpdateUserInv(False, UserIndex, j)
                End If
            End If
        Next j
    End If
    ' Eliminar items newbie de la boveda
    For j = 1 To MAX_BANCOINVENTORY_SLOTS
        If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
            If ObjData(UserList(UserIndex).BancoInvent.Object(j).ObjIndex).Newbie = 1 Then
                UserList(UserIndex).BancoInvent.Object(j).ObjIndex = 0
                UserList(UserIndex).BancoInvent.Object(j).amount = 0
                UserList(UserIndex).BancoInvent.Object(j).ElementalTags = 0
                Call UpdateBanUserInv(False, UserIndex, j, "QuitarNewbieObj")
            End If
        End If
    Next j
    'Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
    'Mandamos a la Isla de la Fortuna
    Call WarpUserChar(UserIndex, Renacimiento.Map, Renacimiento.x, Renacimiento.y, True)
    ' Msg671=Has dejado de ser Newbie.
    Call WriteLocaleMsg(UserIndex, 671, e_FontTypeNames.FONTTYPE_INFO)
    Exit Sub
QuitarNewbieObj_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.QuitarNewbieObj", Erl)
End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)
    On Error GoTo LimpiarInventario_Err
    Dim j As Integer
    With UserList(UserIndex)
        If .CurrentInventorySlots > 0 Then
            For j = 1 To .CurrentInventorySlots
                If j > 0 And j <= UBound(.invent.Object) Then 'Make sure the slot is valid
                    .invent.Object(j).ObjIndex = 0
                    .invent.Object(j).amount = 0
                    .invent.Object(j).Equipped = 0
                End If
            Next
        End If
        .invent.NroItems = 0
        .invent.EquippedArmorObjIndex = 0
        .invent.EquippedArmorSlot = 0
        .invent.EquippedWeaponObjIndex = 0
        .invent.EquippedWeaponSlot = 0
        .invent.EquippedWorkingToolObjIndex = 0
        .invent.EquippedWorkingToolSlot = 0
        .invent.EquippedHelmetObjIndex = 0
        .invent.EquippedHelmetSlot = 0
        .invent.EquippedShieldObjIndex = 0
        .invent.EquippedShieldSlot = 0
        .invent.EquippedRingAccesoryObjIndex = 0
        .invent.EquippedRingAccesorySlot = 0
        .invent.EquippedRingAccesoryObjIndex = 0
        .invent.EquippedRingAccesorySlot = 0
        .invent.EquippedMunitionObjIndex = 0
        .invent.EquippedMunitionSlot = 0
        .invent.EquippedShipObjIndex = 0
        .invent.EquippedShipSlot = 0
        .invent.EquippedSaddleObjIndex = 0
        .invent.EquippedSaddleSlot = 0
        .invent.EquippedAmuletAccesoryObjIndex = 0
        .invent.EquippedAmuletAccesorySlot = 0
        .invent.EquippedBackpackObjIndex = 0
        .invent.EquippedBackpackSlot = 0
    End With
    Exit Sub
LimpiarInventario_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.LimpiarInventario", Erl)
End Sub

Sub ResetUserSkinsInventory(ByVal UserIndex As Integer)
Dim i                           As Byte
    On Error GoTo ResetUserSkinsInventory_Error
    With UserList(UserIndex)

        For i = 1 To MAX_SKINSINVENTORY_SLOTS
            .Invent_Skins.Object(i).ObjIndex = 0
            .Invent_Skins.Object(i).Equipped = False
            .Invent_Skins.Object(i).Type = 0
            .Invent_Skins.Object(i).Deleted = False
        Next i

        .Invent_Skins.ObjIndexArmourEquipped = 0
        .Invent_Skins.ObjIndexHelmetEquipped = 0
        .Invent_Skins.ObjIndexWeaponEquipped = 0
        .Invent_Skins.ObjIndexShieldEquipped = 0
        .Invent_Skins.ObjIndexWindsEquipped = 0
        .Invent_Skins.ObjIndexBoatEquipped = 0
        .Invent_Skins.ObjIndexBackpackEquipped = 0
        .Invent_Skins.SlotArmourEquipped = 0
        .Invent_Skins.SlotHelmetEquipped = 0
        .Invent_Skins.SlotWeaponEquipped = 0
        .Invent_Skins.SlotShieldEquipped = 0
        .Invent_Skins.SlotWindsEquipped = 0
        .Invent_Skins.SlotBoatEquipped = 0
        .Invent_Skins.SlotBackpackEquipped = 0
        .Invent_Skins.count = 0
    End With
    On Error GoTo 0
    Exit Sub
ResetUserSkinsInventory_Error:
    Call Logging.TraceError(Err.Number, Err.Description, "InvUsuario.ResetUserSkinsInventory", Erl())
End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 23/01/2007
    '23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
    '***************************************************
    On Error GoTo ErrHandler
    Dim OriginalAmount As Long
    OriginalAmount = Cantidad
    With UserList(UserIndex)
        ' GM's (excepto Dioses y Admins) no pueden tirar oro
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
            Call LogGM(.name, " trató de tirar " & PonerPuntos(Cantidad) & " de oro en " & .pos.Map & "-" & .pos.x & "-" & .pos.y)
            Exit Sub
        End If
        ' Si el usuario tiene ORO, entonces lo tiramos
        If (Cantidad > 0) And (Cantidad <= .Stats.GLD) Then
            Dim i     As Byte
            Dim MiObj As t_Obj
            'info debug
            Dim loops As Long
            Do While (Cantidad > 0)
                If Cantidad > GetMaxInvOBJ() And .Stats.GLD > GetMaxInvOBJ() Then
                    MiObj.amount = GetMaxInvOBJ()
                    Cantidad = Cantidad - MiObj.amount
                Else
                    MiObj.amount = Cantidad
                    Cantidad = Cantidad - MiObj.amount
                End If
                MiObj.ObjIndex = iORO
                Dim AuxPos As t_WorldPos
                If .clase = e_Class.Pirat Then
                    AuxPos = TirarItemAlPiso(.pos, MiObj, False)
                Else
                    AuxPos = TirarItemAlPiso(.pos, MiObj, True)
                End If
                If AuxPos.x <> 0 And AuxPos.y <> 0 Then
                    .Stats.GLD = .Stats.GLD - MiObj.amount
                End If
                'info debug
                loops = loops + 1
                If loops > 100000 Then 'si entra aca y se cuelga mal el server revisen al tipo porque tiene much oro (NachoP) seguramente es dupero
                    Call LogError("Se ha superado el limite de iteraciones(100000) permitido en el Sub TirarOro() - posible Nacho P")
                    Exit Sub
                End If
            Loop
            ' Si es GM, registramos lo q hizo
            If EsGM(UserIndex) Then
                If MiObj.ObjIndex = iORO Then
                    Call LogGM(.name, "Tiro: " & PonerPuntos(OriginalAmount) & " monedas de oro.")
                Else
                    Call LogGM(.name, "Tiro cantidad:" & PonerPuntos(OriginalAmount) & " Objeto:" & ObjData(MiObj.ObjIndex).name)
                End If
            End If
            Call WriteUpdateGold(UserIndex)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.TirarOro", Erl())
End Sub

Public Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
    On Error GoTo QuitarUserInvItem_Err
    If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
    With UserList(UserIndex).invent.Object(Slot)
        If .amount <= Cantidad And .Equipped = 1 Then
            Call Desequipar(UserIndex, Slot)
        End If
        'Quita un objeto
        .amount = .amount - Cantidad
        '¿Quedan mas?
        If .amount <= 0 Then
            UserList(UserIndex).invent.NroItems = UserList(UserIndex).invent.NroItems - 1
            .ObjIndex = 0
            .amount = 0
        End If
        UserList(UserIndex).flags.ModificoInventario = True
    End With
    Exit Sub
QuitarUserInvItem_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.QuitarUserInvItem", Erl)
End Sub

Public Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
    On Error GoTo UpdateUserInv_Err
    Dim NullObj As t_UserOBJ
    Dim LoopC   As Byte
    'Actualiza un solo slot
    If Not UpdateAll And Slot > 0 Then
        'Actualiza el inventario
        If UserList(UserIndex).invent.Object(Slot).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).invent.Object(Slot))
        Else
            Call ChangeUserInv(UserIndex, Slot, NullObj)
        End If
        UserList(UserIndex).flags.ModificoInventario = True
    Else
        'Actualiza todos los slots
        If UserList(UserIndex).CurrentInventorySlots > 0 Then
            For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
                'Actualiza el inventario
                If LoopC > 0 And LoopC <= UBound(UserList(UserIndex).invent.Object) Then 'Make sure the slot is valid
                    If UserList(UserIndex).invent.Object(LoopC).ObjIndex > 0 Then
                        Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).invent.Object(LoopC))
                    Else
                        Call ChangeUserInv(UserIndex, LoopC, NullObj)
                    End If
                End If
            Next LoopC
        End If
    End If
    Exit Sub
UpdateUserInv_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.UpdateUserInv", Erl)
End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
    On Error GoTo DropObj_Err
    Dim obj As t_Obj
    If num > 0 Then
        With UserList(UserIndex)
            If num > .invent.Object(Slot).amount Then
                num = .invent.Object(Slot).amount
            End If
            obj.ObjIndex = .invent.Object(Slot).ObjIndex
            obj.amount = num
            obj.ElementalTags = .invent.Object(Slot).ElementalTags
            If Not CustomScenarios.UserCanDropItem(UserIndex, Slot, Map, x, y) Then
                Exit Sub
            End If
            If ObjData(obj.ObjIndex).Destruye = 0 Then
                Dim Suma As Long
                Suma = num + MapData(.pos.Map, x, y).ObjInfo.amount
                'Check objeto en el suelo
                If MapData(.pos.Map, x, y).ObjInfo.ObjIndex = 0 Or (MapData(.pos.Map, x, y).ObjInfo.ObjIndex = obj.ObjIndex And MapData(.pos.Map, x, y).ObjInfo.ElementalTags = _
                        obj.ElementalTags And Suma <= GetMaxInvOBJ()) Then
                    If Suma > GetMaxInvOBJ() Then
                        num = GetMaxInvOBJ() - MapData(.pos.Map, x, y).ObjInfo.amount
                    End If
                    ' Si sos Admin, Dios o Usuario, crea el objeto en el piso.
                    If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Admin Or e_PlayerType.Dios)) <> 0 Then
                        ' Tiramos el item al piso
                        Call MakeObj(obj, Map, x, y)
                    End If
                    Call CustomScenarios.UserDropItem(UserIndex, Slot, Map, x, y)
                    Call QuitarUserInvItem(UserIndex, Slot, num)
                    Call UpdateUserInv(False, UserIndex, Slot)
                    If .flags.jugando_captura = 1 Then
                        If Not InstanciaCaptura Is Nothing Then
                            Call InstanciaCaptura.tiraBandera(UserIndex, obj.ObjIndex)
                        End If
                    End If
                    If Not .flags.Privilegios And e_PlayerType.User Then
                        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) <> 0 Then
                            Call LogGM(.name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).name)
                        End If
                    End If
                Else
                    Call WriteLocaleMsg(UserIndex, 262, e_FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                Call QuitarUserInvItem(UserIndex, Slot, num)
                Call UpdateUserInv(False, UserIndex, Slot)
            End If
        End With
    End If
    Exit Sub
DropObj_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.DropObj", Erl)
End Sub

Sub EraseObj(ByVal num As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
    On Error GoTo EraseObj_Err
    Dim Rango As Byte
    MapData(Map, x, y).ObjInfo.amount = MapData(Map, x, y).ObjInfo.amount - num
    If MapData(Map, x, y).ObjInfo.amount <= 0 Then
        MapData(Map, x, y).ObjInfo.ObjIndex = 0
        MapData(Map, x, y).ObjInfo.amount = 0
        MapData(Map, x, y).ObjInfo.ElementalTags = 0
        Call modSendData.SendToAreaByPos(Map, x, y, PrepareMessageObjectDelete(x, y))
    End If
    Exit Sub
EraseObj_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.EraseObj", Erl)
End Sub

Sub MakeObj(ByRef obj As t_Obj, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal Limpiar As Boolean = True)
    On Error GoTo MakeObj_Err
    Dim Color As Long
    Dim Rango As Byte
    If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then
        If MapData(Map, x, y).ObjInfo.ObjIndex = obj.ObjIndex And MapData(Map, x, y).ObjInfo.ElementalTags = obj.ElementalTags Then
            MapData(Map, x, y).ObjInfo.amount = MapData(Map, x, y).ObjInfo.amount + obj.amount
        Else
            MapData(Map, x, y).ObjInfo.ObjIndex = obj.ObjIndex
            MapData(Map, x, y).ObjInfo.ElementalTags = obj.ElementalTags
            If ObjData(obj.ObjIndex).VidaUtil <> 0 Then
                MapData(Map, x, y).ObjInfo.amount = ObjData(obj.ObjIndex).VidaUtil
            Else
                MapData(Map, x, y).ObjInfo.amount = obj.amount
            End If
        End If
        Call modSendData.SendToAreaByPos(Map, x, y, PrepareMessageObjectCreate(obj.ObjIndex, MapData(Map, x, y).ObjInfo.amount, x, y, MapData(Map, x, y).ObjInfo.ElementalTags))
    End If
    Exit Sub
MakeObj_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.MakeObj", Erl)
End Sub

Function GetSlotForItemInInventory(ByVal UserIndex As Integer, ByRef MyObject As t_Obj) As Integer
    On Error GoTo GetSlotForItemInInventory_Err
    GetSlotForItemInInventory = -1
    Dim i As Integer
    For i = 1 To UserList(UserIndex).CurrentInventorySlots
        If UserList(UserIndex).invent.Object(i).ObjIndex = 0 And GetSlotForItemInInventory = -1 Then
            GetSlotForItemInInventory = i 'we found a valid place but keep looking in case we can stack
        ElseIf UserList(UserIndex).invent.Object(i).ObjIndex = MyObject.ObjIndex And UserList(UserIndex).invent.Object(i).ElementalTags = MyObject.ElementalTags And UserList( _
                UserIndex).invent.Object(i).amount + MyObject.amount <= GetMaxInvOBJ() Then
            GetSlotForItemInInventory = i 'we can stack the item, let use this slot
            Exit Function
        End If
    Next i
    Exit Function
GetSlotForItemInInventory_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.GetSlotForItemInInventory", Erl)
End Function

Function GetSlotInInventory(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Integer
    On Error GoTo GetSlotInInventory_Err
    GetSlotInInventory = -1
    Dim i As Integer
    For i = 1 To UserList(UserIndex).CurrentInventorySlots
        If UserList(UserIndex).invent.Object(i).ObjIndex = ObjIndex Then
            GetSlotInInventory = i
            Exit Function
        End If
    Next i
    Exit Function
GetSlotInInventory_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.GetSlotInInventory", Erl)
End Function

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As t_Obj) As Boolean
    On Error GoTo MeterItemEnInventario_Err
    Dim x    As Integer
    Dim y    As Integer
    Dim Slot As Integer
    If MiObj.ObjIndex = 12 Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiObj.amount
        MeterItemEnInventario = True
        Call WriteUpdateGold(UserIndex)
        Exit Function
    End If
    '¿el user ya tiene un objeto del mismo tipo? ?????
    Slot = GetSlotForItemInInventory(UserIndex, MiObj)
    If Slot <= 0 Then
        Call WriteLocaleMsg(UserIndex, MsgInventoryIsFull, e_FontTypeNames.FONTTYPE_FIGHT)
        MeterItemEnInventario = False
        Exit Function
    End If
    If UserList(UserIndex).invent.Object(Slot).ObjIndex = 0 Then
        UserList(UserIndex).invent.NroItems = UserList(UserIndex).invent.NroItems + 1
    End If
    'Mete el objeto
    If UserList(UserIndex).invent.Object(Slot).amount + MiObj.amount <= GetMaxInvOBJ() Then
        'Menor que MAX_INV_OBJS
        UserList(UserIndex).invent.Object(Slot).ObjIndex = MiObj.ObjIndex
        UserList(UserIndex).invent.Object(Slot).amount = UserList(UserIndex).invent.Object(Slot).amount + MiObj.amount
        UserList(UserIndex).invent.Object(Slot).ElementalTags = MiObj.ElementalTags
    Else
        UserList(UserIndex).invent.Object(Slot).amount = GetMaxInvOBJ()
    End If
    Call UpdateUserInv(False, UserIndex, Slot)
    MeterItemEnInventario = True
    UserList(UserIndex).flags.ModificoInventario = True
    Exit Function
MeterItemEnInventario_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.MeterItemEnInventario", Erl)
End Function

Function HayLugarEnInventario(ByVal UserIndex As Integer, ByVal TargetItemIndex As Integer, ByVal ItemCount) As Boolean
    On Error GoTo HayLugarEnInventario_err
    Dim x    As Integer
    Dim y    As Integer
    Dim Slot As Byte
    Slot = 1
    Do Until UserList(UserIndex).invent.Object(Slot).ObjIndex = 0 Or (UserList(UserIndex).invent.Object(Slot).ObjIndex = TargetItemIndex And UserList(UserIndex).invent.Object( _
            Slot).amount + ItemCount < 10000)
        Slot = Slot + 1
        If Slot > UserList(UserIndex).CurrentInventorySlots Then
            HayLugarEnInventario = False
            Exit Function
        End If
    Loop
    HayLugarEnInventario = True
    Exit Function
HayLugarEnInventario_err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.HayLugarEnInventario", Erl)
End Function

Sub PickObj(ByVal UserIndex As Integer)
    On Error GoTo PickObj_Err
    Dim x     As Integer
    Dim y     As Integer
    Dim Slot  As Byte
    Dim obj   As t_ObjData
    Dim MiObj As t_Obj
    '¿Hay algun obj?
    If IsInMapCarcelRestrictedArea(UserList(UserIndex).pos) Then
        Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(MSG_CANNOT_DROP_ITEMS_IN_JAIL, vbNullString, e_FontTypeNames.FONTTYPE_INFO))
        Exit Sub
    End If
    If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).ObjInfo.ObjIndex > 0 Then
        '¿Esta permitido agarrar este obj?
        If ObjData(MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).ObjInfo.ObjIndex).Agarrable <> 1 Then
            If UserList(UserIndex).flags.Montado = 1 Then
                ' Msg672=Debes descender de tu montura para agarrar objetos del suelo.
                Call WriteLocaleMsg(UserIndex, 672, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If Not UserCanPickUpItem(UserIndex) Then
                Exit Sub
            End If
            x = UserList(UserIndex).pos.x
            y = UserList(UserIndex).pos.y
            If UserList(UserIndex).flags.jugando_captura = 1 Then
                If Not InstanciaCaptura Is Nothing Then
                    If Not InstanciaCaptura.tomaBandera(UserIndex, MapData(UserList(UserIndex).pos.Map, x, y).ObjInfo.ObjIndex) Then
                        Exit Sub
                    End If
                End If
            End If
            obj = ObjData(MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).ObjInfo.ObjIndex)
            MiObj.amount = MapData(UserList(UserIndex).pos.Map, x, y).ObjInfo.amount
            MiObj.ObjIndex = MapData(UserList(UserIndex).pos.Map, x, y).ObjInfo.ObjIndex
            MiObj.ElementalTags = MapData(UserList(UserIndex).pos.Map, x, y).ObjInfo.ElementalTags
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                'Call WriteConsoleMsg(UserIndex, "No puedo cargar mas objetos.", e_FontTypeNames.FONTTYPE_INFO)
            Else
                'Quitamos el objeto
                Call EraseObj(MapData(UserList(UserIndex).pos.Map, x, y).ObjInfo.amount, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y)
                If Not UserList(UserIndex).flags.Privilegios And e_PlayerType.User Then Call LogGM(UserList(UserIndex).name, "Agarro:" & MiObj.amount & " Objeto:" & ObjData( _
                        MiObj.ObjIndex).name)
                'Si el obj es oro (12), se muestra la cantidad que agarro arriba del personaje
                If MiObj.ObjIndex = 12 Then
                    Call WriteTextOverTile(UserIndex, "+" & PonerPuntos(MiObj.amount), UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, RGB(212, 175, 55))
                Else
                    Call WriteShowPickUpObj(UserIndex, MiObj.ObjIndex, MiObj.amount)
                End If
                Call UserDidPickupItem(UserIndex, MiObj.ObjIndex)
                If UserList(UserIndex).flags.jugando_captura = 1 Then
                    If Not InstanciaCaptura Is Nothing Then
                        Call InstanciaCaptura.quitarBandera(UserIndex, MiObj.ObjIndex)
                    End If
                End If
                If BusquedaTesoroActiva Then
                    If UserList(UserIndex).pos.Map = TesoroNumMapa And UserList(UserIndex).pos.x = TesoroX And UserList(UserIndex).pos.y = TesoroY Then
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1639, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_TALK)) 'Msg1640=Eventos> ¬1 encontró el tesoro ¡Felicitaciones!
                        BusquedaTesoroActiva = False
                    End If
                End If
                If BusquedaRegaloActiva Then
                    If UserList(UserIndex).pos.Map = RegaloNumMapa And UserList(UserIndex).pos.x = RegaloX And UserList(UserIndex).pos.y = RegaloY Then
                        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1640, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_TALK)) 'Msg1640=Eventos> ¬1 fue el valiente que encontró el gran ítem mágico ¡Felicitaciones!
                        BusquedaRegaloActiva = False
                    End If
                End If
            End If
        End If
    Else
        If Not UserList(UserIndex).flags.UltimoMensaje = MSG_PICKUP_UNAVAILABLE Then
            Call WriteLocaleMsg(UserIndex, MSG_PICKUP_UNAVAILABLE, e_FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = MSG_PICKUP_UNAVAILABLE
        End If
    End If
    Exit Sub
PickObj_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.PickObj", Erl)
End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte, Optional ByVal bSkin As Boolean = False, Optional ByVal eSkinType As e_OBJType)

Dim obj                         As t_ObjData

    On Error GoTo Desequipar_Err

    With UserList(UserIndex)

        If Not bSkin Then
            'Desequipa el item slot del inventario
            If (Slot < LBound(.invent.Object)) Or (Slot > UBound(.invent.Object)) Then
                Exit Sub
            ElseIf .invent.Object(Slot).ObjIndex = 0 Then
                Exit Sub
            End If

            obj = ObjData(.invent.Object(Slot).ObjIndex)

            Select Case obj.OBJType
                Case e_OBJType.otWeapon
                    .invent.Object(Slot).Equipped = 0
                    .invent.EquippedWeaponObjIndex = 0
                    .invent.EquippedWeaponSlot = 0
                    .Char.Arma_Aura = ""
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, 0, True, 1))
                    .Char.WeaponAnim = NingunArma
                    If .flags.Montado = 0 Then
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                    End If
                    If obj.MagicDamageBonus > 0 Then
                        Call WriteUpdateDM(UserIndex)
                    End If
                Case e_OBJType.otArrows
                    .invent.Object(Slot).Equipped = 0
                    .invent.EquippedMunitionObjIndex = 0
                    .invent.EquippedMunitionSlot = 0
                    ' Case e_OBJType.otAnillos
                    '    .Invent.Object(slot).Equipped = 0
                    '    .Invent.AnilloEqpObjIndex = 0
                    ' .Invent.AnilloEqpSlot = 0
                Case e_OBJType.otWorkingTools
                    If .flags.PescandoEspecial = False Then
                        .invent.Object(Slot).Equipped = 0
                        .invent.EquippedWorkingToolObjIndex = 0
                        .invent.EquippedWorkingToolSlot = 0
                        If .flags.UsandoMacro = True Then
                            Call WriteMacroTrabajoToggle(UserIndex, False)
                        End If
                        .Char.WeaponAnim = NingunArma
                        If .flags.Montado = 0 Then
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                        End If
                    End If

                Case e_OBJType.otAmulets
                    Select Case obj.EfectoMagico
                        Case e_MagicItemEffect.eModifyAttributes
                            If obj.QueAtributo <> 0 Then
                                .Stats.UserAtributos(obj.QueAtributo) = .Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
                                .Stats.UserAtributosBackUP(obj.QueAtributo) = .Stats.UserAtributosBackUP(obj.QueAtributo) - obj.CuantoAumento
                                ' .Stats.UserAtributos(obj.QueAtributo) = .Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
                                Call WriteFYA(UserIndex)
                            End If
                        Case e_MagicItemEffect.eModifySkills
                            If obj.Que_Skill <> 0 Then
                                .Stats.UserSkills(obj.Que_Skill) = .Stats.UserSkills(obj.Que_Skill) - obj.CuantoAumento
                            End If
                        Case e_MagicItemEffect.eRegenerateHealth
                            .flags.RegeneracionHP = 0
                        Case e_MagicItemEffect.eRegenerateMana
                            .flags.RegeneracionMana = 0
                        Case e_MagicItemEffect.eIncreaseDamageToNpc
                            .Stats.MaxHit = .Stats.MaxHit - obj.CuantoAumento
                            .Stats.MinHIT = .Stats.MinHIT - obj.CuantoAumento
                        Case e_MagicItemEffect.eInmunityToNpcMagic    'Orbe ignea
                            .flags.NoMagiaEfecto = 0
                        Case e_MagicItemEffect.eIncinerate
                            .flags.incinera = 0
                        Case e_MagicItemEffect.eParalize
                            .flags.Paraliza = 0
                        Case e_MagicItemEffect.eProtectedResources
                            If .flags.Muerto = 0 Then
                                .Char.CartAnim = NoCart
                                Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                            End If
                        Case e_MagicItemEffect.eProtectedInventory
                            .flags.PendienteDelSacrificio = 0
                        Case e_MagicItemEffect.ePreventMagicWords
                            .flags.NoPalabrasMagicas = 0
                        Case e_MagicItemEffect.ePreventInvisibleDetection
                            .flags.NoDetectable = 0
                        Case e_MagicItemEffect.eIncreaseLearningSkills
                            .flags.PendienteDelExperto = 0
                        Case e_MagicItemEffect.ePoison
                            .flags.Envenena = 0
                        Case e_MagicItemEffect.eRingOfShadows
                            .flags.AnilloOcultismo = 0
                        Case e_MagicItemEffect.eTalkToDead
                            Call UnsetMask(.flags.StatusMask, e_StatusMask.eTalkToDead)
                            ' Msg673=Dejas el mundo de los muertos, ya no podrás comunicarte con ellos.
                            Call WriteLocaleMsg(UserIndex, 673, e_FontTypeNames.FONTTYPE_WARNING)
                            Call SendData(SendTarget.ToPCDeadAreaButIndex, UserIndex, PrepareMessageCharacterRemove(4, .Char.charindex, False, True))
                    End Select
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, 0, True, 5))
                    .Char.Otra_Aura = 0
                    .invent.Object(Slot).Equipped = 0
                    .invent.EquippedAmuletAccesoryObjIndex = 0
                    .invent.EquippedAmuletAccesorySlot = 0
                Case e_OBJType.otArmor
                
                    If .Invent_Skins.SlotArmourEquipped > 0 Then
                        If .Invent_Skins.Object(.Invent_Skins.SlotArmourEquipped).ObjIndex > 0 Then
                            If ObjData(.Invent_Skins.Object(.Invent_Skins.SlotArmourEquipped).ObjIndex).RequiereObjeto > 0 Then
                                Call DesequiparSkin(UserIndex, .Invent_Skins.SlotArmourEquipped)
                            End If
                        End If
                    End If
                
                    .invent.Object(Slot).Equipped = 0
                    .invent.EquippedArmorObjIndex = 0
                    .invent.EquippedArmorSlot = 0
                    If .flags.Navegando = 0 Then
                        If .flags.Montado = 0 Then
                            Call SetNakedBody(UserList(UserIndex))
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                        End If
                    End If
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, 0, True, 2))
                    .Char.Body_Aura = 0
                    If obj.ResistenciaMagica > 0 Then
                        Call WriteUpdateRM(UserIndex)
                    End If
                    
                Case e_OBJType.otHelmet
                    .invent.Object(Slot).Equipped = 0
                    .invent.EquippedHelmetObjIndex = 0
                    .invent.EquippedHelmetSlot = 0
                    .Char.Head_Aura = 0
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, 0, True, 4))
                    .Char.CascoAnim = NingunCasco
                    .Char.head = .Char.originalhead
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, _
                                        .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                    If obj.ResistenciaMagica > 0 Then
                        Call WriteUpdateRM(UserIndex)
                    End If
                Case e_OBJType.otShield
                
                    If .Invent_Skins.SlotShieldEquipped > 0 Then
                        If .Invent_Skins.Object(.Invent_Skins.SlotShieldEquipped).ObjIndex > 0 Then
                            If ObjData(.Invent_Skins.Object(.Invent_Skins.SlotShieldEquipped).ObjIndex).RequiereObjeto > 0 Then
                                Call DesequiparSkin(UserIndex, .Invent_Skins.SlotShieldEquipped)
                            End If
                        End If
                    End If
                
                    .invent.Object(Slot).Equipped = 0
                    .invent.EquippedShieldObjIndex = 0
                    .invent.EquippedShieldSlot = 0
                    .Char.Escudo_Aura = 0
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, 0, True, 3))
                    .Char.ShieldAnim = NingunEscudo
                    If .flags.Montado = 0 Then
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                    End If
                    If obj.ResistenciaMagica > 0 Then
                        Call WriteUpdateRM(UserIndex)
                    End If
                Case e_OBJType.otAmulets
                    .invent.Object(Slot).Equipped = 0
                    .invent.EquippedAmuletAccesoryObjIndex = 0
                    .invent.EquippedAmuletAccesorySlot = 0
                    .Char.DM_Aura = 0
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, 0, True, 6))
                    Call WriteUpdateDM(UserIndex)
                    Call WriteUpdateRM(UserIndex)
                Case e_OBJType.otRingAccesory, e_OBJType.otMagicalInstrument
                    .invent.Object(Slot).Equipped = 0
                    .invent.EquippedRingAccesoryObjIndex = 0
                    .invent.EquippedRingAccesorySlot = 0
                    .Char.RM_Aura = 0
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, 0, True, 7))
                    Call WriteUpdateRM(UserIndex)
                    Call WriteUpdateDM(UserIndex)
                Case e_OBJType.otBackpack
                    .invent.Object(Slot).Equipped = 0
                    .invent.EquippedBackpackObjIndex = 0
                    .invent.EquippedBackpackSlot = 0
                    .Char.BackpackAnim = 0
                    If .flags.Navegando = 0 Then
                        If .flags.Montado = 0 Then
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                        End If
                    End If
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, 0, True, 2))
                    .Char.Body_Aura = 0
            End Select
            Call UpdateUserInv(False, UserIndex, Slot)
        Else
            Call DesequiparSkin(UserIndex, Slot)
        End If
    End With
    
    Exit Sub
Desequipar_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.Desequipar", Erl)
End Sub

Sub DesequiparSkin(ByVal UserIndex As Integer, ByVal Slot As Byte)

Dim obj                         As t_ObjData
Dim eSkinType                   As e_OBJType

    On Error GoTo DesequiparSkin_Error
    
    With UserList(UserIndex)
        
        If (Slot < LBound(.Invent_Skins.Object)) Or (Slot > UBound(.Invent_Skins.Object)) Then
            Exit Sub
        ElseIf .Invent_Skins.Object(Slot).ObjIndex = 0 Then
            Exit Sub
        End If
        
        eSkinType = ObjData(.Invent_Skins.Object(Slot).ObjIndex).OBJType
        obj = ObjData(.Invent_Skins.Object(Slot).ObjIndex)

        If .Invent_Skins.Object(Slot).Equipped Then
            .Invent_Skins.Object(Slot).Equipped = False
        End If
        
        Select Case eSkinType
            Case e_OBJType.otSkinsArmours
                .Invent_Skins.ObjIndexArmourEquipped = 0
                .Invent_Skins.SlotArmourEquipped = 0
                If .invent.EquippedArmorObjIndex > 0 Then
                    .Char.body = ObtenerRopaje(UserIndex, ObjData(.invent.EquippedArmorObjIndex))
                Else
                    If SkinRequireObject(UserIndex, Slot) Then
                        Call SetNakedBody(UserList(UserIndex))
                    End If
                End If
                
            Case e_OBJType.otSkinsSpells
               .Stats.UserSkinsHechizos(ObjData(.Invent_Skins.Object(Slot).ObjIndex).HechizoIndex) = 0
               .Invent_Skins.Object(Slot).Equipped = False

            Case e_OBJType.otSkinsHelmets
                If ObjData(.Invent_Skins.ObjIndexHelmetEquipped).Subtipo = 2 Then
                    .Char.head = .OrigChar.head
                End If
                .Invent_Skins.ObjIndexHelmetEquipped = 0
                .Invent_Skins.SlotHelmetEquipped = 0
                
                If .invent.EquippedHelmetObjIndex > 0 Then
                    .Char.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
                Else
                    If SkinRequireObject(UserIndex, Slot) Then
                        .Char.CascoAnim = NingunCasco
                    End If
                End If
                
            Case e_OBJType.otSkinsWings
                If .Invent_Skins.ObjIndexBackpackEquipped > 0 Then
                    .Char.BackpackAnim = NoBackPack
                End If
                'Ojo acá!
                .Invent_Skins.ObjIndexBackpackEquipped = 0
                .Invent_Skins.SlotWindsEquipped = 0
                
            Case e_OBJType.otSkinsBoats
                    .Invent_Skins.ObjIndexBoatEquipped = 0
                    .Invent_Skins.SlotBoatEquipped = 0
                    Call EquiparBarco(UserIndex)
    
            Case e_OBJType.otSkinsShields
                .Invent_Skins.ObjIndexShieldEquipped = 0
                .Invent_Skins.SlotShieldEquipped = 0
                If .invent.EquippedShieldObjIndex > 0 Then
                    .Char.ShieldAnim = ObjData(.invent.EquippedShieldObjIndex).ShieldAnim
                Else
                    If SkinRequireObject(UserIndex, Slot) Then
                        .Char.ShieldAnim = NingunEscudo
                    End If
                End If

            Case e_OBJType.otSkinsWeapons
                .Invent_Skins.ObjIndexWeaponEquipped = 0
                .Invent_Skins.SlotWeaponEquipped = 0
                If .invent.EquippedWeaponObjIndex > 0 Then
                    .Char.WeaponAnim = ObjData(.invent.EquippedWeaponObjIndex).WeaponAnim
                Else
                    If SkinRequireObject(UserIndex, Slot) Then
                        .Char.WeaponAnim = NingunArma
                    End If
                End If
        End Select
        
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
        Call WriteChangeSkinSlot(UserIndex, eSkinType, Slot)
    End With
    
    On Error GoTo 0
    Exit Sub
DesequiparSkin_Error:
    Call Logging.TraceError(Err.Number, Err.Description, "InvUsuario.DesequiparSkin of Módulo", Erl())
End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    On Error GoTo ErrHandler
    If EsGM(UserIndex) Then
        SexoPuedeUsarItem = True
        Exit Function
    End If
    If ObjData(ObjIndex).Mujer = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).genero <> e_Genero.Hombre
    ElseIf ObjData(ObjIndex).Hombre = 1 Then
        SexoPuedeUsarItem = UserList(UserIndex).genero <> e_Genero.Mujer
    Else
        SexoPuedeUsarItem = True
    End If
    Exit Function
ErrHandler:
    Call LogError("SexoPuedeUsarItem")
End Function

Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    On Error GoTo FaccionPuedeUsarItem_Err
    If EsGM(UserIndex) Then
        FaccionPuedeUsarItem = True
        Exit Function
    End If
    If ObjIndex < 1 Then Exit Function
    If ObjData(ObjIndex).Real = 1 Then
        If ObjData(ObjIndex).LeadersOnly Then
            FaccionPuedeUsarItem = (Status(UserIndex) = e_Facciones.consejo)
        ElseIf Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
            FaccionPuedeUsarItem = esArmada(UserIndex)
        Else
            FaccionPuedeUsarItem = False
        End If
    ElseIf ObjData(ObjIndex).Caos = 1 Then
        If ObjData(ObjIndex).LeadersOnly Then
            FaccionPuedeUsarItem = (Status(UserIndex) = e_Facciones.concilio)
        ElseIf Status(UserIndex) = e_Facciones.Caos Or Status(UserIndex) = e_Facciones.concilio Then
            FaccionPuedeUsarItem = esCaos(UserIndex)
        Else
            FaccionPuedeUsarItem = False
        End If
    Else
        FaccionPuedeUsarItem = True
    End If
    Exit Function
FaccionPuedeUsarItem_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.FaccionPuedeUsarItem", Erl)
End Function

Function JerarquiaPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
    With UserList(UserIndex)
        If .Faccion.RecompensasCaos >= ObjData(ObjIndex).Jerarquia Then
            JerarquiaPuedeUsarItem = True
            Exit Function
        End If
        If .Faccion.RecompensasReal >= ObjData(ObjIndex).Jerarquia Then
            JerarquiaPuedeUsarItem = True
            Exit Function
        End If
    End With
End Function

'Equipa barco y hace el cambio de ropaje correspondiente
Sub EquiparBarco(ByVal UserIndex As Integer)
    On Error GoTo EquiparBarco_Err
    Dim Barco As t_ObjData
    With UserList(UserIndex)
        If .invent.EquippedShipObjIndex <= 0 Or .invent.EquippedShipObjIndex > UBound(ObjData) Then Exit Sub
        Barco = ObjData(.invent.EquippedShipObjIndex)
        If .flags.Muerto = 1 Then
            If Barco.Ropaje = iTraje Or Barco.Ropaje = iTrajeAltoNw Or Barco.Ropaje = iTrajeBajoNw Then
                ' No tenemos la cabeza copada que va con iRopaBuceoMuerto,
                ' asique asignamos el casper directamente caminando sobre el agua.
                .Char.body = iCuerpoMuerto 'iRopaBuceoMuerto
                .Char.head = iCabezaMuerto
            ElseIf Barco.Ropaje = iTrajeAltoNw Then
            ElseIf Barco.Ropaje = iTrajeBajoNw Then
            Else
                .Char.body = iFragataFantasmal
                .Char.head = 0
            End If
        Else ' Esta vivo
            If Barco.Ropaje = iTraje Then
                .Char.body = iTraje
                .Char.head = .OrigChar.head
                If .invent.EquippedHelmetObjIndex > 0 Then
                    .Char.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
                End If
            ElseIf Barco.Ropaje = iTrajeAltoNw Then
                .Char.body = iTrajeAltoNw
                .Char.head = .OrigChar.head
                If .invent.EquippedHelmetObjIndex > 0 Then
                    .Char.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
                End If
            ElseIf Barco.Ropaje = iTrajeBajoNw Then
                .Char.body = iTrajeBajoNw
                .Char.head = .OrigChar.head
                If .invent.EquippedHelmetObjIndex > 0 Then
                    .Char.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
                End If
            Else
                .Char.head = 0
                .Char.CascoAnim = NingunCasco
            End If
            If .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Then
                If Barco.Ropaje = iBarca Then .Char.body = iBarcaArmada
                If Barco.Ropaje = iGalera Then .Char.body = iGaleraArmada
                If Barco.Ropaje = iGaleon Then .Char.body = iGaleonArmada
            ElseIf .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then
                If Barco.Ropaje = iBarca Then .Char.body = iBarcaCaos
                If Barco.Ropaje = iGalera Then .Char.body = iGaleraCaos
                If Barco.Ropaje = iGaleon Then .Char.body = iGaleonCaos
            Else
                If Barco.Ropaje = iBarca Then .Char.body = IIf(.Faccion.Status = 0, iBarcaCrimi, iBarcaCiuda)
                If Barco.Ropaje = iGalera Then .Char.body = IIf(.Faccion.Status = 0, iGaleraCrimi, iGaleraCiuda)
                If Barco.Ropaje = iGaleon Then .Char.body = IIf(.Faccion.Status = 0, iGaleonCrimi, iGaleonCiuda)
            End If
        End If
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        
        If .Invent_Skins.ObjIndexBoatEquipped > 0 Then
            Call SkinEquip(UserIndex, .Invent_Skins.SlotBoatEquipped, .Invent_Skins.ObjIndexBoatEquipped)
        End If
        
        Call WriteNavigateToggle(UserIndex, .flags.Navegando)
        Call WriteNadarToggle(UserIndex, (Barco.Ropaje = iTraje Or Barco.Ropaje = iTrajeAltoNw Or Barco.Ropaje = iTrajeBajoNw), (Barco.Ropaje = iTrajeAltoNw Or Barco.Ropaje = iTrajeBajoNw))
        Call ActualizarVelocidadDeUsuario(UserIndex)
    End With
    Exit Sub
EquiparBarco_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.EquiparBarco", Erl)
End Sub

'Equipa un item del inventario
Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, Optional ByVal UserIsLoggingIn As Boolean = False, Optional ByVal bSkin As Boolean = False, Optional ByVal eSkinType As e_OBJType)

Dim bEquipSkin                  As Boolean
Dim obj                         As t_ObjData
Dim ObjIndex                    As Integer
Dim errordesc                   As String
Dim Ropaje                      As Integer

    On Error GoTo ErrHandler
    With UserList(UserIndex)

        If .flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Not bSkin Then
            ObjIndex = .invent.Object(Slot).ObjIndex
            obj = ObjData(ObjIndex)

            If PuedeUsarObjeto(UserIndex, ObjIndex, True) > 0 Then
                Exit Sub
            End If

            Select Case obj.OBJType
                Case e_OBJType.otWeapon
                    errordesc = "Arma"
                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eWeapon) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    'Si esta equipado lo quita
                    If .invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
                        Call Desequipar(UserIndex, Slot)
                        'Animacion por defecto
                        .Char.WeaponAnim = NingunArma
                        If .flags.Montado = 0 Then
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                        End If
                        Exit Sub
                    End If
                    'Quitamos el elemento anterior
                    If .invent.EquippedWeaponObjIndex > 0 Then
                        Call Desequipar(UserIndex, .invent.EquippedWeaponSlot)
                    End If
                    If .invent.EquippedWorkingToolObjIndex > 0 Then
                        Call Desequipar(UserIndex, .invent.EquippedWorkingToolSlot)
                    End If
                    .invent.Object(Slot).Equipped = 1
                    .invent.EquippedWeaponObjIndex = .invent.Object(Slot).ObjIndex
                    .invent.EquippedWeaponSlot = Slot
                    Call ValidateEquippedArrow(UserIndex)
                    If obj.DosManos = 1 Then
                        If .invent.EquippedShieldObjIndex > 0 Then
                            Call Desequipar(UserIndex, .invent.EquippedShieldSlot)
                            ' Msg674=No puedes usar armas dos manos si tienes un escudo equipado. Tu escudo fue desequipado.
                            Call WriteLocaleMsg(UserIndex, 674, e_FontTypeNames.FONTTYPE_INFOIAO)
                        End If
                    End If
                    'Sonido
                    If obj.SndAura = 0 Then
                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .pos.x, .pos.y))
                    Else
                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .pos.x, .pos.y))
                    End If
                    If Len(obj.CreaGRH) <> 0 Then
                        .Char.Arma_Aura = obj.CreaGRH
                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Arma_Aura, False, 1))
                    End If
                    If obj.MagicDamageBonus > 0 Then
                        Call WriteUpdateDM(UserIndex)
                    End If
                    If .flags.Montado = 0 Then
                        If .flags.Navegando = 0 Then
                            '¿Tiene una skin equipada?
                            If .Invent_Skins.ObjIndexWeaponEquipped > 0 Then
                                '¿Esa skin está limitada a un ítem específico y el item que estoy equipado, NO es esa?
                                If ObjData(.Invent_Skins.ObjIndexWeaponEquipped).RequiereObjeto > 0 And obj.ObjNum <> ObjData(.Invent_Skins.ObjIndexWeaponEquipped).RequiereObjeto Then
                                    .Char.WeaponAnim = obj.WeaponAnim
                                Else
                                    .Char.WeaponAnim = ObjData(.Invent_Skins.ObjIndexWeaponEquipped).WeaponAnim
                                End If
                            Else
                                .Char.WeaponAnim = obj.WeaponAnim
                            End If

                            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                        End If
                    End If

            Case e_OBJType.otBackpack
                errordesc = "Backpack"
                If .invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    .Char.BackpackAnim = NoBackPack
                    If .flags.Montado = 0 And .flags.Navegando = 0 Then
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, _
                                            .Char.BackpackAnim)
                    End If
                    Exit Sub
                End If
                If .invent.EquippedBackpackObjIndex > 0 Then
                    Call Desequipar(UserIndex, .invent.EquippedBackpackSlot)
                End If
                Ropaje = ObtenerRopaje(UserIndex, obj)
                If Ropaje = 0 Then
                    ' Msg676=Hay un error con este objeto. Infórmale a un administrador.
                    Call WriteLocaleMsg(UserIndex, 676, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'Lo equipa
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.Body_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Body_Aura, False, 2))
                End If
                .invent.Object(Slot).Equipped = 1
                .invent.EquippedBackpackObjIndex = .invent.Object(Slot).ObjIndex
                .invent.EquippedBackpackSlot = Slot
                If .flags.Montado = 0 And .flags.Navegando = 0 Then
                    .Char.BackpackAnim = Ropaje
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, _
                                        .Char.BackpackAnim)
                End If

            Case e_OBJType.otWorkingTools
                If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eTool) Then
                    Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'Si esta equipado lo quita
                If .invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                'Quitamos el elemento anterior
                If .invent.EquippedWorkingToolObjIndex > 0 Then
                    Call Desequipar(UserIndex, .invent.EquippedWorkingToolSlot)
                End If
                If .invent.EquippedWeaponObjIndex > 0 Then
                    Call Desequipar(UserIndex, .invent.EquippedWeaponSlot)
                End If
                .invent.Object(Slot).Equipped = 1
                .invent.EquippedWorkingToolObjIndex = ObjIndex
                .invent.EquippedWorkingToolSlot = Slot
                If .flags.Montado = 0 Then
                    If .flags.Navegando = 0 Then
                        .Char.WeaponAnim = obj.WeaponAnim
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                    End If
                End If
            Case e_OBJType.otAmulets
                errordesc = "Magico"
                If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eMagicItem) Then
                    Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'Si esta equipado lo quita
                If .invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                'Quitamos el elemento anterior
                If .invent.EquippedAmuletAccesoryObjIndex > 0 Then
                    Call Desequipar(UserIndex, .invent.EquippedAmuletAccesorySlot)
                End If
                .invent.Object(Slot).Equipped = 1
                .invent.EquippedAmuletAccesoryObjIndex = .invent.Object(Slot).ObjIndex
                .invent.EquippedAmuletAccesorySlot = Slot
                Select Case obj.EfectoMagico
                    Case e_MagicItemEffect.eModifyAttributes    'Modif la fuerza, agilidad, carisma, etc
                        .Stats.UserAtributosBackUP(obj.QueAtributo) = .Stats.UserAtributosBackUP(obj.QueAtributo) + obj.CuantoAumento
                        .Stats.UserAtributos(obj.QueAtributo) = MinimoInt(.Stats.UserAtributos(obj.QueAtributo) + obj.CuantoAumento, .Stats.UserAtributosBackUP(obj.QueAtributo) * 2)
                        Call WriteFYA(UserIndex)
                    Case e_MagicItemEffect.eModifySkills
                        .Stats.UserSkills(obj.Que_Skill) = .Stats.UserSkills(obj.Que_Skill) + obj.CuantoAumento
                    Case e_MagicItemEffect.eRegenerateHealth
                        .flags.RegeneracionHP = 1
                    Case e_MagicItemEffect.eRegenerateMana
                        .flags.RegeneracionMana = 1
                    Case e_MagicItemEffect.eIncreaseDamageToNpc
                        .Stats.MaxHit = .Stats.MaxHit + obj.CuantoAumento
                        .Stats.MinHIT = .Stats.MinHIT + obj.CuantoAumento
                    Case e_MagicItemEffect.eInmunityToNpcMagic
                        .flags.NoMagiaEfecto = 1
                    Case e_MagicItemEffect.eIncinerate
                        .flags.incinera = 1
                    Case e_MagicItemEffect.eParalize
                        .flags.Paraliza = 1
                    Case e_MagicItemEffect.eProtectedResources
                        If .flags.Navegando = 0 And .flags.Montado = 0 Then
                            .Char.CartAnim = obj.Ropaje
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                        End If
                    Case e_MagicItemEffect.eProtectedInventory
                        .flags.PendienteDelSacrificio = 1
                    Case e_MagicItemEffect.ePreventMagicWords
                        .flags.NoPalabrasMagicas = 1
                    Case e_MagicItemEffect.ePreventInvisibleDetection
                        .flags.NoDetectable = 1
                    Case e_MagicItemEffect.eIncreaseLearningSkills
                        .flags.PendienteDelExperto = 1
                    Case e_MagicItemEffect.ePoison
                        .flags.Envenena = 1
                    Case e_MagicItemEffect.eRingOfShadows
                        .flags.AnilloOcultismo = 1
                    Case e_MagicItemEffect.eTalkToDead
                        Call SetMask(.flags.StatusMask, e_StatusMask.eTalkToDead)
                        ' Msg675=Entras al mundo de los muertos, ahora podrás comunicarte con ellos.
                        Call WriteLocaleMsg(UserIndex, "675", e_FontTypeNames.FONTTYPE_WARNING)
                        Call CheckUpdateNeededUser(UserIndex, USER_NUEVO, True, 1)
                End Select
                'Sonido
                If obj.SndAura <> 0 Then
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .pos.x, .pos.y))
                End If
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.Otra_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Otra_Aura, False, 5))
                End If
            Case e_OBJType.otArrows
                If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eAmunition) Then
                    Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'Si esta equipado lo quita
                If .invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                Call EquipArrow(UserIndex, Slot)
            Case e_OBJType.otArmor
                If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eArmor) Then
                    Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                '¿Tiene una skin equipada?
                If .Invent_Skins.ObjIndexArmourEquipped > 0 Then
                    '¿Esa skin está limitada a una armadura específica y la armadura que estoy equipado, NO es esa?
                    If ObjData(.Invent_Skins.ObjIndexArmourEquipped).RequiereObjeto > 0 And obj.ObjNum <> ObjData(.Invent_Skins.ObjIndexArmourEquipped).RequiereObjeto Then
                        Ropaje = ObtenerRopaje(UserIndex, obj)
                    Else
                        Ropaje = ObtenerRopaje(UserIndex, ObjData(.Invent_Skins.ObjIndexArmourEquipped))
                    End If
                Else
                    Ropaje = ObtenerRopaje(UserIndex, obj)
                End If
                If Ropaje = 0 Then
                    ' Msg676=Hay un error con este objeto. Infórmale a un administrador.
                    Call WriteLocaleMsg(UserIndex, "676", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'Si esta equipado lo quita
                If .invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    If .flags.Navegando = 0 And .flags.Montado = 0 Then
                        Call SetNakedBody(UserList(UserIndex))
                        If Not UserIsLoggingIn Then
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                        End If
                    Else
                        .flags.Desnudo = 1
                    End If
                    Exit Sub
                End If
                'Quita el anterior
                If .invent.EquippedArmorObjIndex > 0 Then
                    errordesc = "Armadura 2"
                    Call Desequipar(UserIndex, .invent.EquippedArmorSlot)
                    Call ActualizarVelocidadDeUsuario(UserIndex)
                    errordesc = "Armadura 3"
                End If
                'Si esta equipando armadura faccionaria fuera de zona segura o fuera de trigger seguro
                If Not UserIsLoggingIn Then
                    If obj.Real > 0 Or obj.Caos > 0 Then
                        If Not MapData(.pos.Map, .pos.x, .pos.y).trigger = e_Trigger.ZonaSegura And Not MapInfo(.pos.Map).Seguro = 1 Then
                            Call WriteLocaleMsg(UserIndex, 2091, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                    End If
                End If
                'Lo equipa
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.Body_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Body_Aura, False, 2))
                End If
                .invent.Object(Slot).Equipped = 1
                .invent.EquippedArmorObjIndex = .invent.Object(Slot).ObjIndex
                .invent.EquippedArmorSlot = Slot
                Call ActualizarVelocidadDeUsuario(UserIndex)
                If .flags.Montado = 0 And .flags.Navegando = 0 Then
                    .Char.body = Ropaje
                    If Not UserIsLoggingIn Then    'Evitamos redundancia de paquetes durante el loggin.
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                    End If
                End If
                .flags.Desnudo = 0
                If obj.ResistenciaMagica > 0 Then
                    Call WriteUpdateRM(UserIndex)
                End If
            Case e_OBJType.otHelmet
                If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eHelm) Then
                    Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'Si esta equipado lo quita
                If .invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    .Char.CascoAnim = NingunCasco
                    If obj.Subtipo = 2 Then
                        .Char.head = .Char.originalhead
                    End If
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                    Exit Sub
                End If
                'Quita el anterior
                If .invent.EquippedHelmetObjIndex > 0 Then
                    Call Desequipar(UserIndex, .invent.EquippedHelmetSlot)
                End If
                errordesc = "Casco"
                'Lo equipa
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.Head_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Head_Aura, False, 4))
                End If
                .invent.Object(Slot).Equipped = 1
                .invent.EquippedHelmetObjIndex = .invent.Object(Slot).ObjIndex
                .invent.EquippedHelmetSlot = Slot
                If .flags.Navegando = 0 Then
                    If obj.Subtipo = 2 Then
                        ' Si el casco cambia la cabeza entera
                        If Not UserIsLoggingIn Then
                            .Char.originalhead = .Char.head
                        End If

                        '¿Tiene una skin equipada?
                        If .Invent_Skins.ObjIndexHelmetEquipped > 0 Then
                            '¿Esa skin está limitada a un ítem específico y el item que estoy equipado, NO es ese?
                            If ObjData(.Invent_Skins.ObjIndexHelmetEquipped).RequiereObjeto > 0 And obj.ObjNum <> ObjData(.Invent_Skins.ObjIndexHelmetEquipped).RequiereObjeto Then
                                .Char.CascoAnim = obj.CascoAnim
                            Else
                                .Char.CascoAnim = ObjData(.Invent_Skins.ObjIndexHelmetEquipped).CascoAnim
                            End If
                        Else
                            .Char.CascoAnim = obj.CascoAnim
                            .Char.head = obj.CascoAnim
                        End If
                        
                    Else
                        ' Si el casco se superpone (no reemplaza la cabeza)
                        If .Char.head >= PATREON_HEAD Then
                        'lñkñkñlk
                        Else
                            If .Invent_Skins.ObjIndexHelmetEquipped > 0 Then
                               .Char.CascoAnim = ObjData(.Invent_Skins.ObjIndexHelmetEquipped).CascoAnim
                            Else
                                .Char.CascoAnim = obj.CascoAnim
                            End If
                        End If
                    End If
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                End If
                If obj.ResistenciaMagica > 0 Then
                    Call WriteUpdateRM(UserIndex)
                End If
            Case e_OBJType.otShield
                If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eShiled) Then
                    Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'Si esta equipado lo quita
                If .invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    .Char.ShieldAnim = NingunEscudo
                    If .flags.Montado = 0 And .flags.Navegando = 0 Then
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                    End If
                    Exit Sub
                End If
                'Quita el anterior
                If .invent.EquippedShieldObjIndex > 0 Then
                    Call Desequipar(UserIndex, .invent.EquippedShieldSlot)
                End If
                If .invent.EquippedWeaponObjIndex > 0 Then
                    If ObjData(.invent.EquippedWeaponObjIndex).DosManos = 1 Then
                        Call Desequipar(UserIndex, .invent.EquippedWeaponSlot)
                        ' Msg677=No puedes equipar un escudo si tienes un arma dos manos equipada. Tu arma fue desequipada.
                        Call WriteLocaleMsg(UserIndex, 677, e_FontTypeNames.FONTTYPE_INFOIAO)
                    End If
                End If
                errordesc = "Escudo"
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.Escudo_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Escudo_Aura, False, 3))
                End If
                .invent.Object(Slot).Equipped = 1
                .invent.EquippedShieldObjIndex = .invent.Object(Slot).ObjIndex
                .invent.EquippedShieldSlot = Slot
                If .flags.Navegando = 0 And .flags.Montado = 0 Then
                    '¿Tiene una skin equipada?
                    If .Invent_Skins.ObjIndexShieldEquipped > 0 Then
                        '¿Esa skin está limitada a un ítem específico y el item que estoy equipado, NO es ese?
                        If ObjData(.Invent_Skins.ObjIndexShieldEquipped).RequiereObjeto > 0 And obj.ObjNum <> ObjData(.Invent_Skins.ObjIndexShieldEquipped).RequiereObjeto Then
                            .Char.ShieldAnim = obj.ShieldAnim
                        Else
                            .Char.ShieldAnim = ObjData(.Invent_Skins.ObjIndexShieldEquipped).ShieldAnim
                        End If
                    Else
                        .Char.ShieldAnim = obj.ShieldAnim
                    End If
                End If
                If obj.ResistenciaMagica > 0 Then
                    Call WriteUpdateRM(UserIndex)
                End If
                
                If Not UserIsLoggingIn Then
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                End If
                
            Case e_OBJType.otMagicalInstrument, e_OBJType.otRingAccesory
                'Si esta equipado lo quita
                If .invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                'Quita el anterior
                If .invent.EquippedRingAccesorySlot > 0 Then
                    Call Desequipar(UserIndex, .invent.EquippedRingAccesorySlot)
                End If
                .invent.Object(Slot).Equipped = 1
                If ObjData(.invent.Object(Slot).ObjIndex).OBJType = e_OBJType.otRingAccesory Then
                    .invent.EquippedRingAccesoryObjIndex = .invent.Object(Slot).ObjIndex
                    .invent.EquippedRingAccesorySlot = Slot
                    Call WriteUpdateRM(UserIndex)
                ElseIf ObjData(.invent.Object(Slot).ObjIndex).OBJType = e_OBJType.otMagicalInstrument Then
                    .invent.EquippedRingAccesoryObjIndex = .invent.Object(Slot).ObjIndex
                    .invent.EquippedRingAccesorySlot = Slot
                    Call WriteUpdateDM(UserIndex)
                End If
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.DM_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.DM_Aura, False, 6))
                End If
            Case e_OBJType.otDonator
                If obj.Subtipo = 4 Then
                    Call EquipAura(Slot, .invent, UserIndex)
                End If
        End Select
    Else                       'Si es un SKIN:
        ObjIndex = .Invent_Skins.Object(Slot).ObjIndex
        obj = ObjData(ObjIndex)
        Select Case obj.OBJType
            Case e_OBJType.otSkinsArmours, e_OBJType.otSkinsSpells, e_OBJType.otSkinsWeapons, e_OBJType.otSkinsShields, e_OBJType.otSkinsHelmets, e_OBJType.otSkinsBoats, e_OBJType.otSkinsWings
                'Si esta equipado lo quita
                If .Invent_Skins.Object(Slot).Equipped And Not UserIsLoggingIn Then
                    'Sonido
                    'Feat para implementar más adelante.
                    'tmpSoundItem = ObjData(.Invent_Skins.Object(Slot).ObjIndex).Snd2
                    'If tmpSoundItem > 0 Then
                    '    Call SendData(SendTarget.ToPCAreaWithSound, UserIndex, PrepareMessagePlayWave(tmpSoundItem, .pos.x, .pos.y))
                    'End If
                    Call Desequipar(UserIndex, Slot, True, ObjData(ObjIndex).OBJType)
                    Exit Sub   'Revisar este EXIT SUB
                End If

                'Feat para implementar más adelante.
                'tmpSoundItem = ObjData(.Invent_Skins.Object(Slot).ObjIndex).Snd1
                'If tmpSoundItem > 0 Then
                '    Call SendData(SendTarget.ToPCAreaWithSound, UserIndex, PrepareMessagePlayWave(tmpSoundItem, .pos.x, .pos.y))
                'End If
                If CanEquipSkin(UserIndex, Slot, True) Then
                    Call SkinEquip(UserIndex, Slot, ObjIndex)
                End If
        End Select
    End If
End With
'Actualiza
Call UpdateUserInv(False, UserIndex, Slot)
Exit Sub
ErrHandler:
Debug.Print errordesc
Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.Description & "- " & errordesc)
End Sub

Public Sub EquipAura(ByVal Slot As Integer, ByRef inventory As t_Inventario, ByVal UserIndex As Integer)
    If inventory.Object(Slot).Equipped Then
        inventory.Object(Slot).Equipped = False
        Exit Sub
    End If
    If Slot < 1 Or Slot > UBound(inventory.Object) Then Exit Sub
    Dim Index As Integer
    Dim obj   As t_ObjData
    For Index = 1 To UBound(inventory.Object)
        If Index <> Slot And inventory.Object(Index).Equipped Then
            If inventory.Object(Index).ObjIndex > 0 Then
                If inventory.Object(Index).ObjIndex > 0 Then
                    obj = ObjData(inventory.Object(Index).ObjIndex)
                    If obj.OBJType = otDonator And obj.Subtipo = 4 Then
                        inventory.Object(Index).Equipped = 0
                        Call UpdateUserInv(False, UserIndex, Index)
                    End If
                End If
            End If
        End If
    Next Index
    inventory.Object(Slot).Equipped = 1
End Sub

Public Function CheckClaseTipo(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
    On Error GoTo ErrHandler
    If EsGM(UserIndex) Then
        CheckClaseTipo = True
        Exit Function
    End If
    Select Case ObjData(ItemIndex).ClaseTipo
        Case 0
            CheckClaseTipo = True
            Exit Function
        Case 2
            If UserList(UserIndex).clase = e_Class.Mage Then CheckClaseTipo = True
            If UserList(UserIndex).clase = e_Class.Druid Then CheckClaseTipo = True
            Exit Function
        Case 1
            If UserList(UserIndex).clase = e_Class.Warrior Then CheckClaseTipo = True
            If UserList(UserIndex).clase = e_Class.Assasin Then CheckClaseTipo = True
            If UserList(UserIndex).clase = e_Class.Bard Then CheckClaseTipo = True
            If UserList(UserIndex).clase = e_Class.Cleric Then CheckClaseTipo = True
            If UserList(UserIndex).clase = e_Class.Paladin Then CheckClaseTipo = True
            If UserList(UserIndex).clase = e_Class.Trabajador Then CheckClaseTipo = True
            If UserList(UserIndex).clase = e_Class.Hunter Then CheckClaseTipo = True
            Exit Function
    End Select
    Exit Function
ErrHandler:
    Call LogError("Error CheckClaseTipo ItemIndex:" & ItemIndex)
End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal ByClick As Byte)
    Dim ObjIndex As Integer
    Dim nowRaw   As Long
    Dim TargObj  As t_ObjData
    Dim obj      As t_ObjData
    Dim MiObj    As t_Obj
    On Error GoTo hErr
    ' Agrego el Cuerno de la Armada y la Legión.
    'Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
    With UserList(UserIndex)
        If .invent.Object(Slot).amount = 0 Then Exit Sub
        If Not CanUseItem(.flags, .Counters) Then
            Call WriteLocaleMsg(UserIndex, 395, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If PuedeUsarObjeto(UserIndex, .invent.Object(Slot).ObjIndex, True) > 0 Then
            Exit Sub
        End If

        obj = ObjData(.invent.Object(Slot).ObjIndex)
        nowRaw = GetTickCountRaw()
        Dim TimeSinceLastUse As Double: TimeSinceLastUse = TicksElapsed(.CdTimes(obj.cdType), nowRaw)
        If TimeSinceLastUse < obj.Cooldown Then Exit Sub
        If IsSet(obj.ObjFlags, e_ObjFlags.e_UseOnSafeAreaOnly) Then
            If MapInfo(.pos.Map).Seguro = 0 Then
                ' Msg678=Solo podes usar este objeto en mapas seguros.
                Call WriteLocaleMsg(UserIndex, 678, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If

        If obj.OBJType = e_OBJType.otWeapon Then
            If obj.Proyectil = 1 Then
                'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
                If ByClick <> 0 Then
                    If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
                Else
                    If Not IntervaloPermiteUsarClick(UserIndex) Then Exit Sub
                End If
            Else
                'dagas
                If ByClick <> 0 Then
                    If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
                Else
                    If Not IntervaloPermiteUsarClick(UserIndex) Then Exit Sub
                End If
            End If
        Else
            If ByClick <> 0 Then
                If Not IntervaloPermiteUsarClick(UserIndex) Then Exit Sub
            Else
                If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
            End If
        End If

        If .flags.Meditando Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
        End If

        If obj.Newbie = 1 And Not EsNewbie(UserIndex) And Not EsGM(UserIndex) Then
            ' Msg679=Solo los newbies pueden usar estos objetos.
            Call WriteLocaleMsg(UserIndex, 679, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .Stats.ELV < obj.MinELV Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1926, obj.MinELV, e_FontTypeNames.FONTTYPE_INFO))    ' Msg1926=Necesitas ser nivel ¬1 para usar este item.
            Exit Sub
        End If
        If .Stats.ELV > obj.MaxLEV And obj.MaxLEV > 0 Then
            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1982, obj.MaxLEV, e_FontTypeNames.FONTTYPE_INFO))    ' Msg1982=Este objeto no puede ser utilizado por personajes de nivel ¬1 o superior.
            Exit Sub
        End If
        ObjIndex = .invent.Object(Slot).ObjIndex
        .flags.TargetObjInvIndex = ObjIndex
        .flags.TargetObjInvSlot = Slot

        Select Case obj.OBJType
            Case e_OBJType.otSkinsArmours, e_OBJType.otSkinsSpells, e_OBJType.otSkinsBoats, e_OBJType.otSkinsHelmets, e_OBJType.otSkinsShields, e_OBJType.otSkinsWeapons, e_OBJType.otSkinsWings

                If .invent.Object(Slot).ObjIndex = 0 Then Exit Sub
                If ClasePuedeUsarItem(UserIndex, .invent.Object(Slot).ObjIndex) And SexoPuedeUsarItem(UserIndex, .invent.Object(Slot).ObjIndex) And FaccionPuedeUsarItem(UserIndex, .invent.Object(Slot).ObjIndex) And LevelCanUseItem(UserIndex, ObjData(.invent.Object(Slot).ObjIndex)) Then
                    If Not HaveThisSkin(UserIndex, .invent.Object(Slot).ObjIndex) Then
                        If AddSkin(UserIndex, .invent.Object(Slot).ObjIndex) Then
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Call UpdateSingleItemInv(UserIndex, Slot, False)
                        End If
                    Else
                        Call WriteLocaleMsg(UserIndex, 2101, e_FontTypeNames.FONTTYPE_INFO) 'Msg2101=Ya tienes este skin.
                    End If
                Else
                    Call WriteLocaleMsg(UserIndex, 2101, e_FontTypeNames.FONTTYPE_INFO) 'Msg2101=Tu clase o nivel no te permite usar este skin.
                End If

            Case e_OBJType.otUseOnce
                If .flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'Usa el item
                .Stats.MinHam = .Stats.MinHam + obj.MinHam
                If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
                Call WriteUpdateHungerAndThirst(UserIndex)
                If obj.JineteLevel > 0 Then
                    If .Stats.JineteLevel < obj.JineteLevel Then
                        .Stats.JineteLevel = obj.JineteLevel
                    Else
                        'Msg2080 = No puedes consumir un nivel de jinete menor al que posees actualmente
                        Call WriteLocaleMsg(UserIndex, 2079, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                'Sonido
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.SOUND_COMIDA, .pos.x, .pos.y))
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)
                .flags.ModificoInventario = True
            Case e_OBJType.otGoldCoin
                If .flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                .Stats.GLD = .Stats.GLD + .invent.Object(Slot).amount
                .invent.Object(Slot).amount = 0
                .invent.Object(Slot).ObjIndex = 0
                .invent.NroItems = .invent.NroItems - 1
                .flags.ModificoInventario = True
                Call UpdateUserInv(False, UserIndex, Slot)
                Call WriteUpdateGold(UserIndex)
            Case e_OBJType.otWeapon
                If .flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
                    Exit Sub
                End If
                If Not .Stats.MinSta > 0 Then
                    'Msg2129=¡No tengo energía!
                    Call SendData(SendTarget.ToIndex, UserIndex, PrepareLocalizedChatOverHead(2129, UserList(UserIndex).Char.charindex, vbWhite))
                    'Msg93=Estás muy cansado
                    Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If ObjData(ObjIndex).Proyectil = 1 Then
                    If IsSet(.flags.StatusMask, e_StatusMask.eTransformed) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantUseBowTransformed, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    Call WriteWorkRequestTarget(UserIndex, Proyectiles)
                Else
                    If .flags.TargetObj = Wood Then
                        If .invent.Object(Slot).ObjIndex = DAGA Then
                            Call TratarDeHacerFogata(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY, UserIndex)
                        End If
                    End If
                End If
                If .invent.Object(Slot).Equipped = 0 Then
                    Exit Sub
                End If
            Case e_OBJType.otWorkingTools
                If .flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
                    Exit Sub
                End If
                If Not .Stats.MinSta > 0 Then
                    'Msg2129=¡No tengo energía!
                    Call SendData(SendTarget.ToIndex, UserIndex, PrepareLocalizedChatOverHead(2129, UserList(UserIndex).Char.charindex, vbWhite))
                    'Msg93=Estás muy cansado
                    Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
                If .invent.Object(Slot).Equipped = 0 Then
                    Call WriteLocaleMsg(UserIndex, 376, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Select Case obj.Subtipo
                    Case 1, 2  ' Herramientas del Pescador - Caña y Red
                        Call WriteWorkRequestTarget(UserIndex, e_Skill.Pescar)
                    Case 3     ' Herramientas de Alquimia - Tijeras
                        Call WriteWorkRequestTarget(UserIndex, e_Skill.Alquimia)
                    Case 4     ' Herramientas de Alquimia - Olla
                        Call EnivarObjConstruiblesAlquimia(UserIndex)
                        Call WriteShowAlquimiaForm(UserIndex)
                    Case 5     ' Herramientas de Carpinteria - Serrucho
                        Call EnivarObjConstruibles(UserIndex)
                        Call WriteShowCarpenterForm(UserIndex)
                    Case 6     ' Herramientas de Tala - Hacha
                        Call WriteWorkRequestTarget(UserIndex, e_Skill.Talar)
                    Case 7     ' Herramientas de Herrero - Martillo
                        ' Msg680=Debes hacer click derecho sobre el yunque.
                        Call WriteLocaleMsg(UserIndex, 680, e_FontTypeNames.FONTTYPE_INFOIAO)
                    Case 8     ' Herramientas de Mineria - Piquete
                        Call WriteWorkRequestTarget(UserIndex, e_Skill.Mineria)
                    Case 9     ' Herramientas de Sastreria - Costurero
                        Call EnivarObjConstruiblesSastre(UserIndex)
                        Call WriteShowSastreForm(UserIndex)
                End Select
            Case e_OBJType.otPotions
                If .flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
                    Exit Sub
                End If
                If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
                    ' Msg681=¡¡Debes esperar unos momentos para tomar otra poción!!
                    Call WriteLocaleMsg(UserIndex, 681, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                .flags.TomoPocion = True
                .flags.TipoPocion = obj.TipoPocion
                Dim CabezaFinal   As Integer
                Dim CabezaActual  As Integer
                Select Case .flags.TipoPocion
                    Case e_PotionType.ModifiesAgility    'Modif la agilidad
                        .flags.DuracionEfecto = obj.DuracionEfecto
                        'Usa el item
                        .Stats.UserAtributos(e_Atributos.Agilidad) = MinimoInt(.Stats.UserAtributos(e_Atributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador), .Stats.UserAtributosBackUP(e_Atributos.Agilidad) * 2)
                        Call WriteFYA(UserIndex)
                        ' Consumir pocion solo si el usuario no esta en zona de uso libre
                        If Not IsConsumableFreeZone(UserIndex) Then
                            ' Quitamos el ítem del inventario
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                        End If
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                        Else
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
                        End If
                    Case e_PotionType.ModifiesStrength    'Modif la fuerza
                        .flags.DuracionEfecto = obj.DuracionEfecto
                        'Usa el item
                        .Stats.UserAtributos(e_Atributos.Fuerza) = MinimoInt(.Stats.UserAtributos(e_Atributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador), .Stats.UserAtributosBackUP(e_Atributos.Fuerza) * 2)
                        ' Consumir pocion solo si el usuario no esta en zona de uso libre
                        If Not IsConsumableFreeZone(UserIndex) Then
                            ' Quitamos el ítem del inventario
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                        End If
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                        Else
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
                        End If
                        Call WriteFYA(UserIndex)
                    Case e_PotionType.ModifiesHp     'Poción roja, restaura HP
                        ' Usa el ítem
                        If .flags.DivineBlood > 0 Then
                            Call WriteLocaleMsg(UserIndex, 2096, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        Dim HealingAmount As Long
                        Dim Source        As Integer
                        ' Calcula la cantidad de curación
                        HealingAmount = RandomNumber(obj.MinModificador, obj.MaxModificador) * UserMod.GetSelfHealingBonus(UserList(UserIndex))
                        ' Modifica la salud del jugador
                        Call UserMod.ModifyHealth(UserIndex, HealingAmount)
                        ' Consumir pocion solo si el usuario no esta en zona de uso libre
                        If Not IsConsumableFreeZone(UserIndex) Then
                            ' Quitamos el ítem del inventario
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                        End If
                        ' Reproduce sonido al usar la poción
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                        Else
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
                        End If
                    Case e_PotionType.ModifiesMp   'Poción azul, restaura MANA
                        Dim porcentajeRec As Byte
                        porcentajeRec = obj.Porcentaje
                        ' Usa el ítem: restaura el MANA
                        .Stats.MinMAN = IIf(.Stats.MinMAN > 20000, 20000, .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, porcentajeRec))
                        If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
                        ' Consumir pocion solo si el usuario no esta en zona de uso libre
                        If Not IsConsumableFreeZone(UserIndex) Then
                            ' Quitamos el ítem del inventario
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                        End If
                        ' Reproduce sonido al usar la poción
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                        Else
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
                        End If
                    Case e_PotionType.HealsPoison     ' Pocion violeta
                        If .flags.Envenenado > 0 Then
                            .flags.Envenenado = 0
                            ' Msg682=Te has curado del envenenamiento.
                            Call WriteLocaleMsg(UserIndex, 682, e_FontTypeNames.FONTTYPE_INFO)
                            'Quitamos del inv el item
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                            If obj.Snd1 <> 0 Then
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Else
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
                            End If
                        Else
                            ' Msg683=¡No te encuentras envenenado!
                            Call WriteLocaleMsg(UserIndex, 683, e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    Case e_PotionType.HealsParalysis     ' Remueve Parálisis
                        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
                            If .flags.Paralizado = 1 Then
                                .flags.Paralizado = 0
                                Call WriteParalizeOK(UserIndex)
                            End If
                            If .flags.Inmovilizado = 1 Then
                                .Counters.Inmovilizado = 0
                                .flags.Inmovilizado = 0
                                Call WriteInmovilizaOK(UserIndex)
                            End If
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                            If obj.Snd1 <> 0 Then
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Else
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(255, .pos.x, .pos.y))
                            End If
                            ' Msg684=Te has removido la paralizis.
                            Call WriteLocaleMsg(UserIndex, 684, e_FontTypeNames.FONTTYPE_INFOIAO)
                        Else
                            ' Msg685=No estas paralizado.
                            Call WriteLocaleMsg(UserIndex, 685, e_FontTypeNames.FONTTYPE_INFOIAO)
                        End If
                    Case e_PotionType.ModifiesStamina     ' Pocion Naranja
                        .Stats.MinSta = .Stats.MinSta + RandomNumber(obj.MinModificador, obj.MaxModificador)
                        If .Stats.MinSta > .Stats.MaxSta Then .Stats.MinSta = .Stats.MaxSta
                        'Quitamos del inv el item
                        If Not IsConsumableFreeZone(UserIndex) Then
                            ' Quitamos el ítem del inventario
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                        End If
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                        Else
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
                        End If
                    Case e_PotionType.ModifiesHeadRandom    ' Pocion cambio cara
                        Select Case .genero
                            Case e_Genero.Hombre
                                Select Case .raza
                                    Case e_Raza.Humano
                                        CabezaFinal = RandomNumber(1, 40)
                                    Case e_Raza.Elfo
                                        CabezaFinal = RandomNumber(101, 132)
                                    Case e_Raza.Drow
                                        CabezaFinal = RandomNumber(201, 229)
                                    Case e_Raza.Enano
                                        CabezaFinal = RandomNumber(301, 329)
                                    Case e_Raza.Gnomo
                                        CabezaFinal = RandomNumber(401, 429)
                                    Case e_Raza.Orco
                                        CabezaFinal = RandomNumber(501, 529)
                                End Select
                            Case e_Genero.Mujer
                                Select Case .raza
                                    Case e_Raza.Humano
                                        CabezaFinal = RandomNumber(50, 80)
                                    Case e_Raza.Elfo
                                        CabezaFinal = RandomNumber(150, 179)
                                    Case e_Raza.Drow
                                        CabezaFinal = RandomNumber(250, 279)
                                    Case e_Raza.Gnomo
                                        CabezaFinal = RandomNumber(350, 379)
                                    Case e_Raza.Enano
                                        CabezaFinal = RandomNumber(450, 479)
                                    Case e_Raza.Orco
                                        CabezaFinal = RandomNumber(550, 579)
                                End Select
                        End Select
                        .Char.head = CabezaFinal
                        .OrigChar.head = CabezaFinal
                        .OrigChar.originalhead = CabezaFinal    'cabeza final
                        Call ChangeUserChar(UserIndex, .Char.body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                        'Quitamos del inv el item
                        .Counters.timeFx = 3
                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 102, 0, .pos.x, .pos.y))
                        If CabezaActual <> CabezaFinal Then
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                        Else
                            ' Msg686=¡Rayos! Te tocó la misma cabeza, item no consumido. Tienes otra oportunidad.
                            Call WriteLocaleMsg(UserIndex, 686, e_FontTypeNames.FONTTYPE_INFOIAO)
                        End If
                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                    Case e_PotionType.ModifiesSex     ' Pocion sexo
                        Select Case .genero
                            Case e_Genero.Hombre
                                .genero = e_Genero.Mujer
                            Case e_Genero.Mujer
                                .genero = e_Genero.Hombre
                        End Select
                        Select Case .genero
                            Case e_Genero.Hombre
                                Select Case .raza
                                    Case e_Raza.Humano
                                        CabezaFinal = RandomNumber(1, 40)
                                    Case e_Raza.Elfo
                                        CabezaFinal = RandomNumber(101, 132)
                                    Case e_Raza.Drow
                                        CabezaFinal = RandomNumber(201, 229)
                                    Case e_Raza.Enano
                                        CabezaFinal = RandomNumber(301, 329)
                                    Case e_Raza.Gnomo
                                        CabezaFinal = RandomNumber(401, 429)
                                    Case e_Raza.Orco
                                        CabezaFinal = RandomNumber(501, 529)
                                End Select
                            Case e_Genero.Mujer
                                Select Case .raza
                                    Case e_Raza.Humano
                                        CabezaFinal = RandomNumber(50, 80)
                                    Case e_Raza.Elfo
                                        CabezaFinal = RandomNumber(150, 179)
                                    Case e_Raza.Drow
                                        CabezaFinal = RandomNumber(250, 279)
                                    Case e_Raza.Gnomo
                                        CabezaFinal = RandomNumber(350, 379)
                                    Case e_Raza.Enano
                                        CabezaFinal = RandomNumber(450, 479)
                                    Case e_Raza.Orco
                                        CabezaFinal = RandomNumber(550, 579)
                                End Select
                        End Select
                        .Char.head = CabezaFinal
                        .OrigChar.head = CabezaFinal
                        Call ChangeUserChar(UserIndex, .Char.body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                        'Quitamos del inv el item
                        .Counters.timeFx = 3
                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 102, 0, .pos.x, .pos.y))
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                        Else
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
                        End If
                    Case e_PotionType.TurnsYouInvisible    ' Invisibilidad
                        If .flags.invisible = 0 And .Counters.DisabledInvisibility = 0 Then
                            If IsSet(.flags.StatusMask, eTaunting) Then
                                ' Msg687=No tiene efecto.
                                Call WriteLocaleMsg(UserIndex, 687, e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                                Exit Sub
                            End If
                            .flags.invisible = 1
                            .Counters.Invisibilidad = obj.DuracionEfecto
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True, .pos.x, .pos.y))
                            Call WriteContadores(UserIndex)
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                            If obj.Snd1 <> 0 Then
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Else
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(123, .pos.x, .pos.y))
                            End If
                            ' Msg688=Te has escondido entre las sombras...
                            Call WriteLocaleMsg(UserIndex, 688, e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                        Else
                            ' Msg689=Ya estás invisible.
                            Call WriteLocaleMsg(UserIndex, 689, e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                            Exit Sub
                        End If
                        ' Poción que limpia todo
                    Case e_PotionType.HealsAllStatusEffects
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        .flags.Envenenado = 0
                        .flags.Incinerado = 0
                        If .flags.Inmovilizado = 1 Then
                            .Counters.Inmovilizado = 0
                            .flags.Inmovilizado = 0
                            Call WriteInmovilizaOK(UserIndex)
                        End If
                        If .flags.Paralizado = 1 Then
                            .flags.Paralizado = 0
                            Call WriteParalizeOK(UserIndex)
                        End If
                        If .flags.Ceguera = 1 Then
                            .flags.Ceguera = 0
                            Call WriteBlindNoMore(UserIndex)
                        End If
                        If .flags.Maldicion = 1 Then
                            .flags.Maldicion = 0
                            .Counters.Maldicion = 0
                        End If
                        .Stats.MinSta = .Stats.MaxSta
                        .Stats.MinAGU = .Stats.MaxAGU
                        .Stats.MinMAN = .Stats.MaxMAN
                        .Stats.MinHp = .Stats.MaxHp
                        .Stats.MinHam = .Stats.MaxHam
                        Call WriteUpdateHungerAndThirst(UserIndex)
                        ' Msg690=Donador> Te sentís sano y lleno.
                        Call WriteLocaleMsg(UserIndex, 690, e_FontTypeNames.FONTTYPE_WARNING)
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                        Else
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
                        End If
                        ' Poción runa
                    Case 14
                        If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
                            ' Msg691=No podés usar la runa estando en la cárcel.
                            Call WriteLocaleMsg(UserIndex, 691, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        Dim Map     As Integer
                        Dim x       As Byte
                        Dim y       As Byte
                        Dim DeDonde As t_WorldPos
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        Select Case .Hogar
                            Case e_Ciudad.cUllathorpe
                                DeDonde = Ullathorpe
                            Case e_Ciudad.cNix
                                DeDonde = Nix
                            Case e_Ciudad.cBanderbill
                                DeDonde = Banderbill
                            Case e_Ciudad.cLindos
                                DeDonde = Lindos
                            Case e_Ciudad.cArghal
                                DeDonde = Arghal
                            Case e_Ciudad.cArkhein
                                DeDonde = Arkhein
                            Case e_Ciudad.cForgat
                                DeDonde = Forgat
                            Case e_Ciudad.cEldoria
                                DeDonde = Eldoria
                            Case e_Ciudad.cPenthar
                                DeDonde = Penthar
                            Case Else
                                DeDonde = Ullathorpe
                        End Select
                        Map = DeDonde.Map
                        x = DeDonde.x
                        y = DeDonde.y
                        Call FindLegalPos(UserIndex, Map, x, y)
                        Call WarpUserChar(UserIndex, Map, x, y, True)
                        'Msg884= Ya estas a salvo...
                        Call WriteLocaleMsg(UserIndex, 884, e_FontTypeNames.FONTTYPE_WARNING)
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                        Else
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
                        End If
                    Case e_PotionType.ModifiesMarriage    ' Divorcio
                        If .flags.Casado = 1 Then
                            Dim tUser As t_UserReference
                            '.flags.Pareja
                            tUser = NameIndex(GetUserSpouse(.flags.SpouseId))
                            If Not IsValidUserRef(tUser) Then
                                'Msg885= Tu pareja deberás estar conectada para divorciarse.
                                Call WriteLocaleMsg(UserIndex, 885, e_FontTypeNames.FONTTYPE_INFOIAO)
                            Else
                                Call QuitarUserInvItem(UserIndex, Slot, 1)
                                UserList(tUser.ArrayIndex).flags.Casado = 0
                                UserList(tUser.ArrayIndex).flags.SpouseId = 0
                                .flags.Casado = 0
                                .flags.SpouseId = 0
                                'Msg886= Te has divorciado.
                                Call WriteLocaleMsg(UserIndex, 886, e_FontTypeNames.FONTTYPE_INFOIAO)
                                Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1983, .name, e_FontTypeNames.FONTTYPE_INFOIAO)) ' Msg1983=¬1 se ha divorciado de ti.
                                If obj.Snd1 <> 0 Then
                                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                                Else
                                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
                                End If
                            End If
                        Else
                            'Msg887= No estas casado.
                            Call WriteLocaleMsg(UserIndex, 887, e_FontTypeNames.FONTTYPE_INFOIAO)
                        End If
                    Case e_PotionType.ModifiesHeadRandomLegendary    'Cara legendaria
                        Select Case .genero
                            Case e_Genero.Hombre
                                Select Case .raza
                                    Case e_Raza.Humano
                                        CabezaFinal = RandomNumber(684, 686)
                                    Case e_Raza.Elfo
                                        CabezaFinal = RandomNumber(690, 692)
                                    Case e_Raza.Drow
                                        CabezaFinal = RandomNumber(696, 698)
                                    Case e_Raza.Enano
                                        CabezaFinal = RandomNumber(702, 704)
                                    Case e_Raza.Gnomo
                                        CabezaFinal = RandomNumber(708, 710)
                                    Case e_Raza.Orco
                                        CabezaFinal = RandomNumber(714, 716)
                                End Select
                            Case e_Genero.Mujer
                                Select Case .raza
                                    Case e_Raza.Humano
                                        CabezaFinal = RandomNumber(687, 689)
                                    Case e_Raza.Elfo
                                        CabezaFinal = RandomNumber(693, 695)
                                    Case e_Raza.Drow
                                        CabezaFinal = RandomNumber(699, 701)
                                    Case e_Raza.Gnomo
                                        CabezaFinal = RandomNumber(705, 707)
                                    Case e_Raza.Enano
                                        CabezaFinal = RandomNumber(711, 713)
                                    Case e_Raza.Orco
                                        CabezaFinal = RandomNumber(717, 719)
                                End Select
                        End Select
                        CabezaActual = .OrigChar.head
                        .Char.head = CabezaFinal
                        .OrigChar.head = CabezaFinal
                        Call ChangeUserChar(UserIndex, .Char.body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                        'Quitamos del inv el item
                        If CabezaActual <> CabezaFinal Then
                            .Counters.timeFx = 3
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 102, 0, .pos.x, .pos.y))
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                        Else
                            'Msg888= ¡Rayos! No pude asignarte una cabeza nueva, item no consumido. ¡Proba de nuevo!
                            Call WriteLocaleMsg(UserIndex, 888, e_FontTypeNames.FONTTYPE_INFOIAO)
                        End If
                    Case e_PotionType.ModifiesParticlesTemporary   ' tan solo crea una particula por determinado tiempo
                        Dim Particula           As Integer
                        Dim Tiempo              As Long
                        Dim ParticulaPermanente As Byte
                        Dim sobrechar           As Byte
                        If obj.CreaParticula <> "" Then
                            Particula = val(ReadField(1, obj.CreaParticula, Asc(":")))
                            Tiempo = val(ReadField(2, obj.CreaParticula, Asc(":")))
                            ParticulaPermanente = val(ReadField(3, obj.CreaParticula, Asc(":")))
                            sobrechar = val(ReadField(4, obj.CreaParticula, Asc(":")))
                            If ParticulaPermanente = 1 Then
                                .Char.ParticulaFx = Particula
                                .Char.loops = Tiempo
                            End If
                            If sobrechar = 1 Then
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFXToFloor(.pos.x, .pos.y, Particula, Tiempo))
                            Else
                                .Counters.timeFx = 3
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, Particula, Tiempo, False, , .pos.x, .pos.y))
                            End If
                        End If
                        If obj.CreaFX <> 0 Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageFxPiso(obj.CreaFX, .pos.x, .pos.y))
                        End If
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                        End If
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Case 21 'pocion negra
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        If obj.Snd1 <> 0 Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                        Else
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
                        End If
                        'Msg893= Te has suicidado.
                        Call WriteLocaleMsg(UserIndex, 893, e_FontTypeNames.FONTTYPE_EJECUCION)
                        Call CustomScenarios.UserDie(UserIndex)
                        Call UserMod.UserDie(UserIndex)
                    Case 23
                        If obj.ApplyEffectId > 0 Then
                            Call AddOrResetEffect(UserIndex, obj.ApplyEffectId)
                        End If
                        Call UpdateCd(UserIndex, ObjData(ObjIndex).cdType)
                        'Quitamos del inv el item
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        Call UpdateUserInv(False, UserIndex, Slot)
                        Exit Sub
                End Select
                If obj.ApplyEffectId > 0 Then
                    Call AddOrResetEffect(UserIndex, obj.ApplyEffectId)
                End If
                Call WriteUpdateUserStats(UserIndex)
                Call UpdateUserInv(False, UserIndex, Slot)
            Case e_OBJType.otDrinks
                If .flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
                If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
                Call WriteUpdateHungerAndThirst(UserIndex)
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                If obj.ApplyEffectId > 0 Then
                    Call AddOrResetEffect(UserIndex, obj.ApplyEffectId)
                End If
                If obj.Snd1 <> 0 Then
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                Else
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
                End If
                Call UpdateUserInv(False, UserIndex, Slot)
            Case e_OBJType.otChest
                If .flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
                    Exit Sub
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)
                Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1984, obj.name, e_FontTypeNames.FONTTYPE_New_DONADOR))    ' Msg1984=Has abierto un ¬1 y obtuviste...
                If obj.Snd1 <> 0 Then
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                End If
                If obj.CreaFX <> 0 Then
                    .Counters.timeFx = 3
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, obj.CreaFX, 0, .pos.x, .pos.y))
                End If
                Dim i As Byte
                Select Case obj.Subtipo
                    Case 1
                        For i = 1 To obj.CantItem
                            If Not MeterItemEnInventario(UserIndex, obj.Item(i)) Then
                                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Dios Or e_PlayerType.Admin)) Then
                                    Call TirarItemAlPiso(.pos, obj.Item(i))
                                End If
                            End If
                            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).name & " (" & obj.Item(i).amount & ")", e_FontTypeNames.FONTTYPE_INFOBOLD))
                        Next i
                    Case 2
                        For i = 1 To obj.CantEntrega
                            Dim indexobj As Byte
                            indexobj = RandomNumber(1, obj.CantItem)
                            Dim Index As t_Obj
                            Index.ObjIndex = obj.Item(indexobj).ObjIndex
                            Index.amount = obj.Item(indexobj).amount
                            If Not MeterItemEnInventario(UserIndex, Index) Then
                                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Dios Or e_PlayerType.Admin)) Then
                                    Call TirarItemAlPiso(.pos, Index)
                                End If
                            End If
                            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageConsoleMsg(ObjData(Index.ObjIndex).name & " (" & Index.amount & ")", e_FontTypeNames.FONTTYPE_INFOBOLD))
                        Next i
                    Case 3
                        For i = 1 To obj.CantItem
                            If RandomNumber(1, obj.Item(i).data) = 1 Then
                                If Not MeterItemEnInventario(UserIndex, obj.Item(i)) Then
                                    If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Dios Or e_PlayerType.Admin)) Then
                                        Call TirarItemAlPiso(.pos, obj.Item(i))
                                    End If
                                End If
                                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).name & " (" & obj.Item(i).amount & ")", e_FontTypeNames.FONTTYPE_INFOBOLD))
                            End If
                        Next i
                End Select
            Case e_OBJType.otKeys
                If .flags.Muerto = 1 Then
                    'Msg895= ¡¡Estas muerto!! Solo podes usar items cuando estas vivo.
                    Call WriteLocaleMsg(UserIndex, 895, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .flags.TargetObj = 0 Then Exit Sub
                TargObj = ObjData(.flags.TargetObj)
                '¿El objeto clickeado es una puerta?
                If TargObj.OBJType = e_OBJType.otDoors Then
                    If TargObj.clave < 1000 Then
                        'Msg896= Las llaves en el inventario están desactivadas. Sólo se permiten en el llavero.
                        Call WriteLocaleMsg(UserIndex, 896, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    '¿Esta cerrada?
                    If TargObj.Cerrada = 1 Then
                        '¿Cerrada con llave?
                        If TargObj.Llave > 0 Then
                            Dim ClaveLlave As Integer
                            If TargObj.clave = obj.clave Then
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, UserList( _
                                   UserIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, UserList( _
                                   UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
                                'Msg897= Has abierto la puerta.
                                Call WriteLocaleMsg(UserIndex, 897, e_FontTypeNames.FONTTYPE_INFO)
                                ClaveLlave = obj.clave
                                Call EliminarLlaves(ClaveLlave, UserIndex)
                                Exit Sub
                            Else
                                'Msg898= La llave no sirve.
                                Call WriteLocaleMsg(UserIndex, 898, e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        Else
                            If TargObj.clave = obj.clave Then
                                MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex = _
                                   ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, UserList( _
                                   UserIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                                'Msg899= Has cerrado con llave la puerta.
                                Call WriteLocaleMsg(UserIndex, 899, e_FontTypeNames.FONTTYPE_INFO)
                                .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                                Exit Sub
                            Else
                                'Msg900= La llave no sirve.
                                Call WriteLocaleMsg(UserIndex, 900, e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        End If
                    Else
                        'Msg901= No esta cerrada.
                        Call WriteLocaleMsg(UserIndex, 901, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
            Case e_OBJType.otEmptyBottle
                If .flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If Not InMapBounds(.flags.TargetMap, .flags.TargetX, .flags.TargetY) Then
                    Exit Sub
                End If
                If (MapData(.pos.Map, .flags.TargetX, .flags.TargetY).Blocked And FLAG_AGUA) = 0 Then
                    'Msg902= No hay agua allí.
                    Call WriteLocaleMsg(UserIndex, 902, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If Distance(.pos.x, .pos.y, .flags.TargetX, .flags.TargetY) > 2 Then
                    'Msg903= Debes acercarte más al agua.
                    Call WriteLocaleMsg(UserIndex, 903, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                MiObj.amount = 1
                MiObj.ObjIndex = ObjData(.invent.Object(Slot).ObjIndex).IndexAbierta
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.pos, MiObj)
                End If
                Call UpdateUserInv(False, UserIndex, Slot)
            Case e_OBJType.otFullBottle
                If .flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
                    Exit Sub
                End If
                .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
                If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
                Call WriteUpdateHungerAndThirst(UserIndex)
                MiObj.amount = 1
                MiObj.ObjIndex = ObjData(.invent.Object(Slot).ObjIndex).IndexCerrada
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    Call TirarItemAlPiso(.pos, MiObj)
                End If
                Call UpdateUserInv(False, UserIndex, Slot)
            Case e_OBJType.otParchment
                If .flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
                    Exit Sub
                End If
                If ClasePuedeUsarItem(UserIndex, .invent.Object(Slot).ObjIndex, Slot) And RazaPuedeUsarItem(UserIndex, .invent.Object(Slot).ObjIndex, Slot) Then
                    'If .Stats.MaxMAN > 0 Then
                    If .Stats.MinHam > 0 And .Stats.MinAGU > 0 Then
                        Call AgregarHechizo(UserIndex, Slot)
                        Call UpdateUserInv(False, UserIndex, Slot)
                    Else
                        'Msg904= Estas demasiado hambriento y sediento.
                        Call WriteLocaleMsg(UserIndex, 904, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    'Msg906= Por mas que lo intentas, no podés comprender el manuescrito.
                    Call WriteLocaleMsg(UserIndex, 906, e_FontTypeNames.FONTTYPE_INFO)
                End If
            Case e_OBJType.otMinerals
                If .flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
                    Exit Sub
                End If
                Call WriteWorkRequestTarget(UserIndex, FundirMetal)
            Case e_OBJType.otMusicalInstruments
                If .flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
                    Exit Sub
                End If
                If obj.Real Then    '¿Es el Cuerno Real?
                    If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                        If MapInfo(.pos.Map).Seguro = 1 Then
                            'Msg907= No hay Peligro aquí. Es Zona Segura
                            Call WriteLocaleMsg(UserIndex, 907, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        Call SendData(SendTarget.toMap, .pos.Map, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                        Exit Sub
                    Else
                        'Msg908= Solo Miembros de la Armada Real pueden usar este cuerno.
                        Call WriteLocaleMsg(UserIndex, 908, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                ElseIf obj.Caos Then    '¿Es el Cuerno Legión?
                    If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                        If MapInfo(.pos.Map).Seguro = 1 Then
                            'Msg909= No hay Peligro aquí. Es Zona Segura
                            Call WriteLocaleMsg(UserIndex, 909, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        Call SendData(SendTarget.toMap, .pos.Map, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                        Exit Sub
                    Else
                        'Msg910= Solo Miembros de la Legión Oscura pueden usar este cuerno.
                        Call WriteLocaleMsg(UserIndex, 910, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                'Si llega aca es porque es o Laud o Tambor o Flauta
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
            Case e_OBJType.otShips
                ' Piratas y trabajadores navegan al nivel 23
                If .invent.Object(Slot).ObjIndex <> iObjTrajeAltoNw And .invent.Object(Slot).ObjIndex <> iObjTrajeBajoNw And .invent.Object(Slot).ObjIndex <> iObjTraje Then
                    If .clase = e_Class.Trabajador Or .clase = e_Class.Pirat Then
                        If .Stats.ELV < 23 Then
                            'Msg911= Para recorrer los mares debes ser nivel 23 o superior.
                            Call WriteLocaleMsg(UserIndex, 911, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        ' Nivel mínimo 25 para navegar, si no sos pirata ni trabajador
                    ElseIf .Stats.ELV < 25 Then
                        'Msg912= Para recorrer los mares debes ser nivel 25 o superior.
                        Call WriteLocaleMsg(UserIndex, 912, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                ElseIf .invent.Object(Slot).ObjIndex = iObjTrajeAltoNw Or .invent.Object(Slot).ObjIndex = iObjTrajeBajoNw Then
                    If (.flags.Navegando = 0 Or (.invent.EquippedShipObjIndex <> iObjTrajeAltoNw And .invent.EquippedShipObjIndex <> iObjTrajeBajoNw)) And MapData(.pos.Map, _
                       .pos.x + 1, .pos.y).trigger <> e_Trigger.DETALLEAGUA And MapData(.pos.Map, .pos.x - 1, .pos.y).trigger <> e_Trigger.DETALLEAGUA And MapData(.pos.Map, _
                       .pos.x, .pos.y + 1).trigger <> e_Trigger.DETALLEAGUA And MapData(.pos.Map, .pos.x, .pos.y - 1).trigger <> e_Trigger.DETALLEAGUA Then
                        'Msg913= Este traje es para aguas contaminadas.
                        Call WriteLocaleMsg(UserIndex, 913, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                ElseIf .invent.Object(Slot).ObjIndex = iObjTraje Then
                    If (.flags.Navegando = 0 Or .invent.EquippedShipObjIndex <> iObjTraje) And MapData(.pos.Map, .pos.x + 1, .pos.y).trigger <> e_Trigger.NADOCOMBINADO And _
                       MapData(.pos.Map, .pos.x - 1, .pos.y).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.x, .pos.y + 1).trigger <> e_Trigger.NADOCOMBINADO _
                       And MapData(.pos.Map, .pos.x, .pos.y - 1).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.x + 1, .pos.y).trigger <> _
                       e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.x - 1, .pos.y).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.x, .pos.y + 1).trigger <> _
                       e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.x, .pos.y - 1).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.x + 1, .pos.y).trigger <> _
                       e_Trigger.NADOBAJOTECHO And MapData(.pos.Map, .pos.x - 1, .pos.y).trigger <> e_Trigger.NADOBAJOTECHO And MapData(.pos.Map, .pos.x, .pos.y + _
                       1).trigger <> e_Trigger.NADOBAJOTECHO And MapData(.pos.Map, .pos.x, .pos.y - 1).trigger <> e_Trigger.NADOBAJOTECHO Then
                        'Msg914= Este traje es para zonas poco profundas.
                        Call WriteLocaleMsg(UserIndex, 914, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                If .flags.Navegando = 0 Then
                    If LegalWalk(.pos.Map, .pos.x - 1, .pos.y, e_Heading.WEST, True, False) Or LegalWalk(.pos.Map, .pos.x, .pos.y - 1, e_Heading.NORTH, True, False) Or LegalWalk( _
                       .pos.Map, .pos.x + 1, .pos.y, e_Heading.EAST, True, False) Or LegalWalk(.pos.Map, .pos.x, .pos.y + 1, e_Heading.SOUTH, True, False) Then
                        Call DoNavega(UserIndex, obj, Slot)
                    Else
                        'Msg915= ¡Debes aproximarte al agua para usar el barco o traje de baño!
                        Call WriteLocaleMsg(UserIndex, 915, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    If .invent.EquippedShipObjIndex <> .invent.Object(Slot).ObjIndex Then
                        Call DoNavega(UserIndex, obj, Slot)
                    Else
                        If LegalWalk(.pos.Map, .pos.x - 1, .pos.y, e_Heading.WEST, False, True) Or LegalWalk(.pos.Map, .pos.x, .pos.y - 1, e_Heading.NORTH, False, True) Or _
                           LegalWalk(.pos.Map, .pos.x + 1, .pos.y, e_Heading.EAST, False, True) Or LegalWalk(.pos.Map, .pos.x, .pos.y + 1, e_Heading.SOUTH, False, True) Then
                            Call DoNavega(UserIndex, obj, Slot)
                        Else
                            'Msg916= ¡Debes aproximarte a la costa para dejar la barca!
                            Call WriteLocaleMsg(UserIndex, 916, e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                End If
            Case e_OBJType.otSaddles
                'Verifica todo lo que requiere la montura
                If .flags.Muerto = 1 Then
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
                    Exit Sub
                End If
                If .flags.Navegando = 1 Then
                    'Msg917= Debes dejar de navegar para poder cabalgar.
                    Call WriteLocaleMsg(UserIndex, 917, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If MapInfo(.pos.Map).zone = "DUNGEON" Then
                    'Msg918= No podes cabalgar dentro de un dungeon.
                    Call WriteLocaleMsg(UserIndex, 918, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call DoMontar(UserIndex, obj, Slot)
            Case e_OBJType.otDonator
                Select Case obj.Subtipo
                    Case 1
                        If .Counters.Pena <> 0 Then
                            ' Msg691=No podés usar la runa estando en la cárcel.
                            Call WriteLocaleMsg(UserIndex, 691, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
                            ' Msg691=No podés usar la runa estando en la cárcel.
                            Call WriteLocaleMsg(UserIndex, 691, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
                        'Msg919= Has viajado por el mundo.
                        Call WriteLocaleMsg(UserIndex, 919, e_FontTypeNames.FONTTYPE_WARNING)
                        Call QuitarUserInvItem(UserIndex, Slot, 1)
                        Call UpdateUserInv(False, UserIndex, Slot)
                    Case 2
                        Exit Sub
                    Case 3
                        Exit Sub
                End Select
            Case e_OBJType.otPassageTicket
                If .flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
                    Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .flags.TargetNpcTipo <> Pirata Then
                    'Msg920= Primero debes hacer click sobre el pirata.
                    Call WriteLocaleMsg(UserIndex, 920, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 3 Then
                    Call WriteLocaleMsg(UserIndex, 8, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .pos.Map <> obj.DesdeMap Then
                    Call WriteLocaleChatOverHead(UserIndex, 1354, "", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite) ' Msg1354=El pasaje no lo compraste aquí! Largate!
                    Exit Sub
                End If
                If Not MapaValido(obj.HastaMap) Then
                    Call WriteLocaleChatOverHead(UserIndex, 1355, "", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite) ' Msg1355=El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.
                    Exit Sub
                End If
                If obj.NecesitaNave > 0 Then
                    If .Stats.UserSkills(e_Skill.Navegacion) < 80 Then
                        Call WriteLocaleChatOverHead(UserIndex, 1356, "", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite) ' Msg1356=Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.
                        Exit Sub
                    End If
                End If
                Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
                'Msg921= Has viajado por varios días, te sientes exhausto!
                Call WriteLocaleMsg(UserIndex, 921, e_FontTypeNames.FONTTYPE_WARNING)
                .Stats.MinAGU = 0
                .Stats.MinHam = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)
            Case e_OBJType.otRecallStones
                If .Counters.Pena <> 0 Then
                    ' Msg691=No podés usar la runa estando en la cárcel.
                    Call WriteLocaleMsg(UserIndex, 691, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
                    ' Msg691=No podés usar la runa estando en la cárcel.
                    Call WriteLocaleMsg(UserIndex, 691, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If MapInfo(.pos.Map).Seguro = 0 And .flags.Muerto = 0 Then
                    ' Msg692=Solo podes usar tu runa en zonas seguras.
                    Call WriteLocaleMsg(UserIndex, 692, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If .Accion.AccionPendiente Then
                    Exit Sub
                End If
                Select Case ObjData(ObjIndex).TipoRuna
                    Case e_RuneType.ReturnHome
                        .Counters.TimerBarra = HomeTimer
                    Case e_RuneType.Escape
                        .Counters.TimerBarra = HomeTimer
                    Case e_RuneType.MesonSafePassage
                        .Counters.TimerBarra = 5
                End Select
                If Not EsGM(UserIndex) Then
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, e_GraphicEffects.Runa, 400, False))
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageBarFx(.Char.charindex, 350, e_AccionBarra.Runa))
                Else
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, e_GraphicEffects.Runa, 50, False))
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageBarFx(.Char.charindex, 100, e_AccionBarra.Runa))
                End If
                .Accion.Particula = e_GraphicEffects.Runa
                .Accion.AccionPendiente = True
                .Accion.TipoAccion = e_AccionBarra.Runa
                .Accion.RunaObj = ObjIndex
                .Accion.ObjSlot = Slot
            Case e_OBJType.otMap
                Call WriteShowFrmMapa(UserIndex)
            Case e_OBJType.OtQuest
                If obj.QuestId > 0 Then Call WriteObjQuestSend(UserIndex, obj.QuestId, Slot)
            Case e_OBJType.otAmulets
                Select Case ObjData(ObjIndex).Subtipo
                    Case e_MagicItemSubType.TargetUsable
                        Call WriteWorkRequestTarget(UserIndex, e_Skill.TargetableItem)
                End Select
                Select Case ObjData(ObjIndex).EfectoMagico
                    Case e_MagicItemEffect.eProtectedResources
                        If ObjData(ObjIndex).ApplyEffectId <= 0 Then
                            Exit Sub
                        End If
                        Call UpdateCd(UserIndex, ObjData(ObjIndex).cdType)
                        Call AddOrResetEffect(UserIndex, ObjData(ObjIndex).ApplyEffectId)
                End Select
            Case e_OBJType.otUsableOntarget
                .flags.UsingItemSlot = .flags.TargetObjInvSlot
                Call WriteWorkRequestTarget(UserIndex, e_Skill.TargetableItem)
        End Select
    End With
    Exit Sub
hErr:
    LogError "Error en useinvitem Usuario: " & UserList(UserIndex).name & " item:" & obj.name & " index: " & UserList(UserIndex).invent.Object(Slot).ObjIndex
End Sub

'**************************************************************
' Description: Determines whether the user is in a map zone
'              where potions are not consumed upon use.
'
' Parameters:  UserIndex      - Index of the user
'              triggerStatus  - Current trigger status of the user
'
' Returns:     Boolean - True if the user is in a potion-free zone
'                       False if the potion should be consumed
'**************************************************************
Public Function IsConsumableFreeZone(ByVal UserIndex As Integer) As Boolean
    Dim currentMap     As Integer
    Dim isTriggerZone  As Boolean
    Dim isTierUser     As Boolean
    Dim isHouseZone    As Boolean
    Dim isSpecialZone  As Boolean
    Dim isTrainingZone As Boolean
    Dim isArena        As Boolean
    Dim triggerStatus As e_Trigger6

    triggerStatus = TriggerZonaPelea(UserIndex, UserIndex)
    ' Obtener el mapa actual del usuario
    currentMap = UserList(UserIndex).pos.Map
    ' Verificar si está en zona con trigger activo
    isTriggerZone = (triggerStatus = e_Trigger6.TRIGGER6_PERMITE)
    ' Verificar si es un usuario con tier de suscripción
    isTierUser = (UserList(UserIndex).Stats.tipoUsuario = tAventurero Or UserList(UserIndex).Stats.tipoUsuario = tHeroe Or UserList(UserIndex).Stats.tipoUsuario = tLeyenda)
    ' Zona de casas/sotanos arenas: mapas del 600 al 749 con trigger activo
    isHouseZone = (currentMap >= 600 And currentMap <= 749 And isTriggerZone)
    ' Zonas especiales fijas donde no se consumen pociones
    ' 275, 276, 277 - Capture the Flag
    Select Case currentMap
        Case MAP_CAPTURE_THE_FLAG_1, MAP_CAPTURE_THE_FLAG_2, MAP_CAPTURE_THE_FLAG_3
            isSpecialZone = True
        Case Else
            isSpecialZone = False
    End Select
    ' 297 - Arena de Lindos
    isArena = (currentMap = MAP_ARENA_LINDOS And isTriggerZone)
    ' Meson Hostigado - Beneficio Patreon: mapa 172, con trigger activo y jugador con tier
    isTrainingZone = (currentMap = MAP_MESON_HOSTIGADO And isTriggerZone And isTierUser)
    ' Si esta en alguna de las zonas anteriores, no se consume la poción
    IsConsumableFreeZone = (isHouseZone Or isSpecialZone Or isTrainingZone Or isArena)
End Function

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)
    On Error GoTo EnivarArmasConstruibles_Err
    Call WriteBlacksmithWeapons(UserIndex)
    Exit Sub
EnivarArmasConstruibles_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarArmasConstruibles", Erl)
End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)
    On Error GoTo EnivarObjConstruibles_Err
    Call WriteCarpenterObjects(UserIndex)
    Exit Sub
EnivarObjConstruibles_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruibles", Erl)
End Sub

Sub SendCraftableElementRunes(ByVal UserIndex As Integer)
    On Error GoTo SendCraftableElementRunes_Err
    Call WriteBlacksmithElementalRunes(UserIndex)
    Exit Sub
SendCraftableElementRunes_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.SendCraftableElementRunes", Erl)
End Sub

Sub EnivarObjConstruiblesAlquimia(ByVal UserIndex As Integer)
    On Error GoTo EnivarObjConstruiblesAlquimia_Err
    Call WriteAlquimistaObjects(UserIndex)
    Exit Sub
EnivarObjConstruiblesAlquimia_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruiblesAlquimia", Erl)
End Sub

Sub EnivarObjConstruiblesSastre(ByVal UserIndex As Integer)
    On Error GoTo EnivarObjConstruiblesSastre_Err
    Call WriteSastreObjects(UserIndex)
    Exit Sub
EnivarObjConstruiblesSastre_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruiblesSastre", Erl)
End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)
    On Error GoTo EnivarArmadurasConstruibles_Err
    Call WriteBlacksmithArmors(UserIndex)
    Exit Sub
EnivarArmadurasConstruibles_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarArmadurasConstruibles", Erl)
End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean
    On Error GoTo ItemSeCae_Err
    ItemSeCae = (ObjData(Index).Real <> 1 Or ObjData(Index).NoSeCae = 0) And (ObjData(Index).Caos <> 1 Or ObjData(Index).NoSeCae = 0) And ObjData(Index).OBJType <> _
            e_OBJType.otKeys And ObjData(Index).OBJType <> e_OBJType.otShips And ObjData(Index).OBJType <> e_OBJType.otSaddles And ObjData(Index).NoSeCae = 0 And Not ObjData( _
            Index).Intirable = 1 And Not ObjData(Index).Destruye = 1 And Not ObjData(Index).Instransferible = 1
    Exit Function
ItemSeCae_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.ItemSeCae", Erl)
End Function

Public Function PirataCaeItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
    On Error GoTo PirataCaeItem_Err
    With UserList(UserIndex)
        If .clase = e_Class.Pirat And .Stats.ELV >= 37 And .flags.Navegando = 1 Then
            ' Si no está navegando, se caen los items
            If .invent.EquippedShipObjIndex > 0 Then
                ' Con galeón cada item tiene una probabilidad de caerse del 67%
                If ObjData(.invent.EquippedShipObjIndex).Ropaje = iGaleon Then
                    If RandomNumber(1, 100) <= 33 Then
                        Exit Function
                    End If
                End If
            End If
        End If
    End With
    PirataCaeItem = True
    Exit Function
PirataCaeItem_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.PirataCaeItem", Erl)
End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
    On Error GoTo TirarTodosLosItems_Err
    Dim i         As Byte
    Dim NuevaPos  As t_WorldPos
    Dim MiObj     As t_Obj
    Dim ItemIndex As Integer
    With UserList(UserIndex)
        If ((.pos.Map = 58 Or .pos.Map = 59 Or .pos.Map = 60 Or .pos.Map = 61) And EnEventoFaccionario) Then Exit Sub
        ' Tambien se cae el oro de la billetera
        Dim GoldToDrop As Long
        GoldToDrop = .Stats.GLD - (SvrConfig.GetValue("OroPorNivelBilletera") * .Stats.ELV)
        If GoldToDrop > 0 And Not EsGM(UserIndex) Then
            Call TirarOro(GoldToDrop, UserIndex)
        End If
        For i = 1 To .CurrentInventorySlots
            ItemIndex = .invent.Object(i).ObjIndex
            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) And PirataCaeItem(UserIndex, i) And (Not EsNewbie(UserIndex) Or Not ItemNewbie(ItemIndex)) Then
                    NuevaPos.x = 0
                    NuevaPos.y = 0
                    MiObj.amount = DropAmmount(.invent, i)
                    MiObj.ObjIndex = ItemIndex
                    MiObj.ElementalTags = .invent.Object(i).ElementalTags
                    If .flags.Navegando Then
                        Call Tilelibre(.pos, NuevaPos, MiObj, True, True)
                    Else
                        Call Tilelibre(.pos, NuevaPos, MiObj, .flags.Navegando = True, (Not .flags.Navegando) = True)
                        Call ClosestLegalPos(.pos, NuevaPos, .flags.Navegando, Not .flags.Navegando)
                    End If
                    If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then
                        Call DropObj(UserIndex, i, MiObj.amount, NuevaPos.Map, NuevaPos.x, NuevaPos.y)
                        '  Si no hay lugar, quemamos el item del inventario (nada de mochilas gratis)
                    Else
                        Call QuitarUserInvItem(UserIndex, i, MiObj.amount)
                        Call UpdateUserInv(False, UserIndex, i)
                    End If
                End If
            End If
        Next i
    End With
    Exit Sub
TirarTodosLosItems_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.TirarTodosLosItems", Erl)
End Sub

Function DropAmmount(ByRef invent As t_Inventario, ByVal objectIndex As Integer) As Integer
    DropAmmount = invent.Object(objectIndex).amount
    If invent.EquippedAmuletAccesoryObjIndex > 0 Then
        With ObjData(invent.EquippedAmuletAccesoryObjIndex)
            If .EfectoMagico = 12 Then
                Dim unprotected As Single
                unprotected = 1
                If invent.Object(objectIndex).ObjIndex = ORO_MINA Then 'ore types
                    unprotected = CSng(1) - (CSng(.LingO) / 100)
                ElseIf invent.Object(objectIndex).ObjIndex = PLATA_MINA Then
                    unprotected = CSng(1) - (CSng(.LingP) / 100)
                ElseIf invent.Object(objectIndex).ObjIndex = HIERRO_MINA Then
                    unprotected = CSng(1) - (CSng(.LingH) / 100)
                ElseIf invent.Object(objectIndex).ObjIndex = Wood Then ' wood types
                    unprotected = CSng(1) - (CSng(.Madera) / 100)
                ElseIf invent.Object(objectIndex).ObjIndex = ElvenWood Then
                    unprotected = CSng(1) - (CSng(.MaderaElfica) / 100)
                ElseIf invent.Object(objectIndex).ObjIndex = PinoWood Then
                    unprotected = CSng(1) - (CSng(.MaderaPino) / 100)
                ElseIf invent.Object(objectIndex).ObjIndex = BLODIUM_MINA Then
                    unprotected = CSng(1) - (CSng(.Blodium) / 100)
                ElseIf invent.Object(objectIndex).ObjIndex > 0 Then 'fish types
                    If ObjData(invent.Object(objectIndex).ObjIndex).OBJType = otUseOnce And ObjData(invent.Object(objectIndex).ObjIndex).Subtipo = e_UseOnceSubType.eFish Then
                        unprotected = CSng(1) - (CSng(.MaxItems) / 100)
                    End If
                End If
                DropAmmount = Int(DropAmmount * unprotected)
            End If
        End With
    End If
End Function

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
    On Error GoTo ItemNewbie_Err
    ItemNewbie = ObjData(ItemIndex).Newbie = 1
    Exit Function
ItemNewbie_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.ItemNewbie", Erl)
End Function

Public Function IsItemInCooldown(ByRef User As t_User, ByRef obj As t_UserOBJ) As Boolean
    Dim ElapsedTime As Double
    ElapsedTime = TicksElapsed(User.CdTimes(ObjData(obj.ObjIndex).cdType), GetTickCountRaw())
    IsItemInCooldown = ElapsedTime < ObjData(obj.ObjIndex).Cooldown
End Function

Public Sub UserTargetableItem(ByVal UserIndex As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
    On Error GoTo UserTargetableItem_Err
    With UserList(UserIndex)
        If IsItemInCooldown(UserList(UserIndex), .invent.Object(.flags.UsingItemSlot)) Then
            Exit Sub
        End If
        If .flags.UsingItemSlot = 0 Then Exit Sub
        Dim ObjIndex As Integer
        ObjIndex = .invent.Object(.flags.UsingItemSlot).ObjIndex
        With ObjData(ObjIndex)
            If .MinHp > UserList(UserIndex).Stats.MinHp Then
                Call WriteLocaleMsg(UserIndex, MsgRequiresMoreHealth, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If .MinSta > UserList(UserIndex).Stats.MinSta Then
                'Msg2129=¡No tengo energía!
                Call SendData(SendTarget.ToIndex, UserIndex, PrepareLocalizedChatOverHead(2129, UserList(UserIndex).Char.charindex, vbWhite))
                'Msg420=Estas muy cansado para realizar esta acción.
                Call WriteLocaleMsg(UserIndex, MsgTiredToPerformAction, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Select Case .Subtipo
                Case e_UssableOnTarget.eRessurectionItem
                    Call ResurrectWithItem(UserIndex)
                Case e_UssableOnTarget.eTrap
                    Call PlaceTrap(UserIndex, TileX, TileY)
                Case e_UssableOnTarget.eArpon
                    Call UseArpon(UserIndex)
                Case e_UssableOnTarget.eHandCannon
                    Call UseHandCannon(UserIndex, TileX, TileY)
            End Select
        End With
        .flags.UsingItemSlot = 0
    End With
    Exit Sub
UserTargetableItem_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.UserTargetableItem", Erl)
End Sub

Public Sub ResurrectWithItem(ByVal UserIndex As Integer)
    On Error GoTo ResurrectWithItem_Err
    With UserList(UserIndex)
        Dim CanHelpResult As e_InteractionResult
        If Not IsValidUserRef(.flags.TargetUser) Then
            Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.TargetUser.ArrayIndex = UserIndex Then
            Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim TargetUser As Integer
        TargetUser = .flags.TargetUser.ArrayIndex
        If UserList(TargetUser).flags.Muerto = 0 Then
            Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        CanHelpResult = UserMod.CanHelpUser(UserIndex, TargetUser)
        If UserList(TargetUser).flags.SeguroResu Then
            ' Msg693=El usuario tiene el seguro de resurrección activado.
            Call WriteLocaleMsg(UserIndex, 693, e_FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(TargetUser, PrepareMessageLocaleMsg(1985, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1985=¬1 está intentando revivirte. Desactiva el seguro de resurrección para permitirle hacerlo.
            Exit Sub
        End If
        If CanHelpResult <> eInteractionOk Then
            Call SendHelpInteractionMessage(UserIndex, CanHelpResult)
        End If
        Dim costoVidaResu As Long
        costoVidaResu = UserList(TargetUser).Stats.ELV * 1.5 + .Stats.MinHp * 0.5
        Call UserMod.ModifyHealth(UserIndex, -costoVidaResu, 1)
        Call ModifyStamina(UserIndex, -UserList(UserIndex).Stats.MinSta, False, 0)
        Dim ObjIndex As Integer
        ObjIndex = .invent.Object(.flags.UsingItemSlot).ObjIndex
        Call UpdateCd(UserIndex, ObjData(ObjIndex).cdType)
        If Not IsConsumableFreeZone(UserIndex) Then
            Call RemoveItemFromInventory(UserIndex, UserList(UserIndex).flags.UsingItemSlot)
        End If
        Call ResurrectUser(TargetUser)
        If IsFeatureEnabled("remove-inv-on-attack") Then
            Call RemoveUserInvisibility(UserIndex)
        End If
    End With
    Exit Sub
ResurrectWithItem_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.ResurrectWithItem", Erl)
End Sub

Public Sub RemoveItemFromInventory(ByVal UserIndex As Integer, ByVal Slot As Integer)
    Call QuitarUserInvItem(UserIndex, Slot, 1)
    Call UpdateUserInv(True, UserIndex, Slot)
End Sub

Public Sub PlaceTrap(ByVal UserIndex As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
    With UserList(UserIndex)
        If Distance(TileX, TileY, .pos.x, .pos.y) > 3 Then
            Call WriteLocaleMsg(UserIndex, MsgToFar, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not CanAddTrapAt(.pos.Map, TileX, TileY) Then
            Call WriteLocaleMsg(UserIndex, MsgInvalidTile, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim i              As Integer
        Dim OlderTrapTime  As Long
        Dim OlderTrapIndex As Integer
        OlderTrapTime = 0
        Dim TrapCount As Integer
        Dim Trap      As clsTrap
        For i = 0 To .EffectOverTime.EffectCount - 1
            If .EffectOverTime.EffectList(i).TypeId = e_EffectOverTimeType.eTrap Then
                TrapCount = TrapCount + 1
                Set Trap = .EffectOverTime.EffectList(i)
                If Trap.ElapsedTime > OlderTrapTime Then
                    OlderTrapIndex = i
                    OlderTrapTime = Trap.ElapsedTime
                End If
            End If
        Next i
        If TrapCount >= 3 Then
            Set Trap = .EffectOverTime.EffectList(OlderTrapIndex)
            Call Trap.Disable
        End If
        Dim ObjIndex As Integer
        ObjIndex = UserList(UserIndex).invent.Object(UserList(UserIndex).flags.UsingItemSlot).ObjIndex
        Call UpdateCd(UserIndex, ObjData(ObjIndex).cdType)
        Call EffectsOverTime.CreateTrap(UserIndex, eUser, .pos.Map, TileX, TileY, ObjData(ObjIndex).EfectoMagico)
        Call RemoveItemFromInventory(UserIndex, UserList(UserIndex).flags.UsingItemSlot)
    End With
End Sub

Public Sub UseArpon(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim CanAttackResult As e_AttackInteractionResult
        Dim TargetRef       As t_AnyReference
        If IsValidUserRef(.flags.TargetUser) Then
            Call CastUserToAnyRef(.flags.TargetUser, TargetRef)
        Else
            Call CastNpcToAnyRef(.flags.TargetNPC, TargetRef)
        End If
        If Not IsValidRef(TargetRef) Then
            Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If TargetRef.RefType = eUser Then
            If UserList(TargetRef.ArrayIndex).flags.Muerto <> 0 Then
                Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If TargetRef.RefType = eUser And TargetRef.ArrayIndex = UserIndex Then
                Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        CanAttackResult = UserCanAttack(UserIndex, UserList(UserIndex).VersionId, TargetRef)
        If CanAttackResult <> e_AttackInteractionResult.eCanAttack Then
            Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim ObjIndex As Integer
        ObjIndex = .invent.Object(.flags.UsingItemSlot).ObjIndex
        Call UpdateCd(UserIndex, ObjData(ObjIndex).cdType)
        Dim Damage As Integer
        Damage = GetUserDamageWithItem(UserIndex, ObjIndex, 0)
        If TargetRef.RefType = eUser Then
            UserList(TargetRef.ArrayIndex).Counters.timeFx = 3
            Call RemoveUserInvisibility(UserIndex)
            Call SendData(SendTarget.ToPCAliveArea, TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(TargetRef.ArrayIndex).Char.charindex, FXSANGRE, 0, UserList( _
                    TargetRef.ArrayIndex).pos.x, UserList(TargetRef.ArrayIndex).pos.y))
            Call SendData(SendTarget.ToPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(TargetRef.ArrayIndex).pos.x, UserList( _
                    TargetRef.ArrayIndex).pos.y))
        Else
            If NpcList(TargetRef.ArrayIndex).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(NpcList(TargetRef.ArrayIndex).flags.Snd2, NpcList( _
                        TargetRef.ArrayIndex).pos.x, NpcList(TargetRef.ArrayIndex).pos.y))
            Else
                Call SendData(SendTarget.ToNPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(SND_IMPACTO2, NpcList(TargetRef.ArrayIndex).pos.x, NpcList( _
                        TargetRef.ArrayIndex).pos.y))
            End If
        End If
        If DoDamageToTarget(UserIndex, TargetRef, Damage, e_phisical, ObjIndex) = eStillAlive Then
            If TargetRef.RefType = eUser Then
                If Not IsSet(UserList(TargetRef.ArrayIndex).flags.StatusMask, eCCInmunity) Then
                    Call CreateEffect(UserIndex, eUser, TargetRef.ArrayIndex, TargetRef.RefType, ObjData(ObjIndex).ApplyEffectId)
                End If
            End If
        End If
        If .flags.Oculto = 0 Then
            Dim TargetPos As t_WorldPos
            TargetPos = GetPosition(TargetRef)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(.pos.x, .pos.y, TargetPos.x, TargetPos.y, ObjData(ObjIndex).ProjectileType))
        End If
    End With
End Sub

Public Sub UseHandCannon(ByVal UserIndex As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
    With UserList(UserIndex)
        If Distance(TileX, TileY, .pos.x, .pos.y) > 10 Then
            Call WriteLocaleMsg(UserIndex, MsgToFar, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim ObjIndex As Integer
        ObjIndex = .invent.Object(.flags.UsingItemSlot).ObjIndex
        Call UpdateCd(UserIndex, ObjData(ObjIndex).cdType)
        Dim Particula As Integer
        Dim Tiempo    As Long
        Particula = val(ReadField(1, ObjData(ObjIndex).CreaParticula, Asc(":")))
        Tiempo = val(ReadField(2, ObjData(ObjIndex).CreaParticula, Asc(":")))
        UserList(UserIndex).Counters.timeFx = 3
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, Particula, Tiempo, False, , UserList(UserIndex).pos.x, _
                UserList(UserIndex).pos.y))
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(.pos.x, .pos.y, TileX, TileY, ObjData(ObjIndex).ProjectileType))
        Call CreateDelayedBlast(UserIndex, eUser, .pos.Map, TileX, TileY, ObjData(ObjIndex).ApplyEffectId, ObjIndex)
        If ObjData(ObjIndex).Snd1 <> 0 Then Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(ObjData(ObjIndex).Snd1, .pos.x, .pos.y))
        Call RemoveUserInvisibility(UserIndex)
    End With
End Sub

Public Sub AddOrResetEffect(ByVal UserIndex As Integer, ByVal EffectId As Integer)
    With UserList(UserIndex)
        Dim Effect As IBaseEffectOverTime
        Set Effect = EffectsOverTime.FindEffectOnTarget(UserIndex, .EffectOverTime, EffectId)
        If Effect Is Nothing Then
            Call CreateEffect(UserIndex, eUser, UserIndex, eUser, EffectId)
        Else
            If EffectOverTime(EffectId).Override Then
                Call Effect.Reset(UserIndex, eUser, EffectId)
            End If
        End If
    End With
End Sub

Public Sub UpdateCharWithEquipedItems(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        
        If .flags.Navegando > 0 Then
            Call EquiparBarco(UserIndex)
            .Char.CascoAnim = 0
            .Char.CartAnim = 0
            .Char.ShieldAnim = 0
            .Char.WeaponAnim = 0
            .Char.BackpackAnim = 0
            'TODO place ship body
            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, _
                    .Char.BackpackAnim)
            Exit Sub
        End If
        
        If .flags.Muerto = 1 Then
            .Char.body = iCuerpoMuerto
            .Char.head = 0
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            .Char.CartAnim = NoCart
            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, _
                    .Char.BackpackAnim)
            Exit Sub
        End If
        
        .Char.head = .OrigChar.head
        
        'Armas
        If .invent.EquippedWeaponObjIndex > 0 Then
            If .Invent_Skins.ObjIndexWeaponEquipped > 0 Then
                .Char.WeaponAnim = ObjData(.Invent_Skins.ObjIndexWeaponEquipped).WeaponAnim
            Else
                .Char.WeaponAnim = ObjData(.invent.EquippedWeaponObjIndex).WeaponAnim
            End If
        ElseIf .invent.EquippedWorkingToolObjIndex > 0 Then
            .Char.WeaponAnim = ObjData(.invent.EquippedWorkingToolObjIndex).WeaponAnim
        Else
            .Char.WeaponAnim = 0
        End If
        
        'Armaduras/túnicas/vestimentas/etc
        If .invent.EquippedArmorObjIndex > 0 Then
            If .Invent_Skins.ObjIndexArmourEquipped > 0 Then
                .Char.body = ObtenerRopaje(UserIndex, ObjData(.Invent_Skins.ObjIndexArmourEquipped))
            Else
                .Char.body = ObtenerRopaje(UserIndex, ObjData(.invent.EquippedArmorObjIndex))
            End If
        Else
            Call SetNakedBody(UserList(UserIndex))
        End If
        
        'Cascos
        If .invent.EquippedHelmetObjIndex > 0 Then
            If .Invent_Skins.ObjIndexHelmetEquipped > 0 Then
                .Char.CascoAnim = ObjData(.Invent_Skins.ObjIndexHelmetEquipped).CascoAnim
            Else
                .Char.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
            End If
        Else
            .Char.CascoAnim = 0
        End If
        
        'Amuletos/accesorios
        If .invent.EquippedAmuletAccesoryObjIndex > 0 Then
            .Char.CartAnim = ObjData(.invent.EquippedAmuletAccesoryObjIndex).Ropaje
        Else
            .Char.CartAnim = 0
        End If
        
        'Escudos
        If .invent.EquippedShieldObjIndex > 0 Then
            If .Invent_Skins.ObjIndexShieldEquipped > 0 Then
                .Char.ShieldAnim = ObjData(.Invent_Skins.ObjIndexShieldEquipped).ShieldAnim
            Else
                .Char.ShieldAnim = ObjData(.invent.EquippedShieldObjIndex).ShieldAnim
            End If
        Else
            .Char.ShieldAnim = 0
        End If
        
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, _
                .Char.BackpackAnim)
    End With
    
End Sub

Function RemoveGold(ByVal UserIndex As Integer, ByVal amount As Long) As Boolean
    With UserList(UserIndex)
        If .Stats.GLD < amount Then Exit Function
        .Stats.GLD = .Stats.GLD - amount
        Call WriteUpdateGold(UserIndex)
        RemoveGold = True
    End With
End Function

Sub AddGold(ByVal UserIndex As Integer, ByVal amount As Long)
    With UserList(UserIndex)
        .Stats.GLD = .Stats.GLD + amount
        Call WriteUpdateGold(UserIndex)
    End With
End Sub

Function ObtenerRopaje(ByVal UserIndex As Integer, ByRef obj As t_ObjData) As Integer
    Dim race As e_Raza
    race = UserList(UserIndex).raza
    Dim EsMujer As Boolean
    EsMujer = UserList(UserIndex).genero = e_Genero.Mujer
    Dim EsRazaBaja As Boolean
    EsRazaBaja = (race = e_Raza.Gnomo Or race = e_Raza.Enano)
    If obj.OBJType = e_OBJType.otSaddles Then
        If EsRazaBaja Then
            If obj.RazaBajos > 0 Then
                ObtenerRopaje = obj.RazaBajos
                Exit Function
            End If
        Else
            If obj.RazaAltos > 0 Then
                ObtenerRopaje = obj.RazaAltos
                Exit Function
            End If
        End If
    End If
    Select Case race
        Case e_Raza.Humano
            If EsMujer And obj.RopajeHumana > 0 Then
                ObtenerRopaje = obj.RopajeHumana
                Exit Function
            ElseIf obj.RopajeHumano > 0 Then
                ObtenerRopaje = obj.RopajeHumano
                Exit Function
            End If
        Case e_Raza.Elfo
            If EsMujer And obj.RopajeElfa > 0 Then
                ObtenerRopaje = obj.RopajeElfa
                Exit Function
            ElseIf obj.RopajeElfo > 0 Then
                ObtenerRopaje = obj.RopajeElfo
                Exit Function
            End If
        Case e_Raza.Drow
            If EsMujer And obj.RopajeElfaOscura > 0 Then
                ObtenerRopaje = obj.RopajeElfaOscura
                Exit Function
            ElseIf obj.RopajeElfoOscuro > 0 Then
                ObtenerRopaje = obj.RopajeElfoOscuro
                Exit Function
            End If
        Case e_Raza.Orco
            If EsMujer And obj.RopajeOrca > 0 Then
                ObtenerRopaje = obj.RopajeOrca
                Exit Function
            ElseIf obj.RopajeOrco > 0 Then
                ObtenerRopaje = obj.RopajeOrco
                Exit Function
            End If
        Case e_Raza.Enano
            If EsMujer And obj.RopajeEnana > 0 Then
                ObtenerRopaje = obj.RopajeEnana
                Exit Function
            ElseIf obj.RopajeEnano > 0 Then
                ObtenerRopaje = obj.RopajeEnano
                Exit Function
            End If
        Case e_Raza.Gnomo
            If EsMujer And obj.RopajeGnoma > 0 Then
                ObtenerRopaje = obj.RopajeGnoma
                Exit Function
            ElseIf obj.RopajeGnomo > 0 Then
                ObtenerRopaje = obj.RopajeGnomo
                Exit Function
            End If
    End Select
    ObtenerRopaje = obj.Ropaje
End Function

Sub EliminarLlaves(ByVal ClaveLlave As Integer, ByVal UserIndex As Integer)
    ' Abrir el archivo "Eliminarllaves.dat" para lectura
    Open "Eliminarllaves.dat" For Input As #1
    ' Variables para el almacenamiento temporal de datos
    Dim Linea           As String
    Dim clave           As Integer
    Dim Objeto          As Integer
    Dim LlaveEncontrada As Boolean
    LlaveEncontrada = False
    ' Leer cada línea del archivo
    Do Until EOF(1) Or LlaveEncontrada
        Line Input #1, Linea
        If InStr(Linea, "Clave=" & ClaveLlave) > 0 Then
            ' Se encontró la llave con la clave buscada
            LlaveEncontrada = True
            Do Until EOF(1)
                Line Input #1, Linea
                If InStr(Linea, "Objeto=") > 0 Then
                    Objeto = val(mid(Linea, InStr(Linea, "=") + 1))
                    ' Llamar a QuitarObjeto con el objeto encontrado
                    Call QuitarObjetos(Objeto, 1000, UserIndex)
                ElseIf InStr(Linea, "[Llave") > 0 Then
                    ' Se ha alcanzado el final de la sección de la llave actual
                    Exit Do
                End If
            Loop
        End If
    Loop
    ' Cerrar el archivo
    Close #1
End Sub

Public Function CanElementalTagBeApplied(ByVal UserIndex As Integer, ByVal TargetSlot As Integer, ByVal SourceSlot As Integer) As Boolean
    CanElementalTagBeApplied = False
    Dim TargetObj As t_ObjData
    Dim SourceObj As t_ObjData
    If TargetSlot < 1 Or TargetSlot > UserList(UserIndex).CurrentInventorySlots Then
        Exit Function
    End If
    If SourceSlot < 1 Or SourceSlot > UserList(UserIndex).CurrentInventorySlots Then
        Exit Function
    End If
    If UserList(UserIndex).invent.Object(TargetSlot).ObjIndex = 0 Or UserList(UserIndex).invent.Object(SourceSlot).ObjIndex = 0 Then
        Exit Function
    End If
    TargetObj = ObjData(UserList(UserIndex).invent.Object(TargetSlot).ObjIndex)
    SourceObj = ObjData(UserList(UserIndex).invent.Object(SourceSlot).ObjIndex)
    If SourceObj.OBJType <> otElementalRune Then
        Exit Function
    End If
    If TargetObj.OBJType <> otWeapon Then
        Exit Function
    End If
    If TargetObj.ElementalTags <> e_ElementalTags.Normal Then
        Call WriteLocaleMsg(UserIndex, 2087, e_FontTypeNames.FONTTYPE_INFOIAO)
        Exit Function
    End If
    If UserList(UserIndex).invent.Object(TargetSlot).ElementalTags <> e_ElementalTags.Normal Then
        Call WriteLocaleMsg(UserIndex, 2087, e_FontTypeNames.FONTTYPE_INFOIAO)
        Exit Function
    End If
    If UserList(UserIndex).invent.Object(TargetSlot).amount > 1 Then
        Call WriteLocaleMsg(UserIndex, 2088, e_FontTypeNames.FONTTYPE_INFOIAO)
        Exit Function
    End If
    Select Case SourceObj.ElementalTags
        Case e_ElementalTags.Fire
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.Incinerar, 10, False))
        Case e_ElementalTags.Water
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.CurarCrimi, 10, False))
        Case e_ElementalTags.Earth
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.PoisonGas, 10, False))
        Case e_ElementalTags.Wind
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_GraphicEffects.Runa, 10, False))
        Case Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticleEffects.Corazones, 10, False))
    End Select
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_SoundEffects.RUNE_SOUND, NO_3D_SOUND, NO_3D_SOUND))
    UserList(UserIndex).invent.Object(TargetSlot).ElementalTags = SourceObj.ElementalTags
    CanElementalTagBeApplied = True
End Function

Sub SkinEquip(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal ObjIndex As Integer, Optional ByVal bLoggin As Boolean = False)

Dim bNeedChangeUserChar         As Boolean
Dim nuevoHead                   As Integer
Dim nuevoCasco                  As Integer
Dim i                           As Integer
Dim obj                         As t_ObjData
Dim eSkinType                   As e_OBJType

    On Error GoTo SkinEquip_Error:

    If ObjIndex > 0 Then
        obj = ObjData(ObjIndex)
    Else
        Exit Sub
    End If

    With UserList(UserIndex)
    
        If .Invent_Skins.Object(Slot).ObjIndex = 0 Then Exit Sub
        eSkinType = ObjData(.Invent_Skins.Object(Slot).ObjIndex).OBJType
    
        Select Case eSkinType
            Case e_OBJType.otSkinsArmours

                For i = 1 To MAX_SKINSINVENTORY_SLOTS
                    If .Invent_Skins.Object(i).Equipped And .Invent_Skins.Object(i).ObjIndex = .Invent_Skins.ObjIndexArmourEquipped Then
                        Call DesequiparSkin(UserIndex, i)
                        Exit For
                    End If
                Next i

                .Invent_Skins.Object(Slot).Equipped = True
                .Invent_Skins.ObjIndexArmourEquipped = ObjIndex
                .Invent_Skins.SlotArmourEquipped = Slot

                If .flags.Mimetizado = 1 Then
                    .CharMimetizado.body = obj.Ropaje
                Else
                    If .flags.Navegando = 0 Then
                        .OrigChar.body = .Char.body
                        .Char.body = ObtenerRopaje(UserIndex, obj)
                        bNeedChangeUserChar = True
                    Else
                        .OrigChar.body = obj.Ropaje
                    End If
                End If

            Case e_OBJType.otSkinsSpells
                'Buscamos otros skins de este mismo hechizo equipados y lo desequipamos.
                For i = 1 To MAX_SKINSINVENTORY_SLOTS
                    If .Invent_Skins.Object(i).ObjIndex > 0 Then
                        If ObjData(.Invent_Skins.Object(i).ObjIndex).OBJType = e_OBJType.otSkinsSpells Then
                            If ObjData(.Invent_Skins.Object(i).ObjIndex).HechizoIndex = ObjData(.Invent_Skins.Object(Slot).ObjIndex).HechizoIndex And Slot <> i Then
                                Call DesequiparSkin(UserIndex, i)
                                Exit For
                            End If
                        End If
                    End If
                Next i

                'Equipamos ahora nuestro nuevo skin
                .Invent_Skins.Object(Slot).Equipped = True
                .Invent_Skins.Object(Slot).Type = e_OBJType.otSkinsSpells

                If ObjData(.Invent_Skins.Object(Slot).ObjIndex).HechizoIndex > 0 Then
                    .Stats.UserSkinsHechizos(ObjData(.Invent_Skins.Object(Slot).ObjIndex).HechizoIndex) = ObjData(.Invent_Skins.Object(Slot).ObjIndex).CreaFX
                End If

            Case e_OBJType.otSkinsHelmets
            
                'If .invent.EquippedHelmetObjIndex > 0 Then
                    For i = 1 To MAX_SKINSINVENTORY_SLOTS
                        If .Invent_Skins.Object(i).Equipped And .Invent_Skins.Object(i).ObjIndex = .Invent_Skins.ObjIndexHelmetEquipped Then
                            Call DesequiparSkin(UserIndex, i)
                            Exit For
                        End If
                    Next i
    
                    .Invent_Skins.Object(Slot).Equipped = True
                    .Invent_Skins.ObjIndexHelmetEquipped = ObjIndex
                    .Invent_Skins.SlotHelmetEquipped = Slot
    
                    If .flags.Mimetizado = 1 Then
                        .CharMimetizado.body = obj.Ropaje
                    Else
                        If .flags.Navegando = 0 Then
                            If obj.CascoAnim > 0 Then
                                If obj.Subtipo = 2 Then
                                    ' Si el casco cambia la cabeza entera
                                    If .Char.head < PATREON_HEAD Then
                                        .Char.originalhead = .Char.head
                                    End If
                                    nuevoHead = obj.CascoAnim
                                    nuevoCasco = NingunCasco
                                Else
                                    ' Si el casco se superpone (no reemplaza la cabeza)
                                    nuevoHead = .Char.head
                                    If .Char.head >= PATREON_HEAD Then
                                        nuevoCasco = NingunCasco
                                    Else
                                        nuevoCasco = obj.CascoAnim
                                    End If
                                End If
                                
                                ' Asignar cambios y aplicar actualización visual
                                .Char.head = nuevoHead
                                .OrigChar.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
                                .Char.CascoAnim = nuevoCasco
    
                                bNeedChangeUserChar = True
                            End If
                        Else
                            .OrigChar.CascoAnim = obj.CascoAnim
                        End If
                    End If
                'End If
                
            Case e_OBJType.otSkinsWings

                For i = 1 To MAX_SKINSINVENTORY_SLOTS
                    If .Invent_Skins.Object(i).Equipped And .Invent_Skins.Object(i).ObjIndex = .Invent_Skins.ObjIndexWindsEquipped Then
                        Call DesequiparSkin(UserIndex, i)
                        Exit Sub
                    End If
                Next i

                'Lo equipa
                If Len(obj.CreaGRH) <> 0 Then
                    .Char.Body_Aura = obj.CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Body_Aura, False, 2))
                End If

                .Invent_Skins.Object(Slot).Equipped = True
                .Invent_Skins.ObjIndexBackpackEquipped = .Invent_Skins.Object(Slot).ObjIndex
                .Invent_Skins.SlotBackpackEquipped = Slot
                .Invent_Skins.Object(Slot).Type = e_OBJType.otSkinsWings

                If .flags.Montado = 0 And .flags.Navegando = 0 Then
                    If ObtenerRopaje(UserIndex, obj) > 0 Then ' ;)
                        .Char.BackpackAnim = ObtenerRopaje(UserIndex, obj)
                        bNeedChangeUserChar = True
                    End If
                End If
            
            Case e_OBJType.otSkinsBoats
            
                For i = 1 To MAX_SKINSINVENTORY_SLOTS
                    If .Invent_Skins.Object(i).Equipped And .Invent_Skins.Object(i).ObjIndex = .Invent_Skins.ObjIndexBoatEquipped Then
                        Call DesequiparSkin(UserIndex, i)
                        Exit For
                    End If
                Next i

                .Invent_Skins.Object(Slot).Equipped = True
                .Invent_Skins.ObjIndexBoatEquipped = ObjIndex
                .Invent_Skins.SlotBoatEquipped = Slot
                .Invent_Skins.Object(Slot).Type = e_OBJType.otSkinsBoats
                
                If .flags.Mimetizado = 1 Then
                    .CharMimetizado.body = obj.Ropaje
                    If .flags.Navegando = 0 Then
                        .OrigChar.body = .Char.body
                        .CharMimetizado.body = obj.Ropaje
                        bNeedChangeUserChar = True
                    Else
                        .OrigChar.body = obj.Ropaje
                    End If
                Else
                    .OrigChar.body = .Char.body
                    .Char.body = ObtenerRopaje(UserIndex, obj)
                    bNeedChangeUserChar = True
                End If
            
            Case e_OBJType.otSkinsShields
            
                For i = 1 To MAX_SKINSINVENTORY_SLOTS
                    If .Invent_Skins.Object(i).Equipped And .Invent_Skins.Object(i).ObjIndex = .Invent_Skins.ObjIndexShieldEquipped Then
                        Call DesequiparSkin(UserIndex, i)
                        Exit For
                    End If
                Next i

                .Invent_Skins.Object(Slot).Equipped = True
                .Invent_Skins.ObjIndexShieldEquipped = ObjIndex
                .Invent_Skins.SlotShieldEquipped = Slot
                
                If .flags.Mimetizado = 1 Then
                    .CharMimetizado.ShieldAnim = obj.ShieldAnim
                Else
                    If .flags.Navegando = 0 Then
                        .OrigChar.ShieldAnim = .Char.ShieldAnim
                        .Char.ShieldAnim = obj.ShieldAnim
                        bNeedChangeUserChar = True
                    Else
                        .OrigChar.ShieldAnim = obj.ShieldAnim
                    End If
                End If
            
            Case e_OBJType.otSkinsWeapons
            
                For i = 1 To MAX_SKINSINVENTORY_SLOTS
                    If .Invent_Skins.Object(i).Equipped And .Invent_Skins.Object(i).ObjIndex = .Invent_Skins.ObjIndexWeaponEquipped Then
                        Call DesequiparSkin(UserIndex, i)
                        Exit For
                    End If
                Next i

                .Invent_Skins.Object(Slot).Equipped = True
                .Invent_Skins.ObjIndexWeaponEquipped = ObjIndex
                .Invent_Skins.SlotWeaponEquipped = Slot
                
                If .flags.Mimetizado = 1 Then
                    .CharMimetizado.WeaponAnim = obj.WeaponAnim
                Else
                    If .flags.Navegando = 0 Then
                        .OrigChar.WeaponAnim = .Char.WeaponAnim
                        .Char.WeaponAnim = obj.WeaponAnim
                        bNeedChangeUserChar = True
                    Else
                        .OrigChar.WeaponAnim = obj.WeaponAnim
                    End If
                End If
            
        End Select
        
        If bNeedChangeUserChar And Not bLoggin Then
            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
        End If
        
        Call WriteChangeSkinSlot(UserIndex, ObjData(.Invent_Skins.Object(Slot).ObjIndex).OBJType, Slot)
        
    End With
    
    On Error GoTo 0
    Exit Sub
SkinEquip_Error:
    Call Logging.TraceError(Err.Number, Err.Description, "InvUsuario.SkinEquip of Módulo", Erl())
    
End Sub

Public Function AddSkin(ByVal UserIndex As Integer, ByVal SkinIndex As Integer) As Boolean

Dim bAdded                      As Boolean
Dim i                           As Byte
    
    On Error GoTo AddSkin_Error
    
    If SkinIndex = 0 Then Exit Function
    
    If Not IsPatreon(UserIndex) Then
        Call WriteLocaleMsg(UserIndex, 2103, e_FontTypeNames.FONTTYPE_INFO) 'Msg2103=Necesitas mejorar tu cuenta para poder agregar Skins. Para más información visita: https://www.patreon.com/nolandstudios
        Exit Function
    End If
    
    With UserList(UserIndex)
        For i = 1 To MAX_SKINSINVENTORY_SLOTS
            If .Invent_Skins.Object(i).ObjIndex = 0 Or .Invent_Skins.Object(i).Deleted Then
                .Invent_Skins.Object(i).ObjIndex = SkinIndex
                .Invent_Skins.Object(i).Equipped = False
                If Not .Invent_Skins.Object(i).Deleted Then
                    .Invent_Skins.count = .Invent_Skins.count + 1
                Else
                    .Invent_Skins.Object(i).Deleted = False
                End If
                If SkinIndex > 0 Then
                    Call WriteChangeSkinSlot(UserIndex, ObjData(SkinIndex).OBJType, i)
                End If
                Call WriteLocaleMsg(UserIndex, 2097, e_FontTypeNames.FONTTYPE_INFO, ObjData(SkinIndex).name) 'Msg2097=Has agregado con éxito tu nueva skin (" & ObjData(SkinIndex).name & "). Equipala desde el inventario de skins.
                AddSkin = True
                Exit Function
            End If
            
            If i = MAX_SKINSINVENTORY_SLOTS And .Invent_Skins.Object(i).ObjIndex > 0 Then
                Call LogShopTransactions("PJ ID: " & .Id & " Nick: " & .name & " -> Llegó al máximo de skins en su inventario de skins.")
            End If
        Next i

        AddSkin = False
        Call WriteLocaleMsg(UserIndex, 2098, e_FontTypeNames.FONTTYPE_INFO) 'Msg2098=Ya no tienes lugar en el inventario de Skins.
    End With
    AddSkin = False
    On Error GoTo 0
    Exit Function
AddSkin_Error:
    AddSkin = False
    Call Logging.TraceError(Err.Number, Err.Description, "InvUsuario.AddSkin of Módulo", Erl())
    
End Function

Function HaveThisSkin(ByVal UserIndex As Integer, ByVal SkinIndex As Integer) As Boolean

Dim i                           As Byte
    
    On Error GoTo HaveThisSkin_Error
    
    With UserList(UserIndex)
        If SkinIndex = 0 Then Exit Function

        For i = 1 To .Invent_Skins.count
            If .Invent_Skins.Object(i).ObjIndex = SkinIndex Then
                HaveThisSkin = True
                Exit Function
            End If
        Next i

        HaveThisSkin = False
    End With
    
    On Error GoTo 0
    Exit Function
HaveThisSkin_Error:
    HaveThisSkin = False
    Call Logging.TraceError(Err.Number, Err.Description, "InvUsuario.HaveThisSkin of Módulo", Erl())
End Function

Function CanEquipSkin(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal bFromInvent As Boolean) As Boolean

Dim bCanUser                    As Boolean
Dim bDonante                    As Boolean
Dim eSkinType                   As e_OBJType

    On Error GoTo CanEquipSkin_Error

    If Slot <= 0 Then Exit Function

    With UserList(UserIndex)

        If bFromInvent Then
            If .invent.Object(Slot).ObjIndex = 0 Then Exit Function
            bCanUser = ClasePuedeUsarItem(UserIndex, .invent.Object(Slot).ObjIndex) And SexoPuedeUsarItem(UserIndex, .invent.Object(Slot).ObjIndex) And RazaPuedeUsarItem(UserIndex, .invent.Object(Slot).ObjIndex) And FaccionPuedeUsarItem(UserIndex, .invent.Object(Slot).ObjIndex) And LevelCanUseItem(UserIndex, ObjData(.invent.Object(Slot).ObjIndex))
        Else
            bCanUser = ClasePuedeUsarItem(UserIndex, .Invent_Skins.Object(Slot).ObjIndex) And SexoPuedeUsarItem(UserIndex, .Invent_Skins.Object(Slot).ObjIndex) And FaccionPuedeUsarItem(UserIndex, .Invent_Skins.Object(Slot).ObjIndex)
        End If

        eSkinType = ObjData(.Invent_Skins.Object(Slot).ObjIndex).OBJType

        If Not bCanUser Then
            'Añadir el cartel que haga falta
            'Call WriteConsoleMsg(UserIndex, "{315}", e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        '        Revisar si en el futuro pretenden agregar ítems exclusivos para los PATREON, así debería ser la feature, opcional para algunos ítems el requerimiento de Patreon (en este caso Donantes, hay q cambiar los nombres)
        '        If Slot > 0 And Slot <= MAX_INVENTORY_SLOTS Then
        '            If .invent.Object(Slot).ObjIndex > 0 Then
        '                If ObjData(.invent.Object(Slot).ObjIndex).Donante = 0 And ObjData(.invent.Object(Slot).ObjIndex).ValorDonante = 0 Then
        '                    bDonante = False
        '                Else
        '                    bDonante = True
        '                End If
        '            End If
        '        End If
        '
        '        If Not IsSuscribed(UserIndex) Then
        '            If bDonante Then
        '                Call WriteConsoleMsg(UserIndex, "Para equipar este skin, debes tener una suscripción activa.", FontTypeNames.FONTTYPE_PARTY)
        '                Exit Function
        '            End If
        '        End If
        Select Case eSkinType
            Case e_OBJType.otSkinsArmours
                If .invent.EquippedArmorSlot > 0 Then
                    If bFromInvent Then
                        If SkinRequireObject(UserIndex, Slot) Then
                            If .invent.EquippedArmorObjIndex = ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto Then
                                CanEquipSkin = True
                                Exit Function
                            Else
                                Call WriteLocaleMsg(UserIndex, 2099, e_FontTypeNames.FONTTYPE_INFO, ObjData(ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto).name) 'Msg2099=Para equipar este skin, debes tener equipado
                                Exit Function
                            End If
                        Else
                            CanEquipSkin = True
                            Exit Function
                        End If
                    Else
                        If SkinRequireObject(UserIndex, Slot) Then
                            If .invent.EquippedArmorObjIndex = ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto Then
                                CanEquipSkin = True
                                Exit Function
                            Else
                                Call WriteLocaleMsg(UserIndex, 2099, e_FontTypeNames.FONTTYPE_INFO, ObjData(ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto).name) 'Msg2099=Para equipar este skin, debes tener equipado
                                Exit Function
                            End If
                        Else
                            CanEquipSkin = True
                            Exit Function
                        End If
                    End If
                Else
                    Call WriteLocaleMsg(UserIndex, 2100, e_FontTypeNames.FONTTYPE_INFO) 'Msg2100=Para equipar este skin, debes tener equipado un objeto de ese tipo.
                    Exit Function
                End If

            Case e_OBJType.otSkinsWings

                If bFromInvent Then
                    If SkinRequireObject(UserIndex, Slot) Then
                        If .invent.EquippedBackpackObjIndex = ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto Then
                            CanEquipSkin = True
                            Exit Function
                        Else
                            Call WriteLocaleMsg(UserIndex, 2099, e_FontTypeNames.FONTTYPE_INFO, ObjData(ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto).name) 'Msg2099=Para equipar este skin, debes tener equipado
                            Exit Function
                        End If
                    Else
                        CanEquipSkin = True
                        Exit Function
                    End If
                Else
                    If SkinRequireObject(UserIndex, Slot) Then
                        If .invent.EquippedBackpackObjIndex = ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto Then
                            CanEquipSkin = True
                            Exit Function
                        Else
                           Call WriteLocaleMsg(UserIndex, 2099, e_FontTypeNames.FONTTYPE_INFO, ObjData(ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto).name) 'Msg2099=Para equipar este skin, debes tener equipado
                            Exit Function
                        End If
                    Else
                        CanEquipSkin = True
                        Exit Function
                    End If
                End If

            Case e_OBJType.otSkinsHelmets

                If .invent.EquippedHelmetSlot > 0 Then
                    If bFromInvent Then
                        If SkinRequireObject(UserIndex, Slot) Then
                            If .invent.EquippedHelmetObjIndex = ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto Then
                                CanEquipSkin = True
                                Exit Function
                            Else
                                Call WriteLocaleMsg(UserIndex, 2099, e_FontTypeNames.FONTTYPE_INFO, ObjData(ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto).name) 'Msg2099=Para equipar este skin, debes tener equipado
                                Exit Function
                            End If
                        Else
                            CanEquipSkin = True
                            Exit Function
                        End If
                    Else
                        If SkinRequireObject(UserIndex, Slot) Then
                            If .invent.EquippedHelmetObjIndex = ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto Then
                                CanEquipSkin = True
                                Exit Function
                            Else
                                Call WriteLocaleMsg(UserIndex, 2099, e_FontTypeNames.FONTTYPE_INFO, ObjData(ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto).name) 'Msg2099=Para equipar este skin, debes tener equipado
                                Exit Function
                            End If
                        Else
                            CanEquipSkin = True
                            Exit Function
                        End If
                    End If
                Else
                    Call WriteLocaleMsg(UserIndex, 2100, e_FontTypeNames.FONTTYPE_INFO) 'Msg2100=Para equipar este skin, debes tener equipado un objeto de ese tipo.
                    Exit Function
                End If
                
            Case e_OBJType.otSkinsSpells
                CanEquipSkin = True
                Exit Function

            Case e_OBJType.otSkinsBoats
                
                If .invent.EquippedShipObjIndex > 0 Or .flags.Navegando = 1 Then
                    If bFromInvent Then
                        If SkinRequireObject(UserIndex, Slot) Then
                            If .invent.EquippedShipObjIndex = ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto Then
                                CanEquipSkin = True
                                Exit Function
                            Else
                                Call WriteLocaleMsg(UserIndex, 2099, e_FontTypeNames.FONTTYPE_INFO, ObjData(ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto).name) 'Msg2099=Para equipar este skin, debes tener equipado
                                Exit Function
                            End If
                        Else
                            CanEquipSkin = True
                            Exit Function
                        End If
                    Else
                        If SkinRequireObject(UserIndex, Slot) Then
                            If .invent.EquippedShipObjIndex = ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto Then
                                CanEquipSkin = True
                                Exit Function
                            Else
                                Call WriteLocaleMsg(UserIndex, 2099, e_FontTypeNames.FONTTYPE_INFO, ObjData(ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto).name) 'Msg2099=Para equipar este skin, debes tener equipado
                                Exit Function
                            End If
                        Else
                            CanEquipSkin = True
                            Exit Function
                        End If
                    End If
                Else
                    Call WriteLocaleMsg(UserIndex, 2100, e_FontTypeNames.FONTTYPE_INFO) 'Msg2100=Para equipar este skin, debes tener equipado un objeto de ese tipo.
                    Exit Function
                End If

            Case e_OBJType.otSkinsShields
                If .invent.EquippedShieldObjIndex > 0 Then
                    If bFromInvent Then
                        If SkinRequireObject(UserIndex, Slot) Then
                            If .invent.EquippedShieldObjIndex = ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto Then
                                CanEquipSkin = True
                                Exit Function
                            Else
                                Call WriteLocaleMsg(UserIndex, 2099, e_FontTypeNames.FONTTYPE_INFO, ObjData(ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto).name) 'Msg2099=Para equipar este skin, debes tener equipado
                                Exit Function
                            End If
                        Else
                            CanEquipSkin = True
                            Exit Function
                        End If
                    Else
                        If SkinRequireObject(UserIndex, Slot) Then
                            If .invent.EquippedShieldObjIndex = ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto Then
                                CanEquipSkin = True
                                Exit Function
                            Else
                                Call WriteLocaleMsg(UserIndex, 2099, e_FontTypeNames.FONTTYPE_INFO, ObjData(ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto).name) 'Msg2099=Para equipar este skin, debes tener equipado
                                Exit Function
                            End If
                        Else
                            CanEquipSkin = True
                            Exit Function
                        End If
                    End If
                Else
                    Call WriteLocaleMsg(UserIndex, 2100, e_FontTypeNames.FONTTYPE_INFO) 'Msg2100=Para equipar este skin, debes tener equipado un objeto de ese tipo.
                    Exit Function
                End If

            Case e_OBJType.otSkinsWeapons

                If .invent.EquippedWeaponObjIndex > 0 Then
                    If bFromInvent Then
                        If SkinRequireObject(UserIndex, Slot) Then
                            If .invent.EquippedWeaponObjIndex = ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto Then
                                CanEquipSkin = True
                                Exit Function
                            Else
                                Call WriteLocaleMsg(UserIndex, 2099, e_FontTypeNames.FONTTYPE_INFO, ObjData(ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto).name) 'Msg2099=Para equipar este skin, debes tener equipado
                                Exit Function
                            End If
                        Else
                            CanEquipSkin = True
                            Exit Function
                        End If
                    Else
                        If SkinRequireObject(UserIndex, Slot) Then
                            If .invent.EquippedWeaponObjIndex = ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto Then
                                CanEquipSkin = True
                                Exit Function
                            Else
                                Call WriteLocaleMsg(UserIndex, 2099, e_FontTypeNames.FONTTYPE_INFO, ObjData(ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto).name) 'Msg2099=Para equipar este skin, debes tener equipado
                                Exit Function
                            End If
                        Else
                            CanEquipSkin = True
                            Exit Function
                        End If
                    End If
                Else
                    Call WriteLocaleMsg(UserIndex, 2100, e_FontTypeNames.FONTTYPE_INFO) 'Msg2100=Para equipar este skin, debes tener equipado un objeto de ese tipo.
                    Exit Function
                End If
        End Select
    End With
    On Error GoTo 0
    Exit Function
CanEquipSkin_Error:
    CanEquipSkin = False
    Call Logging.TraceError(Err.Number, Err.Description, "InvUsuario.CanEquipSkin of Módulo InvUsuario User: " & UserList(UserIndex).name & " UserIndex: " & UserIndex & " Skin Slot: " & Slot & " eSkinType: " & eSkinType, Erl())

End Function

Sub UpdateSingleItemInv(ByVal UserIndex As Integer, ByVal Slot As Byte, Optional ByVal UpdateFullInfo As Boolean = True)

Dim NullObj                     As t_UserOBJ
    
    On Error GoTo UpdateSingleItemInv_Error
    
    With UserList(UserIndex)
        'Actualiza el inventario
        If .invent.Object(Slot).ObjIndex > 0 Then
            If Slot > 0 Then
                Call ChangeUserInv(UserIndex, Slot, .invent.Object(Slot))    'Actualizamos sólo el slot!
            End If
        Else
            Call ChangeUserInv(UserIndex, Slot, NullObj)
        End If
    End With
    
    On Error GoTo 0
    Exit Sub
UpdateSingleItemInv_Error:
    Call Logging.TraceError(Err.Number, Err.Description, "InvUsuario.UpdateSingleItemInv of Módulo", Erl())
End Sub

Public Function SkinRequireObject(ByVal UserIndex As Integer, ByVal Slot As Byte) As Boolean

   On Error GoTo SkinRequireObject_Error

    With UserList(UserIndex)
        If .Invent_Skins.Object(Slot).ObjIndex > 0 Then
            SkinRequireObject = CBool(ObjData(.Invent_Skins.Object(Slot).ObjIndex).RequiereObjeto > 0)
        End If
    End With

   On Error GoTo 0
   Exit Function

SkinRequireObject_Error:

    Call Logging.TraceError(Err.Number, Err.Description, "InvUsuario.SkinRequireObject", Erl())
    
End Function
Public Sub EquipArrow(ByVal UserIndex As Integer, ByVal Slot As Integer)
    On Error GoTo EquipArrow_Error
    Dim bowIndex  As Integer
    Dim BowCategory   As Byte
    Dim ArrowCategory As Byte
    Dim ArrowObjIndex As Integer

    With UserList(UserIndex).invent
        ArrowObjIndex = .Object(Slot).ObjIndex
        bowIndex = .EquippedWeaponObjIndex
        BowCategory = ObjData(bowIndex).BowCategory
        ArrowCategory = ObjData(ArrowObjIndex).ArrowCategory
        
        ' No hay arco equipado
        If bowIndex <= 0 Then
            'Msg2145=Debes equipar un arco para usar flechas.
            Call WriteLocaleMsg(UserIndex, 2145, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' El arma equipada no es un arco
        If ObjData(bowIndex).WeaponType <> eBow Then
            'Msg2146=El arma equipada no permite usar flechas.
            Call WriteLocaleMsg(UserIndex, 2146, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        ' No se permite flecha de mayor categoria que el arco
        If ArrowCategory > BowCategory Then
            'Msg2147=No podés equipar esta flecha con el arco actual.
            Call WriteLocaleMsg(UserIndex, 2147, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' Quitar flecha previa
        If .EquippedMunitionObjIndex > 0 Then
            Call Desequipar(UserIndex, .EquippedMunitionSlot)
        End If

        ' Equipar flecha
        .Object(Slot).Equipped = 1
        .EquippedMunitionObjIndex = ArrowObjIndex
        .EquippedMunitionSlot = Slot
    End With
    Exit Sub
EquipArrow_Error:
    Call Logging.TraceError(Err.Number, Err.Description, "InvUsuario.EquipArrow", Erl())
End Sub
Public Sub ValidateEquippedArrow(ByVal UserIndex As Integer)
    On Error GoTo ValidateEquippedArrow_Error
    Dim bowIndex   As Integer
    Dim arrowIndex As Integer

    With UserList(UserIndex).invent
        arrowIndex = .EquippedMunitionObjIndex
        bowIndex = .EquippedWeaponObjIndex

        ' No hay flecha equipada ? nada que validar
        If arrowIndex <= 0 Then Exit Sub

        ' No hay arma equipada ? no hacer nada
        If bowIndex <= 0 Then Exit Sub

        ' El arma equipada no es un arco ? no hacer nada
        If ObjData(bowIndex).WeaponType <> eBow Then Exit Sub

        ' Se desequipa la flecha al cambiar a un arco de menor categoría
        If ObjData(bowIndex).BowCategory < ObjData(arrowIndex).ArrowCategory Then
            Call Desequipar(UserIndex, .EquippedMunitionSlot)
            'Msg2148=La flecha fue desequipada porque no es compatible con el arco actual.
            Call WriteLocaleMsg(UserIndex, 2148, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ValidateEquippedArrow_Error:
    Call Logging.TraceError(Err.Number, Err.Description, "InvUsuario.ValidateEquippedArrow", Erl())
End Sub
