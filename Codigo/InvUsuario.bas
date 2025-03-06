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

Public Function get_max_items_inventory(ByVal user_type As e_TipoUsuario) As Integer

    ' Determine inventory slots based on user type
    Select Case user_type

        Case tLeyenda
            get_max_items_inventory = MAX_INVENTORY_SLOTS

        Case tHeroe
            get_max_items_inventory = MAX_USERINVENTORY_HERO_SLOTS

        Case Else
            get_max_items_inventory = MAX_USERINVENTORY_SLOTS

    End Select

End Function

Public Function IsObjecIndextInInventory(ByVal UserIndex As Integer, _
                                         ByVal ObjIndex As Integer) As Boolean

    On Error GoTo IsObjecIndextInInventory_Err

    Debug.Assert UserIndex >= LBound(UserList) And UserIndex <= UBound(UserList)
    ' If no match is found, return False
    IsObjecIndextInInventory = False

    Dim i                 As Integer
    Dim maxItemsInventory As Integer
    Dim currentObjIndex   As Integer

    With UserList(UserIndex)
        maxItemsInventory = get_max_items_inventory(.Stats.tipoUsuario)

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

Public Function get_object_amount_from_inventory(ByVal user_index, _
                                                 ByVal obj_index As Integer) As Integer

    On Error GoTo get_object_amount_from_inventory_Err

    Debug.Assert user_index >= LBound(UserList) And user_index <= UBound(UserList)
    ' If no match is found, return 0
    get_object_amount_from_inventory = 0

    Dim i                 As Integer
    Dim maxItemsInventory As Integer

    With UserList(user_index)
        maxItemsInventory = get_max_items_inventory(.Stats.tipoUsuario)

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

100         For i = 1 To UserList(UserIndex).CurrentInventorySlots
102             ObjIndex = UserList(UserIndex).invent.Object(i).ObjIndex

104             If ObjIndex > 0 Then
106                 If (ObjData(ObjIndex).OBJType <> e_OBJType.otLlaves And ObjData(ObjIndex).OBJType <> e_OBJType.otBarcos And ObjData(ObjIndex).OBJType <> e_OBJType.otMonturas And ObjData(ObjIndex).OBJType <> e_OBJType.OtDonador And ObjData(ObjIndex).OBJType <> e_OBJType.otRunas) Then
108                     TieneObjetosRobables = True
                        Exit Function

                    End If

                End If

110         Next i

        End If

        Exit Function
TieneObjetosRobables_Err:
112     Call TraceError(Err.Number, Err.Description, "InvUsuario.TieneObjetosRobables", Erl)

End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, _
                            ByVal ObjIndex As Integer, _
                            Optional Slot As Byte) As Boolean

        On Error GoTo manejador

        Dim Flag As Boolean

100     If Slot <> 0 Then
102         If UserList(UserIndex).invent.Object(Slot).Equipped Then
104             ClasePuedeUsarItem = True
                Exit Function

            End If

        End If

106     If EsGM(UserIndex) Then
108         ClasePuedeUsarItem = True
            Exit Function

        End If

        Dim i As Integer

110     For i = 1 To NUMCLASES

112         If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
114             ClasePuedeUsarItem = False
                Exit Function

            End If

116     Next i

118     ClasePuedeUsarItem = True
        Exit Function
manejador:
120     LogError ("Error en ClasePuedeUsarItem")

End Function

Function RazaPuedeUsarItem(ByVal UserIndex As Integer, _
                           ByVal ObjIndex As Integer, _
                           Optional Slot As Byte) As Boolean

        On Error GoTo RazaPuedeUsarItem_Err

        Dim Objeto As t_ObjData, i As Long
100     Objeto = ObjData(ObjIndex)

102     If EsGM(UserIndex) Then
104         RazaPuedeUsarItem = True
            Exit Function

        End If

106     For i = 1 To NUMRAZAS

108         If Objeto.RazaProhibida(i) = UserList(UserIndex).raza Then
110             RazaPuedeUsarItem = False
                Exit Function

            End If

112     Next i

        ' Si el objeto no define una raza en particular
114     If Objeto.RazaDrow + Objeto.RazaElfa + Objeto.RazaEnana + Objeto.RazaGnoma + Objeto.RazaHumana + Objeto.RazaOrca = 0 Then
116         RazaPuedeUsarItem = True
        Else ' El objeto esta definido para alguna raza en especial

118         Select Case UserList(UserIndex).raza

                Case e_Raza.Humano
120                 RazaPuedeUsarItem = Objeto.RazaHumana > 0

122             Case e_Raza.Elfo
124                 RazaPuedeUsarItem = Objeto.RazaElfa > 0

126             Case e_Raza.Drow
128                 RazaPuedeUsarItem = Objeto.RazaDrow > 0

130             Case e_Raza.Orco
132                 RazaPuedeUsarItem = Objeto.RazaOrca > 0

134             Case e_Raza.Gnomo
136                 RazaPuedeUsarItem = Objeto.RazaGnoma > 0

138             Case e_Raza.Enano
140                 RazaPuedeUsarItem = Objeto.RazaEnana > 0

            End Select

        End If

        If RazaPuedeUsarItem And Objeto.OBJType = e_OBJType.otArmadura Then
            RazaPuedeUsarItem = ObtenerRopaje(UserIndex, Objeto) <> 0

        End If

        Exit Function
RazaPuedeUsarItem_Err:
142     LogError ("Error en RazaPuedeUsarItem")

End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)

        On Error GoTo QuitarNewbieObj_Err

        Dim j As Integer

        If UserList(UserIndex).CurrentInventorySlots > 0 Then

100         For j = 1 To UserList(UserIndex).CurrentInventorySlots

102             If UserList(UserIndex).invent.Object(j).ObjIndex > 0 Then
104                 If ObjData(UserList(UserIndex).invent.Object(j).ObjIndex).Newbie = 1 Then
106                     Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
108                     Call UpdateUserInv(False, UserIndex, j)

                    End If

                End If

110         Next j

        End If

        ' Eliminar items newbie de la boveda
        For j = 1 To MAX_BANCOINVENTORY_SLOTS

            If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
                If ObjData(UserList(UserIndex).BancoInvent.Object(j).ObjIndex).Newbie = 1 Then
                    UserList(UserIndex).BancoInvent.Object(j).ObjIndex = 0
                    UserList(UserIndex).BancoInvent.Object(j).amount = 0
                    Call UpdateBanUserInv(False, UserIndex, j, "QuitarNewbieObj")

                End If

            End If

        Next j

        'Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
        'Mandamos a la Isla de la Fortuna
        Call WarpUserChar(UserIndex, Renacimiento.Map, Renacimiento.x, Renacimiento.y, True)
        ' Msg671=Has dejado de ser Newbie.
        Call WriteLocaleMsg(UserIndex, "671", e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
QuitarNewbieObj_Err:
144     Call TraceError(Err.Number, Err.Description, "InvUsuario.QuitarNewbieObj", Erl)

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)

        On Error GoTo LimpiarInventario_Err

        Dim j As Integer

        If UserList(UserIndex).CurrentInventorySlots > 0 Then

100         For j = 1 To UserList(UserIndex).CurrentInventorySlots
102             UserList(UserIndex).invent.Object(j).ObjIndex = 0
104             UserList(UserIndex).invent.Object(j).amount = 0
106             UserList(UserIndex).invent.Object(j).Equipped = 0
            Next

        End If

108     UserList(UserIndex).invent.NroItems = 0
110     UserList(UserIndex).invent.ArmourEqpObjIndex = 0
112     UserList(UserIndex).invent.ArmourEqpSlot = 0
114     UserList(UserIndex).invent.WeaponEqpObjIndex = 0
116     UserList(UserIndex).invent.WeaponEqpSlot = 0
118     UserList(UserIndex).invent.HerramientaEqpObjIndex = 0
120     UserList(UserIndex).invent.HerramientaEqpSlot = 0
122     UserList(UserIndex).invent.CascoEqpObjIndex = 0
124     UserList(UserIndex).invent.CascoEqpSlot = 0
126     UserList(UserIndex).invent.EscudoEqpObjIndex = 0
128     UserList(UserIndex).invent.EscudoEqpSlot = 0
130     UserList(UserIndex).invent.DañoMagicoEqpObjIndex = 0
132     UserList(UserIndex).invent.DañoMagicoEqpSlot = 0
134     UserList(UserIndex).invent.ResistenciaEqpObjIndex = 0
136     UserList(UserIndex).invent.ResistenciaEqpSlot = 0
142     UserList(UserIndex).invent.MunicionEqpObjIndex = 0
144     UserList(UserIndex).invent.MunicionEqpSlot = 0
146     UserList(UserIndex).invent.BarcoObjIndex = 0
148     UserList(UserIndex).invent.BarcoSlot = 0
150     UserList(UserIndex).invent.MonturaObjIndex = 0
152     UserList(UserIndex).invent.MonturaSlot = 0
154     UserList(UserIndex).invent.MagicoObjIndex = 0
156     UserList(UserIndex).invent.MagicoSlot = 0
        Exit Sub
LimpiarInventario_Err:
158     Call TraceError(Err.Number, Err.Description, "InvUsuario.LimpiarInventario", Erl)

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

100     With UserList(UserIndex)

            ' GM's (excepto Dioses y Admins) no pueden tirar oro
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
104             Call LogGM(.name, " trató de tirar " & PonerPuntos(Cantidad) & " de oro en " & .pos.Map & "-" & .pos.x & "-" & .pos.y)
                Exit Sub

            End If

            ' Si el usuario tiene ORO, entonces lo tiramos
106         If (Cantidad > 0) And (Cantidad <= .Stats.GLD) Then

                Dim i     As Byte
                Dim MiObj As t_Obj

                'info debug
                Dim loops As Long

116             Do While (Cantidad > 0)

118                 If Cantidad > MAX_INVENTORY_OBJS And .Stats.GLD > MAX_INVENTORY_OBJS Then
120                     MiObj.amount = MAX_INVENTORY_OBJS
122                     Cantidad = Cantidad - MiObj.amount
                    Else
124                     MiObj.amount = Cantidad
126                     Cantidad = Cantidad - MiObj.amount

                    End If

128                 MiObj.ObjIndex = iORO

                    Dim AuxPos As t_WorldPos

130                 If .clase = e_Class.Pirat Then
132                     AuxPos = TirarItemAlPiso(.pos, MiObj, False)
                    Else
134                     AuxPos = TirarItemAlPiso(.pos, MiObj, True)

                    End If

136                 If AuxPos.x <> 0 And AuxPos.y <> 0 Then
138                     .Stats.GLD = .Stats.GLD - MiObj.amount

                    End If

                    'info debug
140                 loops = loops + 1

142                 If loops > 100000 Then 'si entra aca y se cuelga mal el server revisen al tipo porque tiene much oro (NachoP) seguramente es dupero
144                     Call LogError("Se ha superado el limite de iteraciones(100000) permitido en el Sub TirarOro() - posible Nacho P")
                        Exit Sub

                    End If

                Loop

                ' Si es GM, registramos lo q hizo
146             If EsGM(UserIndex) Then
148                 If MiObj.ObjIndex = iORO Then
150                     Call LogGM(.name, "Tiro: " & PonerPuntos(OriginalAmount) & " monedas de oro.")
                    Else
152                     Call LogGM(.name, "Tiro cantidad:" & PonerPuntos(OriginalAmount) & " Objeto:" & ObjData(MiObj.ObjIndex).name)

                    End If

                End If

160             Call WriteUpdateGold(UserIndex)

            End If

        End With

        Exit Sub
ErrHandler:
162     Call TraceError(Err.Number, Err.Description, "InvUsuario.TirarOro", Erl())

End Sub

Public Sub QuitarUserInvItem(ByVal UserIndex As Integer, _
                             ByVal Slot As Byte, _
                             ByVal Cantidad As Integer)

        On Error GoTo QuitarUserInvItem_Err

100     If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub

102     With UserList(UserIndex).invent.Object(Slot)

104         If .amount <= Cantidad And .Equipped = 1 Then
106             Call Desequipar(UserIndex, Slot)

            End If

            'Quita un objeto
108         .amount = .amount - Cantidad

            '¿Quedan mas?
110         If .amount <= 0 Then
112             UserList(UserIndex).invent.NroItems = UserList(UserIndex).invent.NroItems - 1
114             .ObjIndex = 0
116             .amount = 0

            End If

            UserList(UserIndex).flags.ModificoInventario = True

        End With

        If IsValidUserRef(UserList(UserIndex).flags.GMMeSigue) And UserIndex <> UserList(UserIndex).flags.GMMeSigue.ArrayIndex Then
            Call QuitarUserInvItem(UserList(UserIndex).flags.GMMeSigue.ArrayIndex, Slot, Cantidad)

        End If

        Exit Sub
QuitarUserInvItem_Err:
118     Call TraceError(Err.Number, Err.Description, "InvUsuario.QuitarUserInvItem", Erl)

End Sub

Public Sub UpdateUserInv(ByVal UpdateAll As Boolean, _
                         ByVal UserIndex As Integer, _
                         ByVal Slot As Byte)

        On Error GoTo UpdateUserInv_Err

        Dim NullObj As t_UserOBJ
        Dim LoopC   As Byte

        'Actualiza un solo slot
100     If Not UpdateAll And Slot > 0 Then

            'Actualiza el inventario
102         If UserList(UserIndex).invent.Object(Slot).ObjIndex > 0 Then
104             Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).invent.Object(Slot))
            Else
106             Call ChangeUserInv(UserIndex, Slot, NullObj)

            End If

            UserList(UserIndex).flags.ModificoInventario = True
        Else

            'Actualiza todos los slots
            If UserList(UserIndex).CurrentInventorySlots > 0 Then

108             For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots

                    'Actualiza el inventario
110                 If UserList(UserIndex).invent.Object(LoopC).ObjIndex > 0 Then
112                     Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).invent.Object(LoopC))
                    Else
114                     Call ChangeUserInv(UserIndex, LoopC, NullObj)

                    End If

116             Next LoopC

            End If

        End If

        Exit Sub
UpdateUserInv_Err:
118     Call TraceError(Err.Number, Err.Description, "InvUsuario.UpdateUserInv", Erl)

End Sub

Sub DropObj(ByVal UserIndex As Integer, _
            ByVal Slot As Byte, _
            ByVal num As Integer, _
            ByVal Map As Integer, _
            ByVal x As Integer, _
            ByVal y As Integer)

        On Error GoTo DropObj_Err

        Dim obj As t_Obj

100     If num > 0 Then

102         With UserList(UserIndex)

104             If num > .invent.Object(Slot).amount Then
106                 num = .invent.Object(Slot).amount

                End If

108             obj.ObjIndex = .invent.Object(Slot).ObjIndex
110             obj.amount = num

                If Not CustomScenarios.UserCanDropItem(UserIndex, Slot, Map, x, y) Then
                    Exit Sub

                End If

112             If ObjData(obj.ObjIndex).Destruye = 0 Then

                    Dim Suma As Long
                    Suma = num + MapData(.pos.Map, x, y).ObjInfo.amount

                    'Check objeto en el suelo
114                 If MapData(.pos.Map, x, y).ObjInfo.ObjIndex = 0 Or (MapData(.pos.Map, x, y).ObjInfo.ObjIndex = obj.ObjIndex And Suma <= MAX_INVENTORY_OBJS) Then
116                     If Suma > MAX_INVENTORY_OBJS Then
118                         num = MAX_INVENTORY_OBJS - MapData(.pos.Map, x, y).ObjInfo.amount

                        End If

                        ' Si sos Admin, Dios o Usuario, crea el objeto en el piso.
120                     If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Admin Or e_PlayerType.Dios)) <> 0 Then
                            ' Tiramos el item al piso
122                         Call MakeObj(obj, Map, x, y)

                        End If

                        Call CustomScenarios.UserDropItem(UserIndex, Slot, Map, x, y)
124                     Call QuitarUserInvItem(UserIndex, Slot, num)
126                     Call UpdateUserInv(False, UserIndex, Slot)

                        If .flags.jugando_captura = 1 Then
                            If Not InstanciaCaptura Is Nothing Then
                                Call InstanciaCaptura.tiraBandera(UserIndex, obj.ObjIndex)

                            End If

                        End If

128                     If Not .flags.Privilegios And e_PlayerType.user Then
                            If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) <> 0 Then
130                             Call LogGM(.name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).name)

                            End If

                        End If

                    Else
132                     Call WriteLocaleMsg(UserIndex, "262", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
134                 Call QuitarUserInvItem(UserIndex, Slot, num)
136                 Call UpdateUserInv(False, UserIndex, Slot)

                End If

            End With

        End If

        Exit Sub
DropObj_Err:
138     Call TraceError(Err.Number, Err.Description, "InvUsuario.DropObj", Erl)

End Sub

Sub EraseObj(ByVal num As Integer, _
             ByVal Map As Integer, _
             ByVal x As Integer, _
             ByVal y As Integer)

        On Error GoTo EraseObj_Err

        Dim Rango As Byte
100     MapData(Map, x, y).ObjInfo.amount = MapData(Map, x, y).ObjInfo.amount - num

102     If MapData(Map, x, y).ObjInfo.amount <= 0 Then
108         MapData(Map, x, y).ObjInfo.ObjIndex = 0
110         MapData(Map, x, y).ObjInfo.amount = 0
112         Call modSendData.SendToAreaByPos(Map, x, y, PrepareMessageObjectDelete(x, y))

        End If

        Exit Sub
EraseObj_Err:
114     Call TraceError(Err.Number, Err.Description, "InvUsuario.EraseObj", Erl)

End Sub

Sub MakeObj(ByRef obj As t_Obj, _
            ByVal Map As Integer, _
            ByVal x As Integer, _
            ByVal y As Integer, _
            Optional ByVal Limpiar As Boolean = True)

        On Error GoTo MakeObj_Err

        Dim Color As Long
        Dim Rango As Byte

100     If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then
102         If MapData(Map, x, y).ObjInfo.ObjIndex = obj.ObjIndex Then
104             MapData(Map, x, y).ObjInfo.amount = MapData(Map, x, y).ObjInfo.amount + obj.amount
            Else
110             MapData(Map, x, y).ObjInfo.ObjIndex = obj.ObjIndex

112             If ObjData(obj.ObjIndex).VidaUtil <> 0 Then
114                 MapData(Map, x, y).ObjInfo.amount = ObjData(obj.ObjIndex).VidaUtil
                Else
116                 MapData(Map, x, y).ObjInfo.amount = obj.amount

                End If

            End If

118         Call modSendData.SendToAreaByPos(Map, x, y, PrepareMessageObjectCreate(obj.ObjIndex, MapData(Map, x, y).ObjInfo.amount, x, y))

        End If

        Exit Sub
MakeObj_Err:
120     Call TraceError(Err.Number, Err.Description, "InvUsuario.MakeObj", Erl)

End Sub

Function GetSlotForItemInInventory(ByVal UserIndex As Integer, _
                                   ByRef MyObject As t_Obj) As Integer

        On Error GoTo GetSlotForItemInInventory_Err

        GetSlotForItemInInventory = -1

100     Dim i As Integer

102     For i = 1 To UserList(UserIndex).CurrentInventorySlots

104         If UserList(UserIndex).invent.Object(i).ObjIndex = 0 And GetSlotForItemInInventory = -1 Then
106             GetSlotForItemInInventory = i 'we found a valid place but keep looking in case we can stack
108         ElseIf UserList(UserIndex).invent.Object(i).ObjIndex = MyObject.ObjIndex And UserList(UserIndex).invent.Object(i).amount + MyObject.amount <= MAX_INVENTORY_OBJS Then
110             GetSlotForItemInInventory = i 'we can stack the item, let use this slot
112             Exit Function

            End If

        Next i

        Exit Function
GetSlotForItemInInventory_Err:
        Call TraceError(Err.Number, Err.Description, "InvUsuario.GetSlotForItemInInventory", Erl)

End Function

Function GetSlotInInventory(ByVal UserIndex As Integer, _
                            ByVal ObjIndex As Integer) As Integer

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

Function MeterItemEnInventario(ByVal UserIndex As Integer, _
                               ByRef MiObj As t_Obj) As Boolean

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
100     Slot = GetSlotForItemInInventory(UserIndex, MiObj)

        If Slot <= 0 Then
118         Call WriteLocaleMsg(UserIndex, MsgInventoryIsFull, e_FontTypeNames.FONTTYPE_FIGHT)
120         MeterItemEnInventario = False
            Exit Function

        End If

        If UserList(UserIndex).invent.Object(Slot).ObjIndex = 0 Then
            UserList(UserIndex).invent.NroItems = UserList(UserIndex).invent.NroItems + 1

        End If

        'Mete el objeto
124     If UserList(UserIndex).invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
126         UserList(UserIndex).invent.Object(Slot).ObjIndex = MiObj.ObjIndex
128         UserList(UserIndex).invent.Object(Slot).amount = UserList(UserIndex).invent.Object(Slot).amount + MiObj.amount
        Else
130         UserList(UserIndex).invent.Object(Slot).amount = MAX_INVENTORY_OBJS

        End If

132     Call UpdateUserInv(False, UserIndex, Slot)
134     MeterItemEnInventario = True
        UserList(UserIndex).flags.ModificoInventario = True
        Exit Function
MeterItemEnInventario_Err:
        Call TraceError(Err.Number, Err.Description, "InvUsuario.MeterItemEnInventario", Erl)

End Function

Function HayLugarEnInventario(ByVal UserIndex As Integer, _
                              ByVal TargetItemIndex As Integer, _
                              ByVal ItemCount) As Boolean

        On Error GoTo HayLugarEnInventario_err

        Dim x    As Integer
        Dim y    As Integer
        Dim Slot As Byte
100     Slot = 1

102     Do Until UserList(UserIndex).invent.Object(Slot).ObjIndex = 0 Or (UserList(UserIndex).invent.Object(Slot).ObjIndex = TargetItemIndex And UserList(UserIndex).invent.Object(Slot).amount + ItemCount < 10000)
104         Slot = Slot + 1

106         If Slot > UserList(UserIndex).CurrentInventorySlots Then
108             HayLugarEnInventario = False
                Exit Function

            End If

        Loop
110     HayLugarEnInventario = True
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
100     If MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).ObjInfo.ObjIndex > 0 Then

            '¿Esta permitido agarrar este obj?
102         If ObjData(MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).ObjInfo.ObjIndex).Agarrable <> 1 Then
104             If UserList(UserIndex).flags.Montado = 1 Then
106                 ' Msg672=Debes descender de tu montura para agarrar objetos del suelo.
                    Call WriteLocaleMsg(UserIndex, "672", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

                If Not UserCanPickUpItem(UserIndex) Then
                    Exit Sub

                End If

108             x = UserList(UserIndex).pos.x
110             y = UserList(UserIndex).pos.y

                If UserList(UserIndex).flags.jugando_captura = 1 Then
                    If Not InstanciaCaptura Is Nothing Then
                        If Not InstanciaCaptura.tomaBandera(UserIndex, MapData(UserList(UserIndex).pos.Map, x, y).ObjInfo.ObjIndex) Then
                            Exit Sub

                        End If

                    End If

                End If

112             obj = ObjData(MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).ObjInfo.ObjIndex)
114             MiObj.amount = MapData(UserList(UserIndex).pos.Map, x, y).ObjInfo.amount
116             MiObj.ObjIndex = MapData(UserList(UserIndex).pos.Map, x, y).ObjInfo.ObjIndex

118             If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    'Call WriteConsoleMsg(UserIndex, "No puedo cargar mas objetos.", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    'Quitamos el objeto
120                 Call EraseObj(MapData(UserList(UserIndex).pos.Map, x, y).ObjInfo.amount, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y)

122                 If Not UserList(UserIndex).flags.Privilegios And e_PlayerType.user Then Call LogGM(UserList(UserIndex).name, "Agarro:" & MiObj.amount & " Objeto:" & ObjData(MiObj.ObjIndex).name)
                    Call UserDidPickupItem(UserIndex, MiObj.ObjIndex)

                    If UserList(UserIndex).flags.jugando_captura = 1 Then
                        If Not InstanciaCaptura Is Nothing Then
                            Call InstanciaCaptura.quitarBandera(UserIndex, MiObj.ObjIndex)

                        End If

                    End If

124                 If BusquedaTesoroActiva Then
126                     If UserList(UserIndex).pos.Map = TesoroNumMapa And UserList(UserIndex).pos.x = TesoroX And UserList(UserIndex).pos.y = TesoroY Then
128                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> " & UserList(UserIndex).name & " encontro el tesoro ¡Felicitaciones!", e_FontTypeNames.FONTTYPE_TALK))
130                         BusquedaTesoroActiva = False

                        End If

                    End If

132                 If BusquedaRegaloActiva Then
134                     If UserList(UserIndex).pos.Map = RegaloNumMapa And UserList(UserIndex).pos.x = RegaloX And UserList(UserIndex).pos.y = RegaloY Then
136                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> " & UserList(UserIndex).name & " fue el valiente que encontro el gran item magico ¡Felicitaciones!", e_FontTypeNames.FONTTYPE_TALK))
138                         BusquedaRegaloActiva = False

                        End If

                    End If

                End If

            End If

        Else

144         If Not UserList(UserIndex).flags.UltimoMensaje = 261 Then
146             Call WriteLocaleMsg(UserIndex, "261", e_FontTypeNames.FONTTYPE_INFO)
148             UserList(UserIndex).flags.UltimoMensaje = 261

            End If

        End If

        Exit Sub
PickObj_Err:
150     Call TraceError(Err.Number, Err.Description, "InvUsuario.PickObj", Erl)

End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)

        On Error GoTo Desequipar_Err

        'Desequipa el item slot del inventario
        Dim obj As t_ObjData

100     If (Slot < LBound(UserList(UserIndex).invent.Object)) Or (Slot > UBound(UserList(UserIndex).invent.Object)) Then
            Exit Sub
102     ElseIf UserList(UserIndex).invent.Object(Slot).ObjIndex = 0 Then
            Exit Sub

        End If

104     obj = ObjData(UserList(UserIndex).invent.Object(Slot).ObjIndex)

106     Select Case obj.OBJType

            Case e_OBJType.otWeapon
108             UserList(UserIndex).invent.Object(Slot).Equipped = 0
110             UserList(UserIndex).invent.WeaponEqpObjIndex = 0
112             UserList(UserIndex).invent.WeaponEqpSlot = 0
114             UserList(UserIndex).Char.Arma_Aura = ""
116             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 1))
118             UserList(UserIndex).Char.WeaponAnim = NingunArma

120             If UserList(UserIndex).flags.Montado = 0 Then
122                 Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

                End If

124             If obj.MagicDamageBonus > 0 Then
126                 Call WriteUpdateDM(UserIndex)

                End If

128         Case e_OBJType.otFlechas
130             UserList(UserIndex).invent.Object(Slot).Equipped = 0
132             UserList(UserIndex).invent.MunicionEqpObjIndex = 0
134             UserList(UserIndex).invent.MunicionEqpSlot = 0

                ' Case e_OBJType.otAnillos
                '    UserList(UserIndex).Invent.Object(slot).Equipped = 0
                '    UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
                ' UserList(UserIndex).Invent.AnilloEqpSlot = 0
136         Case e_OBJType.otHerramientas

137             If UserList(UserIndex).flags.PescandoEspecial = False Then
138                 UserList(UserIndex).invent.Object(Slot).Equipped = 0
140                 UserList(UserIndex).invent.HerramientaEqpObjIndex = 0
142                 UserList(UserIndex).invent.HerramientaEqpSlot = 0

144                 If UserList(UserIndex).flags.UsandoMacro = True Then
146                     Call WriteMacroTrabajoToggle(UserIndex, False)

                    End If

148                 UserList(UserIndex).Char.WeaponAnim = NingunArma

150                 If UserList(UserIndex).flags.Montado = 0 Then
152                     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

                    End If

                End If

154         Case e_OBJType.otMagicos

156             Select Case obj.EfectoMagico

                    Case e_MagicItemEffect.eModifyAttributes

                        If obj.QueAtributo <> 0 Then
162                         UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
164                         UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) - obj.CuantoAumento
                            ' UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
166                         Call WriteFYA(UserIndex)

                        End If

168                 Case e_MagicItemEffect.eModifySkills

                        If obj.Que_Skill <> 0 Then
170                         UserList(UserIndex).Stats.UserSkills(obj.Que_Skill) = UserList(UserIndex).Stats.UserSkills(obj.Que_Skill) - obj.CuantoAumento

                        End If

172                 Case e_MagicItemEffect.eRegenerateHealth
174                     UserList(UserIndex).flags.RegeneracionHP = 0

176                 Case e_MagicItemEffect.eRegenerateMana
178                     UserList(UserIndex).flags.RegeneracionMana = 0

180                 Case e_MagicItemEffect.eIncreaseDamageToNpc
182                     UserList(UserIndex).Stats.MaxHit = UserList(UserIndex).Stats.MaxHit - obj.CuantoAumento
184                     UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT - obj.CuantoAumento

188                 Case e_MagicItemEffect.eInmunityToNpcMagic 'Orbe ignea
190                     UserList(UserIndex).flags.NoMagiaEfecto = 0

192                 Case e_MagicItemEffect.eIncinerate
194                     UserList(UserIndex).flags.incinera = 0

196                 Case e_MagicItemEffect.eParalize
198                     UserList(UserIndex).flags.Paraliza = 0

200                 Case e_MagicItemEffect.eProtectedResources

202                     If UserList(UserIndex).flags.Muerto = 0 Then
                            UserList(UserIndex).Char.CartAnim = NoCart
203                         Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

                        End If

206                 Case e_MagicItemEffect.eProtectedInventory
208                     UserList(UserIndex).flags.PendienteDelSacrificio = 0

210                 Case e_MagicItemEffect.ePreventMagicWords
212                     UserList(UserIndex).flags.NoPalabrasMagicas = 0

214                 Case e_MagicItemEffect.ePreventInvisibleDetection
216                     UserList(UserIndex).flags.NoDetectable = 0

218                 Case e_MagicItemEffect.eIncreaseLearningSkills
220                     UserList(UserIndex).flags.PendienteDelExperto = 0

222                 Case e_MagicItemEffect.ePoison
224                     UserList(UserIndex).flags.Envenena = 0

226                 Case e_MagicItemEffect.eRingOfShadows
228                     UserList(UserIndex).flags.AnilloOcultismo = 0

                    Case e_MagicItemEffect.eTalkToDead
                        Call UnsetMask(UserList(UserIndex).flags.StatusMask, e_StatusMask.eTalkToDead)
                        ' Msg673=Dejas el mundo de los muertos, ya no podrás comunicarte con ellos.
                        Call WriteLocaleMsg(UserIndex, "673", e_FontTypeNames.FONTTYPE_WARNING)
                        Call SendData(SendTarget.ToPCDeadAreaButIndex, UserIndex, PrepareMessageCharacterRemove(4, UserList(UserIndex).Char.charindex, False, True))

                End Select

230             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 5))
232             UserList(UserIndex).Char.Otra_Aura = 0
234             UserList(UserIndex).invent.Object(Slot).Equipped = 0
236             UserList(UserIndex).invent.MagicoObjIndex = 0
238             UserList(UserIndex).invent.MagicoSlot = 0

256         Case e_OBJType.otArmadura
258             UserList(UserIndex).invent.Object(Slot).Equipped = 0
260             UserList(UserIndex).invent.ArmourEqpObjIndex = 0
262             UserList(UserIndex).invent.ArmourEqpSlot = 0

264             If UserList(UserIndex).flags.Navegando = 0 Then
266                 If UserList(UserIndex).flags.Montado = 0 Then
                        Call SetNakedBody(UserList(UserIndex))
270                     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

                    End If

                End If

272             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 2))
274             UserList(UserIndex).Char.Body_Aura = 0

276             If obj.ResistenciaMagica > 0 Then
278                 Call WriteUpdateRM(UserIndex)

                End If

280         Case e_OBJType.otCasco
282             UserList(UserIndex).invent.Object(Slot).Equipped = 0
284             UserList(UserIndex).invent.CascoEqpObjIndex = 0
286             UserList(UserIndex).invent.CascoEqpSlot = 0
288             UserList(UserIndex).Char.Head_Aura = 0
290             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 4))
292             UserList(UserIndex).Char.CascoAnim = NingunCasco
294             Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

296             If obj.ResistenciaMagica > 0 Then
298                 Call WriteUpdateRM(UserIndex)

                End If

300         Case e_OBJType.otEscudo
302             UserList(UserIndex).invent.Object(Slot).Equipped = 0
304             UserList(UserIndex).invent.EscudoEqpObjIndex = 0
306             UserList(UserIndex).invent.EscudoEqpSlot = 0
308             UserList(UserIndex).Char.Escudo_Aura = 0
310             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 3))
312             UserList(UserIndex).Char.ShieldAnim = NingunEscudo

314             If UserList(UserIndex).flags.Montado = 0 Then
316                 Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

                End If

318             If obj.ResistenciaMagica > 0 Then
320                 Call WriteUpdateRM(UserIndex)

                End If

322         Case e_OBJType.otDañoMagico
324             UserList(UserIndex).invent.Object(Slot).Equipped = 0
326             UserList(UserIndex).invent.DañoMagicoEqpObjIndex = 0
328             UserList(UserIndex).invent.DañoMagicoEqpSlot = 0
330             UserList(UserIndex).Char.DM_Aura = 0
332             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 6))
334             Call WriteUpdateDM(UserIndex)

336         Case e_OBJType.otResistencia
338             UserList(UserIndex).invent.Object(Slot).Equipped = 0
340             UserList(UserIndex).invent.ResistenciaEqpObjIndex = 0
342             UserList(UserIndex).invent.ResistenciaEqpSlot = 0
344             UserList(UserIndex).Char.RM_Aura = 0
346             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 7))
348             Call WriteUpdateRM(UserIndex)

        End Select

350     Call UpdateUserInv(False, UserIndex, Slot)
        Exit Sub
Desequipar_Err:
352     Call TraceError(Err.Number, Err.Description, "InvUsuario.Desequipar", Erl)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, _
                           ByVal ObjIndex As Integer) As Boolean

        On Error GoTo ErrHandler

100     If EsGM(UserIndex) Then
102         SexoPuedeUsarItem = True
            Exit Function

        End If

104     If ObjData(ObjIndex).Mujer = 1 Then
106         SexoPuedeUsarItem = UserList(UserIndex).genero <> e_Genero.Hombre
108     ElseIf ObjData(ObjIndex).Hombre = 1 Then
110         SexoPuedeUsarItem = UserList(UserIndex).genero <> e_Genero.Mujer
        Else
112         SexoPuedeUsarItem = True

        End If

        Exit Function
ErrHandler:
114     Call LogError("SexoPuedeUsarItem")

End Function

Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, _
                              ByVal ObjIndex As Integer) As Boolean

        On Error GoTo FaccionPuedeUsarItem_Err

100     If EsGM(UserIndex) Then
102         FaccionPuedeUsarItem = True
            Exit Function

        End If

104     If ObjIndex < 1 Then Exit Function
106     If ObjData(ObjIndex).Real = 1 Then
108         If Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
110             FaccionPuedeUsarItem = esArmada(UserIndex)
            Else
112             FaccionPuedeUsarItem = False

            End If

114     ElseIf ObjData(ObjIndex).Caos = 1 Then

116         If Status(UserIndex) = e_Facciones.Caos Or Status(UserIndex) = e_Facciones.concilio Then
118             FaccionPuedeUsarItem = esCaos(UserIndex)
            Else
120             FaccionPuedeUsarItem = False

            End If

        Else
122         FaccionPuedeUsarItem = True

        End If

        Exit Function
FaccionPuedeUsarItem_Err:
124     Call TraceError(Err.Number, Err.Description, "InvUsuario.FaccionPuedeUsarItem", Erl)

End Function

Function JerarquiaPuedeUsarItem(ByVal UserIndex As Integer, _
                                ByVal ObjIndex As Integer) As Boolean

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

100     With UserList(UserIndex)

            If .invent.BarcoObjIndex <= 0 Or .invent.BarcoObjIndex > UBound(ObjData) Then Exit Sub
102         Barco = ObjData(.invent.BarcoObjIndex)

104         If .flags.Muerto = 1 Then
106             If Barco.Ropaje = iTraje Or Barco.Ropaje = iTrajeAltoNw Or Barco.Ropaje = iTrajeBajoNw Then
                    ' No tenemos la cabeza copada que va con iRopaBuceoMuerto,
                    ' asique asignamos el casper directamente caminando sobre el agua.
108                 .Char.body = iCuerpoMuerto 'iRopaBuceoMuerto
110                 .Char.head = iCabezaMuerto
                ElseIf Barco.Ropaje = iTrajeAltoNw Then
                ElseIf Barco.Ropaje = iTrajeBajoNw Then
                Else
112                 .Char.body = iFragataFantasmal
114                 .Char.head = 0

                End If

            Else ' Esta vivo

116             If Barco.Ropaje = iTraje Then
118                 .Char.body = iTraje
120                 .Char.head = .OrigChar.head

122                 If .invent.CascoEqpObjIndex > 0 Then
124                     .Char.CascoAnim = ObjData(.invent.CascoEqpObjIndex).CascoAnim

                    End If

                ElseIf Barco.Ropaje = iTrajeAltoNw Then
                    .Char.body = iTrajeAltoNw
                    .Char.head = .OrigChar.head

                    If .invent.CascoEqpObjIndex > 0 Then
                        .Char.CascoAnim = ObjData(.invent.CascoEqpObjIndex).CascoAnim

                    End If

                ElseIf Barco.Ropaje = iTrajeBajoNw Then
                    .Char.body = iTrajeBajoNw
                    .Char.head = .OrigChar.head

                    If .invent.CascoEqpObjIndex > 0 Then
                        .Char.CascoAnim = ObjData(.invent.CascoEqpObjIndex).CascoAnim

                    End If

                Else
126                 .Char.head = 0
128                 .Char.CascoAnim = NingunCasco

                End If

130             If .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Then
132                 If Barco.Ropaje = iBarca Then .Char.body = iBarcaArmada
134                 If Barco.Ropaje = iGalera Then .Char.body = iGaleraArmada
136                 If Barco.Ropaje = iGaleon Then .Char.body = iGaleonArmada
138             ElseIf .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then

140                 If Barco.Ropaje = iBarca Then .Char.body = iBarcaCaos
142                 If Barco.Ropaje = iGalera Then .Char.body = iGaleraCaos
144                 If Barco.Ropaje = iGaleon Then .Char.body = iGaleonCaos
                Else

146                 If Barco.Ropaje = iBarca Then .Char.body = IIf(.Faccion.Status = 0, iBarcaCrimi, iBarcaCiuda)
148                 If Barco.Ropaje = iGalera Then .Char.body = IIf(.Faccion.Status = 0, iGaleraCrimi, iGaleraCiuda)
150                 If Barco.Ropaje = iGaleon Then .Char.body = IIf(.Faccion.Status = 0, iGaleonCrimi, iGaleonCiuda)

                End If

            End If

152         .Char.ShieldAnim = NingunEscudo
154         .Char.WeaponAnim = NingunArma
            Call WriteNavigateToggle(UserIndex, .flags.Navegando)
156         Call WriteNadarToggle(UserIndex, (Barco.Ropaje = iTraje Or Barco.Ropaje = iTrajeAltoNw Or Barco.Ropaje = iTrajeBajoNw), (Barco.Ropaje = iTrajeAltoNw Or Barco.Ropaje = iTrajeBajoNw))
            Call ActualizarVelocidadDeUsuario(UserIndex)

        End With

        Exit Sub
EquiparBarco_Err:
158     Call TraceError(Err.Number, Err.Description, "InvUsuario.EquiparBarco", Erl)

End Sub

'Equipa un item del inventario
Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

        On Error GoTo ErrHandler

        Dim obj       As t_ObjData
        Dim ObjIndex  As Integer
        Dim errordesc As String
100     ObjIndex = UserList(UserIndex).invent.Object(Slot).ObjIndex
102     obj = ObjData(ObjIndex)

104     If PuedeUsarObjeto(UserIndex, ObjIndex, True) > 0 Then
            Exit Sub

        End If

106     With UserList(UserIndex)

108         If .flags.Muerto = 1 Then
                'Msg77=¡¡Estás muerto!!.
110             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

112         Select Case obj.OBJType

                Case e_OBJType.otWeapon
114                 errordesc = "Arma"

                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eWeapon) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    'Si esta equipado lo quita
116                 If .invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
118                     Call Desequipar(UserIndex, Slot)
                        'Animacion por defecto
120                     .Char.WeaponAnim = NingunArma

122                     If .flags.Montado = 0 Then
124                         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

                        End If

                        Exit Sub

                    End If

                    'Quitamos el elemento anterior
126                 If .invent.WeaponEqpObjIndex > 0 Then
128                     Call Desequipar(UserIndex, .invent.WeaponEqpSlot)

                    End If

130                 If .invent.HerramientaEqpObjIndex > 0 Then
132                     Call Desequipar(UserIndex, .invent.HerramientaEqpSlot)

                    End If

138                 .invent.Object(Slot).Equipped = 1
140                 .invent.WeaponEqpObjIndex = .invent.Object(Slot).ObjIndex
142                 .invent.WeaponEqpSlot = Slot

154                 If obj.DosManos = 1 Then
156                     If .invent.EscudoEqpObjIndex > 0 Then
158                         Call Desequipar(UserIndex, .invent.EscudoEqpSlot)
160                         ' Msg674=No puedes usar armas dos manos si tienes un escudo equipado. Tu escudo fue desequipado.
                            Call WriteLocaleMsg(UserIndex, "674", e_FontTypeNames.FONTTYPE_INFOIAO)

                        End If

                    End If

                    'Sonido
162                 If obj.SndAura = 0 Then
164                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .pos.x, .pos.y))
                    Else
166                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .pos.x, .pos.y))

                    End If

168                 If Len(obj.CreaGRH) <> 0 Then
170                     .Char.Arma_Aura = obj.CreaGRH
172                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Arma_Aura, False, 1))

                    End If

174                 If obj.MagicDamageBonus > 0 Then
176                     Call WriteUpdateDM(UserIndex)

                    End If

178                 If .flags.Montado = 0 Then
180                     If .flags.Navegando = 0 Then
182                         .Char.WeaponAnim = obj.WeaponAnim
184                         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

                        End If

                    End If

186             Case e_OBJType.otHerramientas

                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eTool) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    'Si esta equipado lo quita
188                 If .invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
190                     Call Desequipar(UserIndex, Slot)
                        Exit Sub

                    End If

                    'Quitamos el elemento anterior
192                 If .invent.HerramientaEqpObjIndex > 0 Then
194                     Call Desequipar(UserIndex, .invent.HerramientaEqpSlot)

                    End If

196                 If .invent.WeaponEqpObjIndex > 0 Then
198                     Call Desequipar(UserIndex, .invent.WeaponEqpSlot)

                    End If

200                 .invent.Object(Slot).Equipped = 1
202                 .invent.HerramientaEqpObjIndex = ObjIndex
204                 .invent.HerramientaEqpSlot = Slot

206                 If .flags.Montado = 0 Then
208                     If .flags.Navegando = 0 Then
210                         .Char.WeaponAnim = obj.WeaponAnim
212                         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

                        End If

                    End If

214             Case e_OBJType.otMagicos
216                 errordesc = "Magico"

                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eMagicItem) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    'Si esta equipado lo quita
218                 If .invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
220                     Call Desequipar(UserIndex, Slot)
                        Exit Sub

                    End If

                    'Quitamos el elemento anterior
222                 If .invent.MagicoObjIndex > 0 Then
224                     Call Desequipar(UserIndex, .invent.MagicoSlot)

                    End If

226                 .invent.Object(Slot).Equipped = 1
228                 .invent.MagicoObjIndex = .invent.Object(Slot).ObjIndex
230                 .invent.MagicoSlot = Slot

232                 Select Case obj.EfectoMagico

                        Case e_MagicItemEffect.eModifyAttributes 'Modif la fuerza, agilidad, carisma, etc
238                         .Stats.UserAtributosBackUP(obj.QueAtributo) = .Stats.UserAtributosBackUP(obj.QueAtributo) + obj.CuantoAumento
240                         .Stats.UserAtributos(obj.QueAtributo) = MinimoInt(.Stats.UserAtributos(obj.QueAtributo) + obj.CuantoAumento, .Stats.UserAtributosBackUP(obj.QueAtributo) * 2)
242                         Call WriteFYA(UserIndex)

244                     Case e_MagicItemEffect.eModifySkills
246                         .Stats.UserSkills(obj.Que_Skill) = .Stats.UserSkills(obj.Que_Skill) + obj.CuantoAumento

248                     Case e_MagicItemEffect.eRegenerateHealth
250                         .flags.RegeneracionHP = 1

252                     Case e_MagicItemEffect.eRegenerateMana
254                         .flags.RegeneracionMana = 1

256                     Case e_MagicItemEffect.eIncreaseDamageToNpc
258                         .Stats.MaxHit = .Stats.MaxHit + obj.CuantoAumento
260                         .Stats.MinHIT = .Stats.MinHIT + obj.CuantoAumento

262                     Case e_MagicItemEffect.eInmunityToNpcMagic
264                         .flags.NoMagiaEfecto = 1

266                     Case e_MagicItemEffect.eIncinerate
268                         .flags.incinera = 1

270                     Case e_MagicItemEffect.eParalize
272                         .flags.Paraliza = 1

274                     Case e_MagicItemEffect.eProtectedResources

                            If .flags.Navegando = 0 And .flags.Montado = 0 Then
                                .Char.CartAnim = obj.Ropaje
                                Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

                            End If

280                     Case e_MagicItemEffect.eProtectedInventory
282                         .flags.PendienteDelSacrificio = 1

284                     Case e_MagicItemEffect.ePreventMagicWords
286                         .flags.NoPalabrasMagicas = 1

288                     Case e_MagicItemEffect.ePreventInvisibleDetection
290                         .flags.NoDetectable = 1

292                     Case e_MagicItemEffect.eIncreaseLearningSkills
294                         .flags.PendienteDelExperto = 1

296                     Case e_MagicItemEffect.ePoison
298                         .flags.Envenena = 1

300                     Case e_MagicItemEffect.eRingOfShadows
302                         .flags.AnilloOcultismo = 1

                        Case e_MagicItemEffect.eTalkToDead
                            Call SetMask(.flags.StatusMask, e_StatusMask.eTalkToDead)
                            ' Msg675=Entras al mundo de los muertos, ahora podrás comunicarte con ellos.
                            Call WriteLocaleMsg(UserIndex, "675", e_FontTypeNames.FONTTYPE_WARNING)
                            Call CheckUpdateNeededUser(UserIndex, USER_NUEVO, True, 1)

                    End Select

                    'Sonido
304                 If obj.SndAura <> 0 Then
306                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .pos.x, .pos.y))

                    End If

308                 If Len(obj.CreaGRH) <> 0 Then
310                     .Char.Otra_Aura = obj.CreaGRH
312                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Otra_Aura, False, 5))

                    End If

354             Case e_OBJType.otFlechas

                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eAmunition) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    'Si esta equipado lo quita
356                 If .invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
358                     Call Desequipar(UserIndex, Slot)
                        Exit Sub

                    End If

                    'Quitamos el elemento anterior
360                 If .invent.MunicionEqpObjIndex > 0 Then
362                     Call Desequipar(UserIndex, .invent.MunicionEqpSlot)

                    End If

364                 .invent.Object(Slot).Equipped = 1
366                 .invent.MunicionEqpObjIndex = .invent.Object(Slot).ObjIndex
368                 .invent.MunicionEqpSlot = Slot

370             Case e_OBJType.otArmadura

                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eArmor) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    Dim Ropaje As Integer
                    Ropaje = ObtenerRopaje(UserIndex, obj)

372                 If Ropaje = 0 Then
374                     ' Msg676=Hay un error con este objeto. Infórmale a un administrador.
                        Call WriteLocaleMsg(UserIndex, "676", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    'Si esta equipado lo quita
376                 If .invent.Object(Slot).Equipped Then
378                     Call Desequipar(UserIndex, Slot)

380                     If .flags.Navegando = 0 And .flags.Montado = 0 Then
                            Call SetNakedBody(UserList(UserIndex))
384                         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                        Else
386                         .flags.Desnudo = 1

                        End If

                        Exit Sub

                    End If

                    'Quita el anterior
388                 If .invent.ArmourEqpObjIndex > 0 Then
390                     errordesc = "Armadura 2"
392                     Call Desequipar(UserIndex, .invent.ArmourEqpSlot)
394                     errordesc = "Armadura 3"

                    End If

                    'Lo equipa
396                 If Len(obj.CreaGRH) <> 0 Then
398                     .Char.Body_Aura = obj.CreaGRH
400                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Body_Aura, False, 2))

                    End If

402                 .invent.Object(Slot).Equipped = 1
404                 .invent.ArmourEqpObjIndex = .invent.Object(Slot).ObjIndex
406                 .invent.ArmourEqpSlot = Slot

408                 If .flags.Montado = 0 And .flags.Navegando = 0 Then
410                     .Char.body = Ropaje
412                     Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

                    End If

414                 .flags.Desnudo = 0

416                 If obj.ResistenciaMagica > 0 Then
418                     Call WriteUpdateRM(UserIndex)

                    End If

420             Case e_OBJType.otCasco

                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eHelm) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    'Si esta equipado lo quita
422                 If .invent.Object(Slot).Equipped Then
424                     Call Desequipar(UserIndex, Slot)
426                     .Char.CascoAnim = NingunCasco
428                     Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                        Exit Sub

                    End If

                    'Quita el anterior
430                 If .invent.CascoEqpObjIndex > 0 Then
432                     Call Desequipar(UserIndex, .invent.CascoEqpSlot)

                    End If

434                 errordesc = "Casco"

                    'Lo equipa
436                 If Len(obj.CreaGRH) <> 0 Then
438                     .Char.Head_Aura = obj.CreaGRH
440                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Head_Aura, False, 4))

                    End If

442                 .invent.Object(Slot).Equipped = 1
444                 .invent.CascoEqpObjIndex = .invent.Object(Slot).ObjIndex
446                 .invent.CascoEqpSlot = Slot

448                 If .flags.Navegando = 0 Then
450                     .Char.CascoAnim = obj.CascoAnim
452                     Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)

                    End If

454                 If obj.ResistenciaMagica > 0 Then
456                     Call WriteUpdateRM(UserIndex)

                    End If

458             Case e_OBJType.otEscudo

                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eShiled) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    'Si esta equipado lo quita
460                 If .invent.Object(Slot).Equipped Then
462                     Call Desequipar(UserIndex, Slot)
464                     .Char.ShieldAnim = NingunEscudo

466                     If .flags.Montado = 0 And .flags.Navegando = 0 Then
468                         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)

                        End If

                        Exit Sub

                    End If

                    'Quita el anterior
470                 If .invent.EscudoEqpObjIndex > 0 Then
472                     Call Desequipar(UserIndex, .invent.EscudoEqpSlot)

                    End If

484                 If .invent.WeaponEqpObjIndex > 0 Then
486                     If ObjData(.invent.WeaponEqpObjIndex).DosManos = 1 Then
488                         Call Desequipar(UserIndex, .invent.WeaponEqpSlot)
490                         ' Msg677=No puedes equipar un escudo si tienes un arma dos manos equipada. Tu arma fue desequipada.
                            Call WriteLocaleMsg(UserIndex, "677", e_FontTypeNames.FONTTYPE_INFOIAO)

                        End If

                    End If

492                 errordesc = "Escudo"

494                 If Len(obj.CreaGRH) <> 0 Then
496                     .Char.Escudo_Aura = obj.CreaGRH
498                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Escudo_Aura, False, 3))

                    End If

500                 .invent.Object(Slot).Equipped = 1
502                 .invent.EscudoEqpObjIndex = .invent.Object(Slot).ObjIndex
504                 .invent.EscudoEqpSlot = Slot

506                 If .flags.Navegando = 0 And .flags.Montado = 0 Then
508                     .Char.ShieldAnim = obj.ShieldAnim
510                     Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)

                    End If

512                 If obj.ResistenciaMagica > 0 Then
514                     Call WriteUpdateRM(UserIndex)

                    End If

516             Case e_OBJType.otDañoMagico, e_OBJType.otResistencia

                    'Si esta equipado lo quita
518                 If .invent.Object(Slot).Equipped Then
520                     Call Desequipar(UserIndex, Slot)
                        Exit Sub

                    End If

                    'Quita el anterior
522                 If .invent.DañoMagicoEqpSlot > 0 Then
524                     Call Desequipar(UserIndex, .invent.DañoMagicoEqpSlot)

                    End If

546                 If .invent.ResistenciaEqpSlot > 0 Then
548                     Call Desequipar(UserIndex, .invent.ResistenciaEqpSlot)

                    End If

526                 .invent.Object(Slot).Equipped = 1

                    If ObjData(.invent.Object(Slot).ObjIndex).OBJType = e_OBJType.otResistencia Then
                        .invent.ResistenciaEqpObjIndex = .invent.Object(Slot).ObjIndex
530                     .invent.ResistenciaEqpSlot = Slot
                        Call WriteUpdateRM(UserIndex)
                    ElseIf ObjData(.invent.Object(Slot).ObjIndex).OBJType = e_OBJType.otDañoMagico Then
528                     .invent.DañoMagicoEqpObjIndex = .invent.Object(Slot).ObjIndex
                        .invent.DañoMagicoEqpSlot = Slot
538                     Call WriteUpdateDM(UserIndex)

                    End If

532                 If Len(obj.CreaGRH) <> 0 Then
534                     .Char.DM_Aura = obj.CreaGRH
536                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.DM_Aura, False, 6))

                    End If

                Case e_OBJType.OtDonador

                    If obj.Subtipo = 4 Then
                        Call EquipAura(Slot, .invent, UserIndex)

                    End If

            End Select

        End With

        'Actualiza
564     Call UpdateUserInv(False, UserIndex, Slot)
        Exit Sub
ErrHandler:
566     Debug.Print errordesc
568     Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.Description & "- " & errordesc)

End Sub

Public Sub EquipAura(ByVal Slot As Integer, _
                     ByRef inventory As t_Inventario, _
                     ByVal UserIndex As Integer)

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

                    If obj.OBJType = OtDonador And obj.Subtipo = 4 Then
                        inventory.Object(Index).Equipped = 0
                        Call UpdateUserInv(False, UserIndex, Index)

                    End If

                End If

            End If

        End If

    Next Index

    inventory.Object(Slot).Equipped = 1

End Sub

Public Function CheckClaseTipo(ByVal UserIndex As Integer, _
                               ItemIndex As Integer) As Boolean

        On Error GoTo ErrHandler

100     If EsGM(UserIndex) Then
102         CheckClaseTipo = True
            Exit Function

        End If

104     Select Case ObjData(ItemIndex).ClaseTipo

            Case 0
106             CheckClaseTipo = True
                Exit Function

108         Case 2

110             If UserList(UserIndex).clase = e_Class.Mage Then CheckClaseTipo = True
112             If UserList(UserIndex).clase = e_Class.Druid Then CheckClaseTipo = True
                Exit Function

114         Case 1

116             If UserList(UserIndex).clase = e_Class.Warrior Then CheckClaseTipo = True
118             If UserList(UserIndex).clase = e_Class.Assasin Then CheckClaseTipo = True
120             If UserList(UserIndex).clase = e_Class.Bard Then CheckClaseTipo = True
122             If UserList(UserIndex).clase = e_Class.Cleric Then CheckClaseTipo = True
124             If UserList(UserIndex).clase = e_Class.Paladin Then CheckClaseTipo = True
126             If UserList(UserIndex).clase = e_Class.Trabajador Then CheckClaseTipo = True
128             If UserList(UserIndex).clase = e_Class.Hunter Then CheckClaseTipo = True
                Exit Function

        End Select

        Exit Function
ErrHandler:
130     Call LogError("Error CheckClaseTipo ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal ByClick As Byte)

        On Error GoTo hErr

        '*************************************************
        'Author: Unknown
        'Last modified: 24/01/2007
        'Handels the usage of items from inventory box.
        '24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
        '24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
        '*************************************************
        Dim obj      As t_ObjData
        Dim ObjIndex As Integer
        Dim TargObj  As t_ObjData
        Dim MiObj    As t_Obj

100     With UserList(UserIndex)

102         If .invent.Object(Slot).amount = 0 Then Exit Sub
            If Not CanUseItem(.flags, .Counters) Then
                Call WriteLocaleMsg(UserIndex, 395, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            If PuedeUsarObjeto(UserIndex, .invent.Object(Slot).ObjIndex, True) > 0 Then
                Exit Sub

            End If

104         obj = ObjData(.invent.Object(Slot).ObjIndex)

            Dim TimeSinceLastUse As Long: TimeSinceLastUse = GetTickCount() - .CdTimes(obj.cdType)

            If TimeSinceLastUse < obj.Cooldown Then Exit Sub
            If IsSet(obj.ObjFlags, e_ObjFlags.e_UseOnSafeAreaOnly) Then
                If MapInfo(.pos.Map).Seguro = 0 Then
                    ' Msg678=Solo podes usar este objeto en mapas seguros.
                    Call WriteLocaleMsg(UserIndex, "678", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            End If

106         If obj.OBJType = e_OBJType.otWeapon Then
108             If obj.Proyectil = 1 Then

                    'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
110                 If ByClick <> 0 Then
                        If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
                    Else

                        If Not IntervaloPermiteUsarClick(UserIndex) Then Exit Sub

                    End If

                Else

                    'dagas
112                 If ByClick <> 0 Then
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

118         If .flags.Meditando Then
120             .flags.Meditando = False
122             .Char.FX = 0
124             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))

            End If

126         If obj.Newbie = 1 And Not EsNewbie(UserIndex) And Not EsGM(UserIndex) Then
128             ' Msg679=Solo los newbies pueden usar estos objetos.
                Call WriteLocaleMsg(UserIndex, "679", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

130         If .Stats.ELV < obj.MinELV Then
132             Call WriteConsoleMsg(UserIndex, "Necesitas ser nivel " & obj.MinELV & " para usar este item.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

134         ObjIndex = .invent.Object(Slot).ObjIndex
136         .flags.TargetObjInvIndex = ObjIndex
138         .flags.TargetObjInvSlot = Slot

140         Select Case obj.OBJType

                Case e_OBJType.otUseOnce

142                 If .flags.Muerto = 1 Then
144                     Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    'Usa el item
146                 .Stats.MinHam = .Stats.MinHam + obj.MinHam

148                 If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
152                 Call WriteUpdateHungerAndThirst(UserIndex)
                    'Sonido
154                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.SOUND_COMIDA, .pos.x, .pos.y))
                    'Quitamos del inv el item
156                 Call QuitarUserInvItem(UserIndex, Slot, 1)
158                 Call UpdateUserInv(False, UserIndex, Slot)
                    UserList(UserIndex).flags.ModificoInventario = True

160             Case e_OBJType.otGuita

162                 If .flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
164                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

166                 .Stats.GLD = .Stats.GLD + .invent.Object(Slot).amount
168                 .invent.Object(Slot).amount = 0
170                 .invent.Object(Slot).ObjIndex = 0
172                 .invent.NroItems = .invent.NroItems - 1
                    .flags.ModificoInventario = True
174                 Call UpdateUserInv(False, UserIndex, Slot)
176                 Call WriteUpdateGold(UserIndex)

178             Case e_OBJType.otWeapon

180                 If .flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
182                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        'Msg77=¡¡Estás muerto!!.
                        Exit Sub

                    End If

184                 If Not .Stats.MinSta > 0 Then
186                     Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

188                 If ObjData(ObjIndex).Proyectil = 1 Then
                        If IsSet(.flags.StatusMask, e_StatusMask.eTransformed) Then
                            Call WriteLocaleMsg(UserIndex, MsgCantUseBowTransformed, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If

190                     Call WriteWorkRequestTarget(UserIndex, Proyectiles)
                    Else

192                     If .flags.TargetObj = Wood Then
194                         If .invent.Object(Slot).ObjIndex = DAGA Then
196                             Call TratarDeHacerFogata(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY, UserIndex)

                            End If

                        End If

                    End If

198                 If .invent.Object(Slot).Equipped = 0 Then
                        Exit Sub

                    End If

200             Case e_OBJType.otHerramientas

202                 If .flags.Muerto = 1 Then
204                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        'Msg77=¡¡Estás muerto!!.
                        Exit Sub

                    End If

206                 If Not .Stats.MinSta > 0 Then
208                     Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
210                 If .invent.Object(Slot).Equipped = 0 Then
                        'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", e_FontTypeNames.FONTTYPE_INFO)
212                     Call WriteLocaleMsg(UserIndex, "376", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

214                 Select Case obj.Subtipo

                        Case 1, 2  ' Herramientas del Pescador - Caña y Red
216                         Call WriteWorkRequestTarget(UserIndex, e_Skill.Pescar)

218                     Case 3     ' Herramientas de Alquimia - Tijeras
220                         Call WriteWorkRequestTarget(UserIndex, e_Skill.Alquimia)

222                     Case 4     ' Herramientas de Alquimia - Olla
224                         Call EnivarObjConstruiblesAlquimia(UserIndex)
226                         Call WriteShowAlquimiaForm(UserIndex)

228                     Case 5     ' Herramientas de Carpinteria - Serrucho
230                         Call EnivarObjConstruibles(UserIndex)
232                         Call WriteShowCarpenterForm(UserIndex)

234                     Case 6     ' Herramientas de Tala - Hacha
236                         Call WriteWorkRequestTarget(UserIndex, e_Skill.Talar)

238                     Case 7     ' Herramientas de Herrero - Martillo
240                         ' Msg680=Debes hacer click derecho sobre el yunque.
                            Call WriteLocaleMsg(UserIndex, "680", e_FontTypeNames.FONTTYPE_INFOIAO)

242                     Case 8     ' Herramientas de Mineria - Piquete
244                         Call WriteWorkRequestTarget(UserIndex, e_Skill.Mineria)

246                     Case 9     ' Herramientas de Sastreria - Costurero
248                         Call EnivarObjConstruiblesSastre(UserIndex)
250                         Call WriteShowSastreForm(UserIndex)

                    End Select

252             Case e_OBJType.otPociones

254                 If .flags.Muerto = 1 Then
256                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        'Msg77=¡¡Estás muerto!!.
                        Exit Sub

                    End If

                    If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
                        ' Msg681=¡¡Debes esperar unos momentos para tomar otra poción!!
                        Call WriteLocaleMsg(UserIndex, "681", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

258                 .flags.TomoPocion = True
260                 .flags.TipoPocion = obj.TipoPocion

                    Dim CabezaFinal  As Integer
                    Dim CabezaActual As Integer

262                 Select Case .flags.TipoPocion

                        Case 1 'Modif la agilidad
264                         .flags.DuracionEfecto = obj.DuracionEfecto
                            'Usa el item
266                         .Stats.UserAtributos(e_Atributos.Agilidad) = MinimoInt(.Stats.UserAtributos(e_Atributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador), .Stats.UserAtributosBackUP(e_Atributos.Agilidad) * 2)
268                         Call WriteFYA(UserIndex)
                            'Quitamos del inv el item
270                         Call QuitarUserInvItem(UserIndex, Slot, 1)

272                         If obj.Snd1 <> 0 Then
274                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Else
276                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))

                            End If

278                     Case 2 'Modif la fuerza
280                         .flags.DuracionEfecto = obj.DuracionEfecto
                            'Usa el item
282                         .Stats.UserAtributos(e_Atributos.Fuerza) = MinimoInt(.Stats.UserAtributos(e_Atributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador), .Stats.UserAtributosBackUP(e_Atributos.Fuerza) * 2)
                            'Quitamos del inv el item
284                         Call QuitarUserInvItem(UserIndex, Slot, 1)

286                         If obj.Snd1 <> 0 Then
288                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Else
290                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))

                            End If

292                         Call WriteFYA(UserIndex)

                        Case 3 'Poción roja, restaura HP

                            ' Usa el ítem
                            Dim HealingAmount As Long
                            Dim Source        As Integer
                            Dim T             As e_Trigger6
                            ' Calcula la cantidad de curación
                            HealingAmount = RandomNumber(obj.MinModificador, obj.MaxModificador) * UserMod.GetSelfHealingBonus(UserList(UserIndex))
                            ' Modifica la salud del jugador
                            Call UserMod.ModifyHealth(UserIndex, HealingAmount)
                            ' Verifica si el jugador está en la ARENA
                            T = TriggerZonaPelea(UserIndex, UserIndex)

                            ' Si NO está en un mapa entre 600 y 749 o NO está en la ARENA, se consume la poción
                            If Not ((UserList(UserIndex).pos.Map >= 600 And UserList(UserIndex).pos.Map <= 749 And T = e_Trigger6.TRIGGER6_PERMITE) Or (UserList(UserIndex).pos.Map = 275 Or UserList(UserIndex).pos.Map = 276 Or UserList(UserIndex).pos.Map = 277) Or ((UserList(UserIndex).pos.Map = 172 And T = e_Trigger6.TRIGGER6_PERMITE And (.Stats.tipoUsuario = tAventurero Or .Stats.tipoUsuario = tHeroe Or .Stats.tipoUsuario = tLeyenda)))) Then
                                Call QuitarUserInvItem(UserIndex, Slot, 1)

                            End If

                            ' Reproduce sonido al usar la poción
                            If obj.Snd1 <> 0 Then
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Else
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))

                            End If

                        Case 4 'Poción azul, restaura MANA

                            Dim porcentajeRec As Byte
                            porcentajeRec = obj.Porcentaje
                            ' Usa el ítem: restaura el MANA
                            .Stats.MinMAN = IIf(.Stats.MinMAN > 20000, 20000, .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, porcentajeRec))

                            If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN

                            ' Verifica si el jugador está en la ARENA
                            Dim triggerStatus As e_Trigger6
                            triggerStatus = TriggerZonaPelea(UserIndex, UserIndex)

                            ' Si NO está en las zonas permitidas, se consume la poción
                            If Not ((UserList(UserIndex).pos.Map >= 600 And UserList(UserIndex).pos.Map <= 749 And triggerStatus = e_Trigger6.TRIGGER6_PERMITE) Or (UserList(UserIndex).pos.Map = 275 Or UserList(UserIndex).pos.Map = 276 Or UserList(UserIndex).pos.Map = 277) Or (UserList(UserIndex).pos.Map = 172 And triggerStatus = e_Trigger6.TRIGGER6_PERMITE And (UserList(UserIndex).Stats.tipoUsuario = tAventurero Or UserList(UserIndex).Stats.tipoUsuario = tHeroe Or UserList(UserIndex).Stats.tipoUsuario = tLeyenda))) Then
                                ' Quitamos el ítem del inventario
                                Call QuitarUserInvItem(UserIndex, Slot, 1)

                            End If

                            ' Reproduce sonido al usar la poción
                            If obj.Snd1 <> 0 Then
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Else
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))

                            End If

324                     Case 5 ' Pocion violeta

326                         If .flags.Envenenado > 0 Then
328                             .flags.Envenenado = 0
330                             ' Msg682=Te has curado del envenenamiento.
                                Call WriteLocaleMsg(UserIndex, "682", e_FontTypeNames.FONTTYPE_INFO)
                                'Quitamos del inv el item
332                             Call QuitarUserInvItem(UserIndex, Slot, 1)

334                             If obj.Snd1 <> 0 Then
336                                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                                Else
338                                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))

                                End If

                            Else
340                             ' Msg683=¡No te encuentras envenenado!
                                Call WriteLocaleMsg(UserIndex, "683", e_FontTypeNames.FONTTYPE_INFO)

                            End If

342                     Case 6  ' Remueve Parálisis

344                         If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
346                             If .flags.Paralizado = 1 Then
348                                 .flags.Paralizado = 0
350                                 Call WriteParalizeOK(UserIndex)

                                End If

352                             If .flags.Inmovilizado = 1 Then
354                                 .Counters.Inmovilizado = 0
356                                 .flags.Inmovilizado = 0
358                                 Call WriteInmovilizaOK(UserIndex)

                                End If

360                             Call QuitarUserInvItem(UserIndex, Slot, 1)

362                             If obj.Snd1 <> 0 Then
364                                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                                Else
366                                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(255, .pos.x, .pos.y))

                                End If

368                             ' Msg684=Te has removido la paralizis.
                                Call WriteLocaleMsg(UserIndex, "684", e_FontTypeNames.FONTTYPE_INFOIAO)
                            Else
370                             ' Msg685=No estas paralizado.
                                Call WriteLocaleMsg(UserIndex, "685", e_FontTypeNames.FONTTYPE_INFOIAO)

                            End If

372                     Case 7  ' Pocion Naranja
374                         .Stats.MinSta = .Stats.MinSta + RandomNumber(obj.MinModificador, obj.MaxModificador)

376                         If .Stats.MinSta > .Stats.MaxSta Then .Stats.MinSta = .Stats.MaxSta
                            'Quitamos del inv el item
378                         Call QuitarUserInvItem(UserIndex, Slot, 1)

380                         If obj.Snd1 <> 0 Then
382                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Else
384                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))

                            End If

386                     Case 8  ' Pocion cambio cara

388                         Select Case .genero

                                Case e_Genero.Hombre

390                                 Select Case .raza

                                        Case e_Raza.Humano
392                                         CabezaFinal = RandomNumber(1, 40)

394                                     Case e_Raza.Elfo
396                                         CabezaFinal = RandomNumber(101, 132)

398                                     Case e_Raza.Drow
400                                         CabezaFinal = RandomNumber(201, 229)

402                                     Case e_Raza.Enano
404                                         CabezaFinal = RandomNumber(301, 329)

406                                     Case e_Raza.Gnomo
408                                         CabezaFinal = RandomNumber(401, 429)

410                                     Case e_Raza.Orco
412                                         CabezaFinal = RandomNumber(501, 529)

                                    End Select

414                             Case e_Genero.Mujer

416                                 Select Case .raza

                                        Case e_Raza.Humano
418                                         CabezaFinal = RandomNumber(50, 80)

420                                     Case e_Raza.Elfo
422                                         CabezaFinal = RandomNumber(150, 179)

424                                     Case e_Raza.Drow
426                                         CabezaFinal = RandomNumber(250, 279)

428                                     Case e_Raza.Gnomo
430                                         CabezaFinal = RandomNumber(350, 379)

432                                     Case e_Raza.Enano
434                                         CabezaFinal = RandomNumber(450, 479)

436                                     Case e_Raza.Orco
438                                         CabezaFinal = RandomNumber(550, 579)

                                    End Select

                            End Select

440                         .Char.head = CabezaFinal
442                         .OrigChar.head = CabezaFinal
444                         Call ChangeUserChar(UserIndex, .Char.body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)
                            'Quitamos del inv el item
                            UserList(UserIndex).Counters.timeFx = 3
446                         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 102, 0, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))

448                         If CabezaActual <> CabezaFinal Then
450                             Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Else
452                             ' Msg686=¡Rayos! Te tocó la misma cabeza, item no consumido. Tienes otra oportunidad.
                                Call WriteLocaleMsg(UserIndex, "686", e_FontTypeNames.FONTTYPE_INFOIAO)

                            End If

454                         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))

456                     Case 9  ' Pocion sexo

458                         Select Case .genero

                                Case e_Genero.Hombre
460                                 .genero = e_Genero.Mujer

462                             Case e_Genero.Mujer
464                                 .genero = e_Genero.Hombre

                            End Select

466                         Select Case .genero

                                Case e_Genero.Hombre

468                                 Select Case .raza

                                        Case e_Raza.Humano
470                                         CabezaFinal = RandomNumber(1, 40)

472                                     Case e_Raza.Elfo
474                                         CabezaFinal = RandomNumber(101, 132)

476                                     Case e_Raza.Drow
478                                         CabezaFinal = RandomNumber(201, 229)

480                                     Case e_Raza.Enano
482                                         CabezaFinal = RandomNumber(301, 329)

484                                     Case e_Raza.Gnomo
486                                         CabezaFinal = RandomNumber(401, 429)

488                                     Case e_Raza.Orco
490                                         CabezaFinal = RandomNumber(501, 529)

                                    End Select

492                             Case e_Genero.Mujer

494                                 Select Case .raza

                                        Case e_Raza.Humano
496                                         CabezaFinal = RandomNumber(50, 80)

498                                     Case e_Raza.Elfo
500                                         CabezaFinal = RandomNumber(150, 179)

502                                     Case e_Raza.Drow
504                                         CabezaFinal = RandomNumber(250, 279)

506                                     Case e_Raza.Gnomo
508                                         CabezaFinal = RandomNumber(350, 379)

510                                     Case e_Raza.Enano
512                                         CabezaFinal = RandomNumber(450, 479)

514                                     Case e_Raza.Orco
516                                         CabezaFinal = RandomNumber(550, 579)

                                    End Select

                            End Select

518                         .Char.head = CabezaFinal
520                         .OrigChar.head = CabezaFinal
522                         Call ChangeUserChar(UserIndex, .Char.body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)
                            'Quitamos del inv el item
                            UserList(UserIndex).Counters.timeFx = 3
524                         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 102, 0, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
526                         Call QuitarUserInvItem(UserIndex, Slot, 1)

528                         If obj.Snd1 <> 0 Then
530                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Else
532                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))

                            End If

534                     Case 10  ' Invisibilidad

536                         If .flags.invisible = 0 And .Counters.DisabledInvisibility = 0 Then
                                If IsSet(.flags.StatusMask, eTaunting) Then
                                    ' Msg687=No tiene efecto.
                                    Call WriteLocaleMsg(UserIndex, "687", e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                                    Exit Sub

                                End If

538                             .flags.invisible = 1
540                             .Counters.Invisibilidad = obj.DuracionEfecto
542                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True, .pos.x, .pos.y))
544                             Call WriteContadores(UserIndex)
546                             Call QuitarUserInvItem(UserIndex, Slot, 1)

548                             If obj.Snd1 <> 0 Then
550                                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                                Else
552                                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave("123", .pos.x, .pos.y))

                                End If

554                             ' Msg688=Te has escondido entre las sombras...
                                Call WriteLocaleMsg(UserIndex, "688", e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                            Else
556                             ' Msg689=Ya estás invisible.
                                Call WriteLocaleMsg(UserIndex, "689", e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                                Exit Sub

                            End If

                            ' Poción que limpia todo
626                     Case 13
628                         Call QuitarUserInvItem(UserIndex, Slot, 1)
630                         .flags.Envenenado = 0
632                         .flags.Incinerado = 0

634                         If .flags.Inmovilizado = 1 Then
636                             .Counters.Inmovilizado = 0
638                             .flags.Inmovilizado = 0
640                             Call WriteInmovilizaOK(UserIndex)

                            End If

642                         If .flags.Paralizado = 1 Then
644                             .flags.Paralizado = 0
646                             Call WriteParalizeOK(UserIndex)

                            End If

648                         If .flags.Ceguera = 1 Then
650                             .flags.Ceguera = 0
652                             Call WriteBlindNoMore(UserIndex)

                            End If

654                         If .flags.Maldicion = 1 Then
656                             .flags.Maldicion = 0
658                             .Counters.Maldicion = 0

                            End If

660                         .Stats.MinSta = .Stats.MaxSta
662                         .Stats.MinAGU = .Stats.MaxAGU
664                         .Stats.MinMAN = .Stats.MaxMAN
666                         .Stats.MinHp = .Stats.MaxHp
668                         .Stats.MinHam = .Stats.MaxHam
674                         Call WriteUpdateHungerAndThirst(UserIndex)
676                         ' Msg690=Donador> Te sentís sano y lleno.
                            Call WriteLocaleMsg(UserIndex, "690", e_FontTypeNames.FONTTYPE_WARNING)

678                         If obj.Snd1 <> 0 Then
680                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Else
682                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))

                            End If

                            ' Poción runa
684                     Case 14

686                         If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
688                             ' Msg691=No podés usar la runa estando en la cárcel.
                                Call WriteLocaleMsg(UserIndex, "691", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If

                            Dim Map     As Integer
                            Dim x       As Byte
                            Dim y       As Byte
                            Dim DeDonde As t_WorldPos
690                         Call QuitarUserInvItem(UserIndex, Slot, 1)

692                         Select Case .Hogar

                                Case e_Ciudad.cUllathorpe
694                                 DeDonde = Ullathorpe

696                             Case e_Ciudad.cNix
698                                 DeDonde = Nix

700                             Case e_Ciudad.cBanderbill
702                                 DeDonde = Banderbill

704                             Case e_Ciudad.cLindos
706                                 DeDonde = Lindos

708                             Case e_Ciudad.cArghal
710                                 DeDonde = Arghal

712                             Case e_Ciudad.cArkhein
714                                 DeDonde = Arkhein

716                             Case Else
718                                 DeDonde = Ullathorpe

                            End Select

720                         Map = DeDonde.Map
722                         x = DeDonde.x
724                         y = DeDonde.y
726                         Call FindLegalPos(UserIndex, Map, x, y)
728                         Call WarpUserChar(UserIndex, Map, x, y, True)
                            'Msg884= Ya estas a salvo...
                            Call WriteLocaleMsg(UserIndex, "884", e_FontTypeNames.FONTTYPE_WARNING)

732                         If obj.Snd1 <> 0 Then
734                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Else
736                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))

                            End If

774                     Case 16 ' Divorcio

776                         If .flags.Casado = 1 Then

                                Dim tUser As t_UserReference
                                '.flags.Pareja
778                             tUser = NameIndex(GetUserSpouse(.flags.SpouseId))

782                             If Not IsValidUserRef(tUser) Then
                                    'Msg885= Tu pareja deberás estar conectada para divorciarse.
                                    Call WriteLocaleMsg(UserIndex, "885", e_FontTypeNames.FONTTYPE_INFOIAO)
                                Else
780                                 Call QuitarUserInvItem(UserIndex, Slot, 1)
794                                 UserList(tUser.ArrayIndex).flags.Casado = 0
796                                 UserList(tUser.ArrayIndex).flags.SpouseId = 0
798                                 .flags.Casado = 0
800                                 .flags.SpouseId = 0
                                    'Msg886= Te has divorciado.
                                    Call WriteLocaleMsg(UserIndex, "886", e_FontTypeNames.FONTTYPE_INFOIAO)
804                                 Call WriteConsoleMsg(tUser.ArrayIndex, .name & " se ha divorciado de ti.", e_FontTypeNames.FONTTYPE_INFOIAO)

                                    If obj.Snd1 <> 0 Then
808                                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                                    Else
810                                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))

                                    End If

                                End If

806
                            Else
                                'Msg887= No estas casado.
                                Call WriteLocaleMsg(UserIndex, "887", e_FontTypeNames.FONTTYPE_INFOIAO)

                            End If

814                     Case 17 'Cara legendaria

816                         Select Case .genero

                                Case e_Genero.Hombre

818                                 Select Case .raza

                                        Case e_Raza.Humano
820                                         CabezaFinal = RandomNumber(684, 686)

822                                     Case e_Raza.Elfo
824                                         CabezaFinal = RandomNumber(690, 692)

826                                     Case e_Raza.Drow
828                                         CabezaFinal = RandomNumber(696, 698)

830                                     Case e_Raza.Enano
832                                         CabezaFinal = RandomNumber(702, 704)

834                                     Case e_Raza.Gnomo
836                                         CabezaFinal = RandomNumber(708, 710)

838                                     Case e_Raza.Orco
840                                         CabezaFinal = RandomNumber(714, 716)

                                    End Select

842                             Case e_Genero.Mujer

844                                 Select Case .raza

                                        Case e_Raza.Humano
846                                         CabezaFinal = RandomNumber(687, 689)

848                                     Case e_Raza.Elfo
850                                         CabezaFinal = RandomNumber(693, 695)

852                                     Case e_Raza.Drow
854                                         CabezaFinal = RandomNumber(699, 701)

856                                     Case e_Raza.Gnomo
858                                         CabezaFinal = RandomNumber(705, 707)

860                                     Case e_Raza.Enano
862                                         CabezaFinal = RandomNumber(711, 713)

864                                     Case e_Raza.Orco
866                                         CabezaFinal = RandomNumber(717, 719)

                                    End Select

                            End Select

868                         CabezaActual = .OrigChar.head
870                         .Char.head = CabezaFinal
872                         .OrigChar.head = CabezaFinal
874                         Call ChangeUserChar(UserIndex, .Char.body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)

                            'Quitamos del inv el item
876                         If CabezaActual <> CabezaFinal Then
                                UserList(UserIndex).Counters.timeFx = 3
878                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 102, 0, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
880                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
882                             Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Else
                                'Msg888= ¡Rayos! No pude asignarte una cabeza nueva, item no consumido. ¡Proba de nuevo!
                                Call WriteLocaleMsg(UserIndex, "888", e_FontTypeNames.FONTTYPE_INFOIAO)

                            End If

886                     Case 18  ' tan solo crea una particula por determinado tiempo

                            Dim Particula           As Integer
                            Dim Tiempo              As Long
                            Dim ParticulaPermanente As Byte
                            Dim sobrechar           As Byte

888                         If obj.CreaParticula <> "" Then
890                             Particula = val(ReadField(1, obj.CreaParticula, Asc(":")))
892                             Tiempo = val(ReadField(2, obj.CreaParticula, Asc(":")))
894                             ParticulaPermanente = val(ReadField(3, obj.CreaParticula, Asc(":")))
896                             sobrechar = val(ReadField(4, obj.CreaParticula, Asc(":")))

898                             If ParticulaPermanente = 1 Then
900                                 .Char.ParticulaFx = Particula
902                                 .Char.loops = Tiempo

                                End If

904                             If sobrechar = 1 Then
906                                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFXToFloor(.pos.x, .pos.y, Particula, Tiempo))
                                Else
                                    UserList(UserIndex).Counters.timeFx = 3
908                                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, Particula, Tiempo, False, , UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))

                                End If

                            End If

910                         If obj.CreaFX <> 0 Then
912                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageFxPiso(obj.CreaFX, .pos.x, .pos.y))

                            End If

914                         If obj.Snd1 <> 0 Then
916                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))

                            End If

918                         Call QuitarUserInvItem(UserIndex, Slot, 1)

920                     Case 19 ' Reseteo de skill

                            Dim s As Byte

922                         If .Stats.UserSkills(e_Skill.liderazgo) >= 80 Then
                                'Msg889= Has fundado un clan, no podes resetar tus skills.
                                Call WriteLocaleMsg(UserIndex, "889", e_FontTypeNames.FONTTYPE_INFOIAO)
                                Exit Sub

                            End If

926                         For s = 1 To NUMSKILLS
928                             .Stats.UserSkills(s) = 0
930                         Next s

                            Dim SkillLibres As Integer
932                         SkillLibres = 5
934                         SkillLibres = SkillLibres + (5 * .Stats.ELV)
936                         .Stats.SkillPts = SkillLibres
938                         Call WriteLevelUp(UserIndex, .Stats.SkillPts)
                            'Msg890= Tus skills han sido reseteados.
                            Call WriteLocaleMsg(UserIndex, "890", e_FontTypeNames.FONTTYPE_INFOIAO)
942                         Call QuitarUserInvItem(UserIndex, Slot, 1)

                            ' Mochila
944                     Case 20

946                         If .Stats.InventLevel < INVENTORY_EXTRA_ROWS Then
948                             .Stats.InventLevel = .Stats.InventLevel + 1
950                             .CurrentInventorySlots = getMaxInventorySlots(UserIndex)
952                             Call WriteInventoryUnlockSlots(UserIndex)
                                'Msg891= Has aumentado el espacio de tu inventario!
                                Call WriteLocaleMsg(UserIndex, "891", e_FontTypeNames.FONTTYPE_INFO)
956                             Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Else
                                'Msg892= Ya has desbloqueado todos los casilleros disponibles.
                                Call WriteLocaleMsg(UserIndex, "892", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If

                            ' Poción negra (suicidio)
960                     Case 21
                            'Quitamos del inv el item
962                         Call QuitarUserInvItem(UserIndex, Slot, 1)

964                         If obj.Snd1 <> 0 Then
966                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Else
968                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))

                            End If

                            'Msg893= Te has suicidado.
                            Call WriteLocaleMsg(UserIndex, "893", e_FontTypeNames.FONTTYPE_EJECUCION)
                            Call CustomScenarios.UserDie(UserIndex)
972                         Call UserMod.UserDie(UserIndex)

                            'Poción de reset (resetea el personaje)
                        Case 22

                            If GetTickCount - .Counters.LastResetTick > 3000 Then
                                Call writeAnswerReset(UserIndex)
                                .Counters.LastResetTick = GetTickCount
                            Else
                                'Msg894= Debes esperar unos momentos para tomar esta poción.
                                Call WriteLocaleMsg(UserIndex, "894", e_FontTypeNames.FONTTYPE_INFO)

                            End If

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

974                 Call WriteUpdateUserStats(UserIndex)
976                 Call UpdateUserInv(False, UserIndex, Slot)

978             Case e_OBJType.otBebidas

980                 If .flags.Muerto = 1 Then
982                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

984                 .Stats.MinAGU = .Stats.MinAGU + obj.MinSed

986                 If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
990                 Call WriteUpdateHungerAndThirst(UserIndex)
                    'Quitamos del inv el item
992                 Call QuitarUserInvItem(UserIndex, Slot, 1)

                    If obj.ApplyEffectId > 0 Then
                        Call AddOrResetEffect(UserIndex, obj.ApplyEffectId)

                    End If

994                 If obj.Snd1 <> 0 Then
996                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                    Else
998                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))

                    End If

1000                Call UpdateUserInv(False, UserIndex, Slot)

1002            Case e_OBJType.OtCofre

1004                If .flags.Muerto = 1 Then
1006                    Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        'Msg77=¡¡Estás muerto!!.
                        Exit Sub

                    End If

                    'Quitamos del inv el item
1008                Call QuitarUserInvItem(UserIndex, Slot, 1)
1010                Call UpdateUserInv(False, UserIndex, Slot)
1012                Call WriteConsoleMsg(UserIndex, "Has abierto un " & obj.name & " y obtuviste...", e_FontTypeNames.FONTTYPE_New_DONADOR)

1014                If obj.Snd1 <> 0 Then
1016                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))

                    End If

1018                If obj.CreaFX <> 0 Then
                        UserList(UserIndex).Counters.timeFx = 3
1020                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, obj.CreaFX, 0, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))

                    End If

                    Dim i As Byte

1022                Select Case obj.Subtipo

                        Case 1

1024                        For i = 1 To obj.CantItem

1026                            If Not MeterItemEnInventario(UserIndex, obj.Item(i)) Then
1028                                If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Dios Or e_PlayerType.Admin)) Then
1030                                    Call TirarItemAlPiso(.pos, obj.Item(i))

                                    End If

                                End If

1032                            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).name & " (" & obj.Item(i).amount & ")", e_FontTypeNames.FONTTYPE_INFOBOLD))
1034                        Next i

                        Case 2

1036                        For i = 1 To obj.CantEntrega

                                Dim indexobj As Byte
1038                            indexobj = RandomNumber(1, obj.CantItem)

                                Dim Index As t_Obj
1040                            Index.ObjIndex = obj.Item(indexobj).ObjIndex
1042                            Index.amount = obj.Item(indexobj).amount

1044                            If Not MeterItemEnInventario(UserIndex, Index) Then
1046                                If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Dios Or e_PlayerType.Admin)) Then
1048                                    Call TirarItemAlPiso(.pos, Index)

                                    End If

                                End If

1050                            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageConsoleMsg(ObjData(Index.ObjIndex).name & " (" & Index.amount & ")", e_FontTypeNames.FONTTYPE_INFOBOLD))
1052                        Next i

                        Case 3

                            For i = 1 To obj.CantItem

                                If RandomNumber(1, obj.Item(i).data) = 1 Then
                                    If Not MeterItemEnInventario(UserIndex, obj.Item(i)) Then
                                        If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Dios Or e_PlayerType.Admin)) Then
                                            Call TirarItemAlPiso(.pos, obj.Item(i))

                                        End If

                                    End If

                                    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).name & " (" & obj.Item(i).amount & ")", e_FontTypeNames.FONTTYPE_INFOBOLD))

                                End If

                            Next i

                    End Select

1054            Case e_OBJType.otLlaves

                    If UserList(UserIndex).flags.Muerto = 1 Then
                        'Msg895= ¡¡Estas muerto!! Solo podes usar items cuando estas vivo.
                        Call WriteLocaleMsg(UserIndex, "895", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
                    TargObj = ObjData(UserList(UserIndex).flags.TargetObj)

                    '¿El objeto clickeado es una puerta?
                    If TargObj.OBJType = e_OBJType.otPuertas Then
                        If TargObj.clave < 1000 Then
                            'Msg896= Las llaves en el inventario están desactivadas. Sólo se permiten en el llavero.
                            Call WriteLocaleMsg(UserIndex, "896", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If

                        '¿Esta cerrada?
                        If TargObj.Cerrada = 1 Then

                            '¿Cerrada con llave?
                            If TargObj.Llave > 0 Then

                                Dim ClaveLlave As Integer

                                If TargObj.clave = obj.clave Then
                                    MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                                    UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
                                    'Msg897= Has abierto la puerta.
                                    Call WriteLocaleMsg(UserIndex, "897", e_FontTypeNames.FONTTYPE_INFO)
                                    ClaveLlave = obj.clave
                                    Call EliminarLlaves(ClaveLlave, UserIndex)
                                    Exit Sub
                                Else
                                    'Msg898= La llave no sirve.
                                    Call WriteLocaleMsg(UserIndex, "898", e_FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub

                                End If

                            Else

                                If TargObj.clave = obj.clave Then
                                    MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                                    'Msg899= Has cerrado con llave la puerta.
                                    Call WriteLocaleMsg(UserIndex, "899", e_FontTypeNames.FONTTYPE_INFO)
                                    UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
                                    Exit Sub
                                Else
                                    'Msg900= La llave no sirve.
                                    Call WriteLocaleMsg(UserIndex, "900", e_FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub

                                End If

                            End If

                        Else
                            'Msg901= No esta cerrada.
                            Call WriteLocaleMsg(UserIndex, "901", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If

                    End If

1058            Case e_OBJType.otBotellaVacia

1060                If .flags.Muerto = 1 Then
1062                    Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If Not InMapBounds(.flags.TargetMap, .flags.TargetX, .flags.TargetY) Then
                        Exit Sub

                    End If

1064                If (MapData(.pos.Map, .flags.TargetX, .flags.TargetY).Blocked And FLAG_AGUA) = 0 Then
                        'Msg902= No hay agua allí.
                        Call WriteLocaleMsg(UserIndex, "902", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                    If Distance(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, .flags.TargetX, .flags.TargetY) > 2 Then
                        'Msg903= Debes acercarte más al agua.
                        Call WriteLocaleMsg(UserIndex, "903", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

1068                MiObj.amount = 1
1070                MiObj.ObjIndex = ObjData(.invent.Object(Slot).ObjIndex).IndexAbierta
1072                Call QuitarUserInvItem(UserIndex, Slot, 1)

1074                If Not MeterItemEnInventario(UserIndex, MiObj) Then
1076                    Call TirarItemAlPiso(.pos, MiObj)

                    End If

1078                Call UpdateUserInv(False, UserIndex, Slot)

1080            Case e_OBJType.otBotellaLlena

1082                If .flags.Muerto = 1 Then
1084                    Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        'Msg77=¡¡Estás muerto!!.
                        Exit Sub

                    End If

1086                .Stats.MinAGU = .Stats.MinAGU + obj.MinSed

1088                If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
1092                Call WriteUpdateHungerAndThirst(UserIndex)
1094                MiObj.amount = 1
1096                MiObj.ObjIndex = ObjData(.invent.Object(Slot).ObjIndex).IndexCerrada
1098                Call QuitarUserInvItem(UserIndex, Slot, 1)

1100                If Not MeterItemEnInventario(UserIndex, MiObj) Then
1102                    Call TirarItemAlPiso(.pos, MiObj)

                    End If

1104                Call UpdateUserInv(False, UserIndex, Slot)

1106            Case e_OBJType.otPergaminos

1108                If .flags.Muerto = 1 Then
1110                    Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        'Msg77=¡¡Estás muerto!!.
                        Exit Sub

                    End If

                    'Call LogError(.Name & " intento aprender el hechizo " & ObjData(.Invent.Object(slot).ObjIndex).HechizoIndex)
1112                If ClasePuedeUsarItem(UserIndex, .invent.Object(Slot).ObjIndex, Slot) And RazaPuedeUsarItem(UserIndex, .invent.Object(Slot).ObjIndex, Slot) Then

                        'If .Stats.MaxMAN > 0 Then
1114                    If .Stats.MinHam > 0 And .Stats.MinAGU > 0 Then
1116                        Call AgregarHechizo(UserIndex, Slot)
1118                        Call UpdateUserInv(False, UserIndex, Slot)
                            ' Call LogError(.Name & " lo aprendio.")
                        Else
                            'Msg904= Estas demasiado hambriento y sediento.
                            Call WriteLocaleMsg(UserIndex, "904", e_FontTypeNames.FONTTYPE_INFO)

                        End If

                        ' Else
                        '    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", e_FontTypeNames.FONTTYPE_WARNING)
                        'End If
                    Else
                        'Msg906= Por mas que lo intentas, no podés comprender el manuescrito.
                        Call WriteLocaleMsg(UserIndex, "906", e_FontTypeNames.FONTTYPE_INFO)

                    End If

1124            Case e_OBJType.otMinerales

1126                If .flags.Muerto = 1 Then
1128                    Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        'Msg77=¡¡Estás muerto!!.
                        Exit Sub

                    End If

1130                Call WriteWorkRequestTarget(UserIndex, FundirMetal)

1132            Case e_OBJType.otInstrumentos

1134                If .flags.Muerto = 1 Then
1136                    Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        'Msg77=¡¡Estás muerto!!.
                        Exit Sub

                    End If

1138                If obj.Real Then '¿Es el Cuerno Real?
1140                    If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
1142                        If MapInfo(.pos.Map).Seguro = 1 Then
                                'Msg907= No hay Peligro aquí. Es Zona Segura
                                Call WriteLocaleMsg(UserIndex, "907", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If

1146                        Call SendData(SendTarget.toMap, .pos.Map, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Exit Sub
                        Else
                            'Msg908= Solo Miembros de la Armada Real pueden usar este cuerno.
                            Call WriteLocaleMsg(UserIndex, "908", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If

1150                ElseIf obj.Caos Then '¿Es el Cuerno Legión?

1152                    If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
1154                        If MapInfo(.pos.Map).Seguro = 1 Then
                                'Msg909= No hay Peligro aquí. Es Zona Segura
                                Call WriteLocaleMsg(UserIndex, "909", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If

1158                        Call SendData(SendTarget.toMap, .pos.Map, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
                            Exit Sub
                        Else
                            'Msg910= Solo Miembros de la Legión Oscura pueden usar este cuerno.
                            Call WriteLocaleMsg(UserIndex, "910", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If

                    End If

                    'Si llega aca es porque es o Laud o Tambor o Flauta
1162                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))

1164            Case e_OBJType.otBarcos

                    ' Piratas y trabajadores navegan al nivel 23
                    If .invent.Object(Slot).ObjIndex <> iObjTrajeAltoNw And .invent.Object(Slot).ObjIndex <> iObjTrajeBajoNw And .invent.Object(Slot).ObjIndex <> iObjTraje Then
1166                    If .clase = e_Class.Trabajador Or .clase = e_Class.Pirat Then
1168                        If .Stats.ELV < 23 Then
                                'Msg911= Para recorrer los mares debes ser nivel 23 o superior.
                                Call WriteLocaleMsg(UserIndex, "911", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If

                            ' Nivel mínimo 25 para navegar, si no sos pirata ni trabajador
1172                    ElseIf .Stats.ELV < 25 Then
                            'Msg912= Para recorrer los mares debes ser nivel 25 o superior.
                            Call WriteLocaleMsg(UserIndex, "912", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If

                    ElseIf .invent.Object(Slot).ObjIndex = iObjTrajeAltoNw Or .invent.Object(Slot).ObjIndex = iObjTrajeBajoNw Then

                        If (.flags.Navegando = 0 Or (.invent.BarcoObjIndex <> iObjTrajeAltoNw And .invent.BarcoObjIndex <> iObjTrajeBajoNw)) And MapData(.pos.Map, .pos.x + 1, .pos.y).trigger <> e_Trigger.DETALLEAGUA And MapData(.pos.Map, .pos.x - 1, .pos.y).trigger <> e_Trigger.DETALLEAGUA And MapData(.pos.Map, .pos.x, .pos.y + 1).trigger <> e_Trigger.DETALLEAGUA And MapData(.pos.Map, .pos.x, .pos.y - 1).trigger <> e_Trigger.DETALLEAGUA Then
                            'Msg913= Este traje es para aguas contaminadas.
                            Call WriteLocaleMsg(UserIndex, "913", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If

                    ElseIf .invent.Object(Slot).ObjIndex = iObjTraje Then

                        If (.flags.Navegando = 0 Or .invent.BarcoObjIndex <> iObjTraje) And MapData(.pos.Map, .pos.x + 1, .pos.y).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.x - 1, .pos.y).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.x, .pos.y + 1).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.x, .pos.y - 1).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.x + 1, .pos.y).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.x - 1, .pos.y).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.x, .pos.y + 1).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.x, .pos.y - 1).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.x + 1, .pos.y).trigger <> e_Trigger.NADOBAJOTECHO And MapData(.pos.Map, .pos.x - 1, .pos.y).trigger <> e_Trigger.NADOBAJOTECHO And MapData(.pos.Map, .pos.x, .pos.y + 1).trigger <> e_Trigger.NADOBAJOTECHO And MapData(.pos.Map, .pos.x, .pos.y - 1).trigger <> e_Trigger.NADOBAJOTECHO Then
                            'Msg914= Este traje es para zonas poco profundas.
                            Call WriteLocaleMsg(UserIndex, "914", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If

                    End If

1176                If .flags.Navegando = 0 Then
1178                    If LegalWalk(.pos.Map, .pos.x - 1, .pos.y, e_Heading.WEST, True, False) Or LegalWalk(.pos.Map, .pos.x, .pos.y - 1, e_Heading.NORTH, True, False) Or LegalWalk(.pos.Map, .pos.x + 1, .pos.y, e_Heading.EAST, True, False) Or LegalWalk(.pos.Map, .pos.x, .pos.y + 1, e_Heading.SOUTH, True, False) Then
1180                        Call DoNavega(UserIndex, obj, Slot)
                        Else
                            'Msg915= ¡Debes aproximarte al agua para usar el barco o traje de baño!
                            Call WriteLocaleMsg(UserIndex, "915", e_FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else

1184                    If .invent.BarcoObjIndex <> .invent.Object(Slot).ObjIndex Then
1186                        Call DoNavega(UserIndex, obj, Slot)
                        Else

1188                        If LegalWalk(.pos.Map, .pos.x - 1, .pos.y, e_Heading.WEST, False, True) Or LegalWalk(.pos.Map, .pos.x, .pos.y - 1, e_Heading.NORTH, False, True) Or LegalWalk(.pos.Map, .pos.x + 1, .pos.y, e_Heading.EAST, False, True) Or LegalWalk(.pos.Map, .pos.x, .pos.y + 1, e_Heading.SOUTH, False, True) Then
1190                            Call DoNavega(UserIndex, obj, Slot)
                            Else
                                'Msg916= ¡Debes aproximarte a la costa para dejar la barca!
                                Call WriteLocaleMsg(UserIndex, "916", e_FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

                    End If

1194            Case e_OBJType.otMonturas

                    'Verifica todo lo que requiere la montura
1196                If .flags.Muerto = 1 Then
1198                    Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        'Msg77=¡¡Estás muerto!!.
                        Exit Sub

                    End If

1200                If .flags.Navegando = 1 Then
                        'Msg917= Debes dejar de navegar para poder cabalgar.
                        Call WriteLocaleMsg(UserIndex, "917", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

1204                If MapInfo(.pos.Map).zone = "DUNGEON" Then
                        'Msg918= No podes cabalgar dentro de un dungeon.
                        Call WriteLocaleMsg(UserIndex, "918", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

1208                Call DoMontar(UserIndex, obj, Slot)

                Case e_OBJType.OtDonador

                    Select Case obj.Subtipo

                        Case 1

1214                        If .Counters.Pena <> 0 Then
1216                            ' Msg691=No podés usar la runa estando en la cárcel.
                                Call WriteLocaleMsg(UserIndex, "691", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If

1218                        If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
1220                            ' Msg691=No podés usar la runa estando en la cárcel.
                                Call WriteLocaleMsg(UserIndex, "691", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If

1222                        Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
                            'Msg919= Has viajado por el mundo.
                            Call WriteLocaleMsg(UserIndex, "919", e_FontTypeNames.FONTTYPE_WARNING)
1226                        Call QuitarUserInvItem(UserIndex, Slot, 1)
1228                        Call UpdateUserInv(False, UserIndex, Slot)

1230                    Case 2
                            Exit Sub

1252                    Case 3
                            Exit Sub

                    End Select

1262            Case e_OBJType.otpasajes

1264                If .flags.Muerto = 1 Then
                        'Msg77=¡¡Estás muerto!!.
1266                    Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

1268                If .flags.TargetNpcTipo <> Pirata Then
                        'Msg920= Primero debes hacer click sobre el pirata.
                        Call WriteLocaleMsg(UserIndex, "920", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

1272                If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 3 Then
1274                    Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

1276                If .pos.Map <> obj.DesdeMap Then
1278                    Call WriteChatOverHead(UserIndex, "El pasaje no lo compraste aquí! Largate!", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite)
                        Exit Sub

                    End If

1280                If Not MapaValido(obj.HastaMap) Then
1282                    Call WriteChatOverHead(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite)
                        Exit Sub

                    End If

1284                If obj.NecesitaNave > 0 Then
1286                    If .Stats.UserSkills(e_Skill.Navegacion) < 80 Then
1288                        Call WriteChatOverHead(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite)
                            Exit Sub

                        End If

                    End If

1290                Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
                    'Msg921= Has viajado por varios días, te sientes exhausto!
                    Call WriteLocaleMsg(UserIndex, "921", e_FontTypeNames.FONTTYPE_WARNING)
1294                .Stats.MinAGU = 0
1296                .Stats.MinHam = 0
1302                Call WriteUpdateHungerAndThirst(UserIndex)
1304                Call QuitarUserInvItem(UserIndex, Slot, 1)
1306                Call UpdateUserInv(False, UserIndex, Slot)

1308            Case e_OBJType.otRunas

1310                If .Counters.Pena <> 0 Then
1312                    ' Msg691=No podés usar la runa estando en la cárcel.
                        Call WriteLocaleMsg(UserIndex, "691", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

1314                If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
1316                    ' Msg691=No podés usar la runa estando en la cárcel.
                        Call WriteLocaleMsg(UserIndex, "691", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

1318                If MapInfo(.pos.Map).Seguro = 0 And .flags.Muerto = 0 Then
1320                    ' Msg692=Solo podes usar tu runa en zonas seguras.
                        Call WriteLocaleMsg(UserIndex, "692", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

1322                If .Accion.AccionPendiente Then
                        Exit Sub

                    End If

1324                Select Case ObjData(ObjIndex).TipoRuna

                        Case 1, 2

1326                        If Not EsGM(UserIndex) Then
1328                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, e_ParticulasIndex.Runa, 400, False))
1330                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageBarFx(.Char.charindex, 350, e_AccionBarra.Runa))
                            Else
1332                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, e_ParticulasIndex.Runa, 50, False))
1334                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageBarFx(.Char.charindex, 100, e_AccionBarra.Runa))

                            End If

1336                        .Accion.Particula = e_ParticulasIndex.Runa
1338                        .Accion.AccionPendiente = True
1340                        .Accion.TipoAccion = e_AccionBarra.Runa
1342                        .Accion.RunaObj = ObjIndex
1344                        .Accion.ObjSlot = Slot

                    End Select

1346            Case e_OBJType.otmapa
1348                Call WriteShowFrmMapa(UserIndex)

                Case e_OBJType.OtQuest

1349                If obj.QuestId > 0 Then Call WriteObjQuestSend(UserIndex, obj.QuestId, Slot)

                Case e_OBJType.otMagicos

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
1350    LogError "Error en useinvitem Usuario: " & UserList(UserIndex).name & " item:" & obj.name & " index: " & UserList(UserIndex).invent.Object(Slot).ObjIndex

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)

        On Error GoTo EnivarArmasConstruibles_Err

100     Call WriteBlacksmithWeapons(UserIndex)
        Exit Sub
EnivarArmasConstruibles_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarArmasConstruibles", Erl)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)

        On Error GoTo EnivarObjConstruibles_Err

100     Call WriteCarpenterObjects(UserIndex)
        Exit Sub
EnivarObjConstruibles_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruibles", Erl)

End Sub

Sub EnivarObjConstruiblesAlquimia(ByVal UserIndex As Integer)

        On Error GoTo EnivarObjConstruiblesAlquimia_Err

100     Call WriteAlquimistaObjects(UserIndex)
        Exit Sub
EnivarObjConstruiblesAlquimia_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruiblesAlquimia", Erl)

End Sub

Sub EnivarObjConstruiblesSastre(ByVal UserIndex As Integer)

        On Error GoTo EnivarObjConstruiblesSastre_Err

100     Call WriteSastreObjects(UserIndex)
        Exit Sub
EnivarObjConstruiblesSastre_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruiblesSastre", Erl)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)

        On Error GoTo EnivarArmadurasConstruibles_Err

100     Call WriteBlacksmithArmors(UserIndex)
        Exit Sub
EnivarArmadurasConstruibles_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarArmadurasConstruibles", Erl)

End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean

        On Error GoTo ItemSeCae_Err

100     ItemSeCae = (ObjData(Index).Real <> 1 Or ObjData(Index).NoSeCae = 0) And (ObjData(Index).Caos <> 1 Or ObjData(Index).NoSeCae = 0) And ObjData(Index).OBJType <> e_OBJType.otLlaves And ObjData(Index).OBJType <> e_OBJType.otBarcos And ObjData(Index).OBJType <> e_OBJType.otMonturas And ObjData(Index).NoSeCae = 0 And Not ObjData(Index).Intirable = 1 And Not ObjData(Index).Destruye = 1 And Not ObjData(Index).Instransferible = 1
        Exit Function
ItemSeCae_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.ItemSeCae", Erl)

End Function

Public Function PirataCaeItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

        On Error GoTo PirataCaeItem_Err

100     With UserList(UserIndex)

102         If .clase = e_Class.Pirat And .Stats.ELV >= 37 And .flags.Navegando = 1 Then

                ' Si no está navegando, se caen los items
104             If .invent.BarcoObjIndex > 0 Then

                    ' Con galeón cada item tiene una probabilidad de caerse del 67%
106                 If ObjData(.invent.BarcoObjIndex).Ropaje = iGaleon Then
108                     If RandomNumber(1, 100) <= 33 Then
                            Exit Function

                        End If

                    End If

                End If

            End If

        End With

110     PirataCaeItem = True
        Exit Function
PirataCaeItem_Err:
112     Call TraceError(Err.Number, Err.Description, "InvUsuario.PirataCaeItem", Erl)

End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)

        On Error GoTo TirarTodosLosItems_Err

        Dim i         As Byte
        Dim NuevaPos  As t_WorldPos
        Dim MiObj     As t_Obj
        Dim ItemIndex As Integer

100     With UserList(UserIndex)

            If ((.pos.Map = 58 Or .pos.Map = 59 Or .pos.Map = 60 Or .pos.Map = 61) And EnEventoFaccionario) Then Exit Sub

            ' Tambien se cae el oro de la billetera
            Dim GoldToDrop As Long
            GoldToDrop = .Stats.GLD - (SvrConfig.GetValue("OroPorNivelBilletera") * .Stats.ELV)

102         If GoldToDrop > 0 And Not EsGM(UserIndex) Then
104             Call TirarOro(GoldToDrop, UserIndex)

            End If

106         For i = 1 To .CurrentInventorySlots
108             ItemIndex = .invent.Object(i).ObjIndex

110             If ItemIndex > 0 Then
112                 If ItemSeCae(ItemIndex) And PirataCaeItem(UserIndex, i) And (Not EsNewbie(UserIndex) Or Not ItemNewbie(ItemIndex)) Then
114                     NuevaPos.x = 0
116                     NuevaPos.y = 0
118                     MiObj.amount = DropAmmount(.invent, i)
120                     MiObj.ObjIndex = ItemIndex

                        If .flags.Navegando Then
128                         Call Tilelibre(.pos, NuevaPos, MiObj, True, True)
                        Else
129                         Call Tilelibre(.pos, NuevaPos, MiObj, .flags.Navegando = True, (Not .flags.Navegando) = True)
                            Call ClosestLegalPos(.pos, NuevaPos, .flags.Navegando, Not .flags.Navegando)

                        End If

130                     If NuevaPos.x <> 0 And NuevaPos.y <> 0 Then
132                         Call DropObj(UserIndex, i, MiObj.amount, NuevaPos.Map, NuevaPos.x, NuevaPos.y)
                            ' WyroX: Si no hay lugar, quemamos el item del inventario (nada de mochilas gratis)
                        Else
134                         Call QuitarUserInvItem(UserIndex, i, MiObj.amount)
136                         Call UpdateUserInv(False, UserIndex, i)

                        End If

                    End If

                End If

138         Next i

        End With

        Exit Sub
TirarTodosLosItems_Err:
140     Call TraceError(Err.Number, Err.Description, "InvUsuario.TirarTodosLosItems", Erl)

End Sub

Function DropAmmount(ByRef invent As t_Inventario, _
                     ByVal objectIndex As Integer) As Integer
100     DropAmmount = invent.Object(objectIndex).amount

102     If invent.MagicoObjIndex > 0 Then

            With ObjData(invent.MagicoObjIndex)

104             If .EfectoMagico = 12 Then

                    Dim unprotected As Single
                    unprotected = 1

106                 If invent.Object(objectIndex).ObjIndex = ORO_MINA Then 'ore types
108                     unprotected = CSng(1) - (CSng(.LingO) / 100)
110                 ElseIf invent.Object(objectIndex).ObjIndex = PLATA_MINA Then
112                     unprotected = CSng(1) - (CSng(.LingP) / 100)
114                 ElseIf invent.Object(objectIndex).ObjIndex = HIERRO_MINA Then
116                     unprotected = CSng(1) - (CSng(.LingH) / 100)
118                 ElseIf invent.Object(objectIndex).ObjIndex = Wood Then ' wood types
120                     unprotected = CSng(1) - (CSng(.Madera) / 100)
122                 ElseIf invent.Object(objectIndex).ObjIndex = ElvenWood Then
124                     unprotected = CSng(1) - (CSng(.MaderaElfica) / 100)
129                 ElseIf invent.Object(objectIndex).ObjIndex = PinoWood Then
130                     unprotected = CSng(1) - (CSng(.MaderaPino) / 100)
131                 ElseIf invent.Object(objectIndex).ObjIndex = BLODIUM_MINA Then
132                     unprotected = CSng(1) - (CSng(.Blodium) / 100)
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

100     ItemNewbie = ObjData(ItemIndex).Newbie = 1
        Exit Function
ItemNewbie_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.ItemNewbie", Erl)

End Function

Public Function IsItemInCooldown(ByRef user As t_User, ByRef obj As t_UserOBJ) As Boolean

    Dim ElapsedTime As Long
    ElapsedTime = GetTickCount() - user.CdTimes(ObjData(obj.ObjIndex).cdType)
    IsItemInCooldown = ElapsedTime < ObjData(obj.ObjIndex).Cooldown

End Function

Public Sub UserTargetableItem(ByVal UserIndex As Integer, _
                              ByVal TileX As Integer, _
                              ByVal TileY As Integer)

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

100         Dim CanHelpResult As e_InteractionResult

102         If Not IsValidUserRef(.flags.TargetUser) Then
104             Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            If .flags.TargetUser.ArrayIndex = UserIndex Then
                Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

114         Dim TargetUser As Integer
116         TargetUser = .flags.TargetUser.ArrayIndex

            If UserList(TargetUser).flags.Muerto = 0 Then
                Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

106         CanHelpResult = UserMod.CanHelpUser(UserIndex, TargetUser)

            If UserList(TargetUser).flags.SeguroResu Then
                ' Msg693=El usuario tiene el seguro de resurrección activado.
                Call WriteLocaleMsg(UserIndex, "693", e_FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(TargetUser, UserList(UserIndex).name & " está intentando revivirte. Desactiva el seguro de resurrección para permitirle hacerlo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            If CanHelpResult <> eInteractionOk Then
                Call SendHelpInteractionMessage(UserIndex, CanHelpResult)

            End If

118         Dim costoVidaResu As Long
120         costoVidaResu = UserList(TargetUser).Stats.ELV * 1.5 + .Stats.MinHp * 0.5
122         Call UserMod.ModifyHealth(UserIndex, -costoVidaResu, 1)
124         Call ModifyStamina(UserIndex, -UserList(UserIndex).Stats.MinSta, False, 0)

            Dim ObjIndex As Integer
126         ObjIndex = .invent.Object(.flags.UsingItemSlot).ObjIndex
128         Call UpdateCd(UserIndex, ObjData(ObjIndex).cdType)
192         Call RemoveItemFromInventory(UserIndex, UserList(UserIndex).flags.UsingItemSlot)
196         Call ResurrectUser(TargetUser)

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

Public Sub PlaceTrap(ByVal UserIndex As Integer, _
                     ByVal TileX As Integer, _
                     ByVal TileY As Integer)

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

100         Dim CanAttackResult As e_AttackInteractionResult
            Dim TargetRef       As t_AnyReference

            If IsValidUserRef(.flags.TargetUser) Then
                Call CastUserToAnyRef(.flags.TargetUser, TargetRef)
            Else
                Call CastNpcToAnyRef(.flags.TargetNPC, TargetRef)

            End If

102         If Not IsValidRef(TargetRef) Then
104             Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
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
                Call SendData(SendTarget.ToPCAliveArea, TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(TargetRef.ArrayIndex).Char.charindex, FXSANGRE, 0, UserList(TargetRef.ArrayIndex).pos.x, UserList(TargetRef.ArrayIndex).pos.y))
                Call SendData(SendTarget.ToPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(TargetRef.ArrayIndex).pos.x, UserList(TargetRef.ArrayIndex).pos.y))
            Else

                If NpcList(TargetRef.ArrayIndex).flags.Snd2 > 0 Then
                    Call SendData(SendTarget.ToNPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(NpcList(TargetRef.ArrayIndex).flags.Snd2, NpcList(TargetRef.ArrayIndex).pos.x, NpcList(TargetRef.ArrayIndex).pos.y))
                Else
                    Call SendData(SendTarget.ToNPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(SND_IMPACTO2, NpcList(TargetRef.ArrayIndex).pos.x, NpcList(TargetRef.ArrayIndex).pos.y))

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

Public Sub UseHandCannon(ByVal UserIndex As Integer, _
                         ByVal TileX As Integer, _
                         ByVal TileY As Integer)

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
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, Particula, Tiempo, False, , UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
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
                'TODO place ship body
                Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                Exit Sub

            End If

            If .flags.Muerto = 1 Then
                .Char.body = iCuerpoMuerto
204             .Char.head = 0
206             .Char.ShieldAnim = NingunEscudo
208             .Char.WeaponAnim = NingunArma
210             .Char.CascoAnim = NingunCasco
211             .Char.CartAnim = NoCart
                Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                Exit Sub

            End If

            .Char.head = .OrigChar.head

            If .invent.WeaponEqpObjIndex > 0 Then
                .Char.WeaponAnim = ObjData(.invent.WeaponEqpObjIndex).WeaponAnim
            ElseIf .invent.HerramientaEqpObjIndex > 0 Then
                .Char.WeaponAnim = ObjData(.invent.HerramientaEqpObjIndex).WeaponAnim
            Else
                .Char.WeaponAnim = 0

            End If

            If .invent.ArmourEqpObjIndex > 0 Then
                .Char.body = ObtenerRopaje(UserIndex, ObjData(.invent.ArmourEqpObjIndex))
            Else
                Call SetNakedBody(UserList(UserIndex))

            End If

            If .invent.CascoEqpObjIndex > 0 Then
                .Char.CascoAnim = ObjData(.invent.CascoEqpObjIndex).CascoAnim
            Else
                .Char.CascoAnim = 0

            End If

            If .invent.MagicoObjIndex > 0 Then
                .Char.CartAnim = ObjData(.invent.MagicoObjIndex).Ropaje
            Else
                .Char.CartAnim = 0

            End If

            If .invent.EscudoEqpObjIndex > 0 Then
                .Char.ShieldAnim = ObjData(.invent.EscudoEqpObjIndex).ShieldAnim
            Else
                .Char.ShieldAnim = 0

            End If

            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

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
