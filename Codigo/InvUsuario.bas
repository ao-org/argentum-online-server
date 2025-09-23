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

Public Function IsObjecIndextInInventory(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo IsObjecIndextInInventory_Err
    Debug.Assert UserIndex >= LBound(UserList) And UserIndex <= UBound(UserList)
    ' If no match is found, return False
    IsObjecIndextInInventory = False
    Dim i As Integer
    Dim maxItemsInventory As Integer
    Dim currentObjIndex As Integer
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
    Dim i As Integer
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
100         For i = 1 To UserList(UserIndex).CurrentInventorySlots
102             ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    
104             If ObjIndex > 0 Then
106                 If (ObjData(ObjIndex).OBJType <> e_OBJType.otKeys And ObjData(ObjIndex).OBJType <> e_OBJType.otShips And ObjData(ObjIndex).OBJType <> e_OBJType.otSaddles And ObjData(ObjIndex).OBJType <> e_OBJType.otDonator And ObjData(ObjIndex).OBJType <> e_OBJType.otRecallStones) Then
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

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional Slot As Byte) As Boolean

        On Error GoTo manejador

        Dim flag As Boolean

100     If Slot <> 0 Then
102         If UserList(UserIndex).Invent.Object(Slot).Equipped Then
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

Function RazaPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional Slot As Byte) As Boolean
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
        
        If RazaPuedeUsarItem And Objeto.OBJType = e_OBJType.otArmor Then
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
    
102             If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
                 
104                 If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then
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
                    UserList(UserIndex).BancoInvent.Object(j).ElementalTags = 0
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
        End With
        
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
132                     AuxPos = TirarItemAlPiso(.Pos, MiObj, False)
                    Else
134                     AuxPos = TirarItemAlPiso(.Pos, MiObj, True)
                    End If
            
136                 If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
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
162 Call TraceError(Err.Number, Err.Description, "InvUsuario.TirarOro", Erl())
    
End Sub

Public Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarUserInvItem_Err
        

100     If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
    
102     With UserList(UserIndex).Invent.Object(Slot)

104         If .amount <= Cantidad And .Equipped = 1 Then
106             Call Desequipar(UserIndex, Slot)
            End If
        
            'Quita un objeto
108         .amount = .amount - Cantidad

            '¿Quedan mas?
110         If .amount <= 0 Then
112             UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
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

Public Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
 On Error GoTo UpdateUserInv_Err

        Dim NullObj As t_UserOBJ
        Dim LoopC   As Byte

        'Actualiza un solo slot
100     If Not UpdateAll And Slot > 0 Then
    
            'Actualiza el inventario
102         If UserList(UserIndex).Invent.Object(Slot).ObjIndex > 0 Then
104             Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
            Else
106             Call ChangeUserInv(UserIndex, Slot, NullObj)
            End If
                        
            UserList(UserIndex).flags.ModificoInventario = True
        Else

            'Actualiza todos los slots
            If UserList(UserIndex).CurrentInventorySlots > 0 Then
108             For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
                    'Actualiza el inventario
                    If LoopC > 0 And LoopC <= UBound(UserList(UserIndex).invent.Object) Then 'Make sure the slot is valid
                      If UserList(UserIndex).invent.Object(LoopC).ObjIndex > 0 Then
                         Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).invent.Object(LoopC))
                      Else
                         Call ChangeUserInv(UserIndex, LoopC, NullObj)
                      End If
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
            ByVal X As Integer, _
            ByVal Y As Integer)
        
        On Error GoTo DropObj_Err
        Dim obj As t_Obj

100     If num > 0 Then
102         With UserList(UserIndex)
104             If num > .Invent.Object(Slot).amount Then
106                 num = .Invent.Object(Slot).amount
                End If
108             obj.ObjIndex = .Invent.Object(Slot).ObjIndex
110             obj.amount = num
                obj.ElementalTags = .Invent.Object(Slot).ElementalTags
                If Not CustomScenarios.UserCanDropItem(UserIndex, Slot, Map, x, y) Then
                    Exit Sub
                End If
                
112             If ObjData(obj.ObjIndex).Destruye = 0 Then
                    Dim Suma As Long
                    Suma = num + MapData(.Pos.Map, X, Y).ObjInfo.amount
                    'Check objeto en el suelo
114                 If MapData(.Pos.Map, x, y).ObjInfo.ObjIndex = 0 Or (MapData(.Pos.Map, x, y).ObjInfo.ObjIndex = obj.ObjIndex And MapData(.Pos.Map, x, y).ObjInfo.ElementalTags = obj.ElementalTags And Suma <= MAX_INVENTORY_OBJS) Then

116                     If Suma > MAX_INVENTORY_OBJS Then
118                         num = MAX_INVENTORY_OBJS - MapData(.Pos.Map, X, Y).ObjInfo.amount
                        End If
                        ' Si sos Admin, Dios o Usuario, crea el objeto en el piso.
120                     If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Admin Or e_PlayerType.Dios)) <> 0 Then
                            ' Tiramos el item al piso
122                         Call MakeObj(obj, Map, X, Y)
                        End If
                        Call CustomScenarios.UserDropItem(UserIndex, Slot, Map, x, y)
124                     Call QuitarUserInvItem(UserIndex, Slot, num)
126                     Call UpdateUserInv(False, UserIndex, Slot)

                        If .flags.jugando_captura = 1 Then
                            If Not InstanciaCaptura Is Nothing Then
                                Call InstanciaCaptura.tiraBandera(UserIndex, obj.objIndex)
                            End If
                        End If
                        
128                     If Not .flags.Privilegios And e_PlayerType.user Then
                            If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) <> 0 Then
130                             Call LogGM(.Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).Name)
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

Sub EraseObj(ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo EraseObj_Err
        

        Dim Rango As Byte

100     MapData(Map, X, Y).ObjInfo.amount = MapData(Map, X, Y).ObjInfo.amount - num

102     If MapData(Map, X, Y).ObjInfo.amount <= 0 Then

            
108         MapData(Map, X, Y).ObjInfo.ObjIndex = 0
110         MapData(Map, X, Y).ObjInfo.amount = 0
            MapData(Map, x, y).ObjInfo.ElementalTags = 0
    
    
112         Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectDelete(X, Y))

        End If

        
        Exit Sub

EraseObj_Err:
114     Call TraceError(Err.Number, Err.Description, "InvUsuario.EraseObj", Erl)

        
End Sub

Sub MakeObj(ByRef obj As t_Obj, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal Limpiar As Boolean = True)
        
        On Error GoTo MakeObj_Err

        Dim Color As Long

        Dim Rango As Byte

100     If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then
    
102         If MapData(Map, x, y).ObjInfo.ObjIndex = obj.ObjIndex And MapData(Map, x, y).ObjInfo.ElementalTags = obj.ElementalTags Then
104             MapData(Map, X, Y).ObjInfo.amount = MapData(Map, X, Y).ObjInfo.amount + obj.amount
            Else
110             MapData(Map, X, Y).ObjInfo.ObjIndex = obj.ObjIndex
                MapData(Map, x, y).ObjInfo.ElementalTags = obj.ElementalTags

112             If ObjData(obj.ObjIndex).VidaUtil <> 0 Then
114                 MapData(Map, X, Y).ObjInfo.amount = ObjData(obj.ObjIndex).VidaUtil
                Else
116                 MapData(Map, X, Y).ObjInfo.amount = obj.amount

                End If
                
            End If
            
118         Call modSendData.SendToAreaByPos(Map, x, y, PrepareMessageObjectCreate(obj.ObjIndex, MapData(Map, x, y).ObjInfo.amount, x, y, MapData(Map, x, y).ObjInfo.ElementalTags))
    
        End If
        
        Exit Sub

MakeObj_Err:
120     Call TraceError(Err.Number, Err.Description, "InvUsuario.MakeObj", Erl)
End Sub

Function GetSlotForItemInInventory(ByVal UserIndex As Integer, ByRef MyObject As t_Obj) As Integer
On Error GoTo GetSlotForItemInInventory_Err
    GetSlotForItemInInventory = -1
100 Dim i As Integer
    
102 For i = 1 To UserList(UserIndex).CurrentInventorySlots
104    If UserList(UserIndex).invent.Object(i).objIndex = 0 And GetSlotForItemInInventory = -1 Then
106        GetSlotForItemInInventory = i 'we found a valid place but keep looking in case we can stack
108    ElseIf UserList(UserIndex).invent.Object(i).objIndex = MyObject.objIndex And _
              UserList(UserIndex).invent.Object(i).ElementalTags = MyObject.ElementalTags And _
              UserList(UserIndex).invent.Object(i).amount + MyObject.amount <= MAX_INVENTORY_OBJS Then
110        GetSlotForItemInInventory = i 'we can stack the item, let use this slot
112        Exit Function
       End If
    Next i
    Exit Function
GetSlotForItemInInventory_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.GetSlotForItemInInventory", Erl)
End Function

Function GetSlotInInventory(ByVal UserIndex As Integer, ByVal objIndex As Integer) As Integer
    On Error GoTo GetSlotInInventory_Err
    GetSlotInInventory = -1
    Dim i As Integer
    
    For i = 1 To UserList(UserIndex).CurrentInventorySlots
        If UserList(UserIndex).invent.Object(i).objIndex = objIndex Then
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

        Dim X    As Integer

        Dim Y    As Integer

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
118        Call WriteLocaleMsg(UserIndex, MsgInventoryIsFull, e_FontTypeNames.FONTTYPE_FIGHT)
120        MeterItemEnInventario = False
           Exit Function
        End If
        If UserList(UserIndex).invent.Object(Slot).objIndex = 0 Then
            UserList(UserIndex).invent.NroItems = UserList(UserIndex).invent.NroItems + 1
        End If
        'Mete el objeto
124     If UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
126         UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
128         UserList(UserIndex).Invent.Object(Slot).amount = UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount
            UserList(UserIndex).invent.Object(Slot).ElementalTags = MiObj.ElementalTags
            
        
        Else
130         UserList(UserIndex).Invent.Object(Slot).amount = MAX_INVENTORY_OBJS
        End If
        
132     Call UpdateUserInv(False, UserIndex, Slot)
        
134     MeterItemEnInventario = True
        UserList(UserIndex).flags.ModificoInventario = True

        Exit Function
MeterItemEnInventario_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.MeterItemEnInventario", Erl)
End Function

Function HayLugarEnInventario(ByVal UserIndex As Integer, ByVal TargetItemIndex As Integer, ByVal ItemCount) As Boolean
On Error GoTo HayLugarEnInventario_err
        Dim X    As Integer
        Dim Y    As Integer
        Dim Slot As Byte
100     Slot = 1

102     Do Until UserList(UserIndex).invent.Object(Slot).ObjIndex = 0 Or _
            (UserList(UserIndex).invent.Object(Slot).ObjIndex = TargetItemIndex And UserList(UserIndex).invent.Object(Slot).amount + ItemCount < 10000)
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
        
        Dim X    As Integer
        Dim Y    As Integer
        Dim Slot As Byte
        Dim obj   As t_ObjData
        Dim MiObj As t_Obj

        '¿Hay algun obj?
100     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then

            '¿Esta permitido agarrar este obj?
102         If ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then

104             If UserList(UserIndex).flags.Montado = 1 Then
106                 ' Msg672=Debes descender de tu montura para agarrar objetos del suelo.
                    Call WriteLocaleMsg(UserIndex, "672", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If Not UserCanPickUpItem(UserIndex) Then
                    Exit Sub
                End If
108             X = UserList(UserIndex).Pos.X
110             Y = UserList(UserIndex).Pos.Y

                If UserList(UserIndex).flags.jugando_captura = 1 Then
                    If Not InstanciaCaptura Is Nothing Then
                        If Not InstanciaCaptura.tomaBandera(UserIndex, MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.objIndex) Then
                            Exit Sub
                        End If
                    End If
                End If
                
                
112             obj = ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex)
114             MiObj.amount = MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.amount
116             MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.ObjIndex
117             MiObj.ElementalTags = MapData(UserList(UserIndex).Pos.Map, x, y).ObjInfo.ElementalTags
        
118             If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    'Call WriteConsoleMsg(UserIndex, "No puedo cargar mas objetos.", e_FontTypeNames.FONTTYPE_INFO)
                Else
            
                    'Quitamos el objeto
120                 Call EraseObj(MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

122                 If Not UserList(UserIndex).flags.Privilegios And e_PlayerType.user Then Call LogGM(UserList(UserIndex).Name, "Agarro:" & MiObj.amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
                    
                    'Si el obj es oro (12), se muestra la cantidad que agarro arriba del personaje
                    If MiObj.ObjIndex = 12 Then
                        Call WriteTextOverTile(UserIndex, "+" & PonerPuntos(MiObj.amount), UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, RGB(212, 175, 55))
                    End If
                    
                    Call UserDidPickupItem(UserIndex, MiObj.ObjIndex)
                    If UserList(UserIndex).flags.jugando_captura = 1 Then
                    If Not InstanciaCaptura Is Nothing Then
                            Call InstanciaCaptura.quitarBandera(UserIndex, MiObj.objIndex)
                    End If
                    End If
    
124                 If BusquedaTesoroActiva Then
126                     If UserList(UserIndex).Pos.Map = TesoroNumMapa And UserList(UserIndex).Pos.X = TesoroX And UserList(UserIndex).Pos.Y = TesoroY Then
    
128                         Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1639, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_TALK)) 'Msg1640=Eventos> ¬1 encontró el tesoro ¡Felicitaciones!
130                         BusquedaTesoroActiva = False

                        End If

                    End If
                
132                 If BusquedaRegaloActiva Then
134                     If UserList(UserIndex).Pos.Map = RegaloNumMapa And UserList(UserIndex).Pos.X = RegaloX And UserList(UserIndex).Pos.Y = RegaloY Then
136                         Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1640, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_TALK)) 'Msg1640=Eventos> ¬1 fue el valiente que encontró el gran ítem mágico ¡Felicitaciones!
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

    If (Slot < LBound(UserList(UserIndex).invent.Object)) Or (Slot > UBound(UserList(UserIndex).invent.Object)) Then
        Exit Sub
    ElseIf UserList(UserIndex).invent.Object(Slot).ObjIndex = 0 Then
        Exit Sub
    End If

    obj = ObjData(UserList(UserIndex).invent.Object(Slot).ObjIndex)

    Select Case obj.OBJType

        Case e_OBJType.otWeapon
            UserList(UserIndex).invent.Object(Slot).Equipped = 0
            UserList(UserIndex).invent.EquippedWeaponObjIndex = 0
            UserList(UserIndex).invent.EquippedWeaponSlot = 0
            UserList(UserIndex).Char.Arma_Aura = ""
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 1))
        
            UserList(UserIndex).Char.WeaponAnim = NingunArma
            
            If UserList(UserIndex).flags.Montado = 0 Then
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)
            End If
                
            If obj.MagicDamageBonus > 0 Then
                Call WriteUpdateDM(UserIndex)
            End If
    
        Case e_OBJType.otArrows
            UserList(UserIndex).invent.Object(Slot).Equipped = 0
            UserList(UserIndex).invent.EquippedMunitionObjIndex = 0
            UserList(UserIndex).invent.EquippedMunitionSlot = 0
    
            ' Case e_OBJType.otAnillos
            '    UserList(UserIndex).Invent.Object(slot).Equipped = 0
            '    UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
            ' UserList(UserIndex).Invent.AnilloEqpSlot = 0
            
        Case e_OBJType.otWorkingTools
            If UserList(UserIndex).flags.PescandoEspecial = False Then
                UserList(UserIndex).invent.Object(Slot).Equipped = 0
                UserList(UserIndex).invent.EquippedWorkingToolObjIndex = 0
                UserList(UserIndex).invent.EquippedWorkingToolSlot = 0
    
                If UserList(UserIndex).flags.UsandoMacro = True Then
                    Call WriteMacroTrabajoToggle(UserIndex, False)
                End If
            
                UserList(UserIndex).Char.WeaponAnim = NingunArma
                
                If UserList(UserIndex).flags.Montado = 0 Then
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)
                End If
            End If
       
        Case e_OBJType.otAmulets
    
            Select Case obj.EfectoMagico

                Case e_MagicItemEffect.eModifyAttributes
                    If obj.QueAtributo <> 0 Then
                        UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
                        UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) - obj.CuantoAumento
                        ' UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
                            
                        Call WriteFYA(UserIndex)
                    End If

                Case e_MagicItemEffect.eModifySkills
                    If obj.Que_Skill <> 0 Then
                        UserList(UserIndex).Stats.UserSkills(obj.Que_Skill) = UserList(UserIndex).Stats.UserSkills(obj.Que_Skill) - obj.CuantoAumento
                    End If
                        
                Case e_MagicItemEffect.eRegenerateHealth
                    UserList(UserIndex).flags.RegeneracionHP = 0

                Case e_MagicItemEffect.eRegenerateMana
                    UserList(UserIndex).flags.RegeneracionMana = 0

                Case e_MagicItemEffect.eIncreaseDamageToNpc
                    UserList(UserIndex).Stats.MaxHit = UserList(UserIndex).Stats.MaxHit - obj.CuantoAumento
                    UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT - obj.CuantoAumento

                Case e_MagicItemEffect.eInmunityToNpcMagic 'Orbe ignea
                    UserList(UserIndex).flags.NoMagiaEfecto = 0

                Case e_MagicItemEffect.eIncinerate
                    UserList(UserIndex).flags.incinera = 0

                Case e_MagicItemEffect.eParalize
                    UserList(UserIndex).flags.Paraliza = 0

                Case e_MagicItemEffect.eProtectedResources
                    If UserList(UserIndex).flags.Muerto = 0 Then
                        UserList(UserIndex).Char.CartAnim = NoCart
                        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)
                    End If
                        
                Case e_MagicItemEffect.eProtectedInventory
                    UserList(UserIndex).flags.PendienteDelSacrificio = 0
                 
                Case e_MagicItemEffect.ePreventMagicWords
                    UserList(UserIndex).flags.NoPalabrasMagicas = 0

                Case e_MagicItemEffect.ePreventInvisibleDetection
                    UserList(UserIndex).flags.NoDetectable = 0

                Case e_MagicItemEffect.eIncreaseLearningSkills
                    UserList(UserIndex).flags.PendienteDelExperto = 0

                Case e_MagicItemEffect.ePoison
                    UserList(UserIndex).flags.Envenena = 0

                Case e_MagicItemEffect.eRingOfShadows
                    UserList(UserIndex).flags.AnilloOcultismo = 0

                Case e_MagicItemEffect.eTalkToDead
                    Call UnsetMask(UserList(UserIndex).flags.StatusMask, e_StatusMask.eTalkToDead)
                    ' Msg673=Dejas el mundo de los muertos, ya no podrás comunicarte con ellos.
                    Call WriteLocaleMsg(UserIndex, "673", e_FontTypeNames.FONTTYPE_WARNING)
                    Call SendData(SendTarget.ToPCDeadAreaButIndex, UserIndex, PrepareMessageCharacterRemove(4, UserList(UserIndex).Char.charindex, False, True))
            End Select
        
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 5))
            UserList(UserIndex).Char.Otra_Aura = 0
            UserList(UserIndex).invent.Object(Slot).Equipped = 0
            UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex = 0
            UserList(UserIndex).invent.EquippedAmuletAccesorySlot = 0
        
        Case e_OBJType.otArmor
            UserList(UserIndex).invent.Object(Slot).Equipped = 0
            UserList(UserIndex).invent.EquippedArmorObjIndex = 0
            UserList(UserIndex).invent.EquippedArmorSlot = 0
        
            If UserList(UserIndex).flags.Navegando = 0 Then
                If UserList(UserIndex).flags.Montado = 0 Then
                    Call SetNakedBody(UserList(UserIndex))
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)
                End If
            End If
        
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 2))
        
            UserList(UserIndex).Char.Body_Aura = 0

            If obj.ResistenciaMagica > 0 Then
                Call WriteUpdateRM(UserIndex)
            End If
    
        Case e_OBJType.otHelmet
            UserList(UserIndex).invent.Object(Slot).Equipped = 0
            UserList(UserIndex).invent.EquippedHelmetObjIndex = 0
            UserList(UserIndex).invent.EquippedHelmetSlot = 0
            UserList(UserIndex).Char.Head_Aura = 0
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 4))

            UserList(UserIndex).Char.CascoAnim = NingunCasco
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)
    
            If obj.ResistenciaMagica > 0 Then
                Call WriteUpdateRM(UserIndex)
            End If
    
        Case e_OBJType.otShield
            UserList(UserIndex).invent.Object(Slot).Equipped = 0
            UserList(UserIndex).invent.EquippedShieldObjIndex = 0
            UserList(UserIndex).invent.EquippedShieldSlot = 0
            UserList(UserIndex).Char.Escudo_Aura = 0
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 3))
        
            UserList(UserIndex).Char.ShieldAnim = NingunEscudo

            If UserList(UserIndex).flags.Montado = 0 Then
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)
            End If
                
            If obj.ResistenciaMagica > 0 Then
                Call WriteUpdateRM(UserIndex)
            End If
                
        Case e_OBJType.otAmulets
            UserList(UserIndex).invent.Object(Slot).Equipped = 0
            UserList(UserIndex).invent.EquippedAmuletAccesoryObjIndex = 0
            UserList(UserIndex).invent.EquippedAmuletAccesorySlot = 0
            UserList(UserIndex).Char.DM_Aura = 0
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 6))
            Call WriteUpdateDM(UserIndex)
            Call WriteUpdateRM(UserIndex)
                
        Case e_OBJType.otRingAccesory, e_OBJType.otMagicalInstrument
            UserList(UserIndex).invent.Object(Slot).Equipped = 0
            UserList(UserIndex).invent.EquippedRingAccesoryObjIndex = 0
            UserList(UserIndex).invent.EquippedRingAccesorySlot = 0
            UserList(UserIndex).Char.RM_Aura = 0
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 7))
            Call WriteUpdateRM(UserIndex)
            Call WriteUpdateDM(UserIndex)
                
                
        Case e_OBJType.otBackpack
        
            UserList(UserIndex).invent.Object(Slot).Equipped = 0
            UserList(UserIndex).invent.EquippedBackpackObjIndex = 0
            UserList(UserIndex).invent.EquippedBackpackSlot = 0
            UserList(UserIndex).Char.BackpackAnim = 0
                
            If UserList(UserIndex).flags.Navegando = 0 Then
                If UserList(UserIndex).flags.Montado = 0 Then
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)
                End If
            End If
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 2))
            UserList(UserIndex).Char.Body_Aura = 0
        
    End Select
        
    Call UpdateUserInv(False, UserIndex, Slot)

        
    Exit Sub

Desequipar_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.Desequipar", Erl)

        
End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

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

Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
        
        On Error GoTo FaccionPuedeUsarItem_Err
        
100     If EsGM(UserIndex) Then
102         FaccionPuedeUsarItem = True
            Exit Function
        End If
        
104     If ObjIndex < 1 Then Exit Function

106     If ObjData(ObjIndex).Real = 1 Then
107         If ObjData(ObjIndex).LeadersOnly Then
108             FaccionPuedeUsarItem = (Status(UserIndex) = e_Facciones.consejo)
109         ElseIf Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
110             FaccionPuedeUsarItem = esArmada(UserIndex)
            Else
112             FaccionPuedeUsarItem = False
            End If

114     ElseIf ObjData(ObjIndex).Caos = 1 Then
115         If ObjData(ObjIndex).LeadersOnly Then
116             FaccionPuedeUsarItem = (Status(UserIndex) = e_Facciones.concilio)
117         ElseIf Status(UserIndex) = e_Facciones.Caos Or Status(UserIndex) = e_Facciones.concilio Then
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

Function JerarquiaPuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean
       
    With UserList(userindex)
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
            If .invent.EquippedShipObjIndex <= 0 Or .invent.EquippedShipObjIndex > UBound(ObjData) Then Exit Sub
102         Barco = ObjData(.invent.EquippedShipObjIndex)

104         If .flags.Muerto = 1 Then
106             If Barco.Ropaje = iTraje Or Barco.Ropaje = iTrajeAltoNw Or Barco.Ropaje = iTrajeBajoNw Then
                    ' No tenemos la cabeza copada que va con iRopaBuceoMuerto,
                    ' asique asignamos el casper directamente caminando sobre el agua.
108                 .Char.Body = iCuerpoMuerto 'iRopaBuceoMuerto
110                 .Char.Head = iCabezaMuerto
                ElseIf Barco.Ropaje = iTrajeAltoNw Then
          
                ElseIf Barco.Ropaje = iTrajeBajoNw Then
          
                Else
112                 .Char.Body = iFragataFantasmal
114                 .Char.Head = 0
                End If
      
            Else ' Esta vivo

116             If Barco.Ropaje = iTraje Then
118                 .Char.Body = iTraje
120                 .Char.Head = .OrigChar.Head

122                 If .invent.EquippedHelmetObjIndex > 0 Then
124                     .Char.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
                    End If
                ElseIf Barco.Ropaje = iTrajeAltoNw Then
                    .Char.Body = iTrajeAltoNw
                    .Char.Head = .OrigChar.Head

                    If .invent.EquippedHelmetObjIndex > 0 Then
                        .Char.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
                    End If
                ElseIf Barco.Ropaje = iTrajeBajoNw Then
                    .Char.Body = iTrajeBajoNw
                    .Char.Head = .OrigChar.Head

                    If .invent.EquippedHelmetObjIndex > 0 Then
                        .Char.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
                    End If
                Else
126                 .Char.Head = 0
128                 .Char.CascoAnim = NingunCasco
                End If

130             If .Faccion.status = e_Facciones.Armada Or .Faccion.status = e_Facciones.consejo Then
132                 If Barco.Ropaje = iBarca Then .Char.Body = iBarcaArmada
134                 If Barco.Ropaje = iGalera Then .Char.Body = iGaleraArmada
136                 If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonArmada

138             ElseIf .Faccion.status = e_Facciones.Caos Or .Faccion.status = e_Facciones.concilio Then

140                 If Barco.Ropaje = iBarca Then .Char.Body = iBarcaCaos
142                 If Barco.Ropaje = iGalera Then .Char.Body = iGaleraCaos
144                 If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonCaos
          
                Else

146                 If Barco.Ropaje = iBarca Then .Char.Body = IIf(.Faccion.Status = 0, iBarcaCrimi, iBarcaCiuda)
148                 If Barco.Ropaje = iGalera Then .Char.Body = IIf(.Faccion.Status = 0, iGaleraCrimi, iGaleraCiuda)
150                 If Barco.Ropaje = iGaleon Then .Char.Body = IIf(.Faccion.Status = 0, iGaleonCrimi, iGaleonCiuda)
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
Sub EquiparInvItem(ByVal UserIndex As Integer, _
                   ByVal Slot As Byte, _
                   Optional ByVal UserIsLoggingIn As Boolean = False)

    On Error GoTo ErrHandler

    Dim obj       As t_ObjData

    Dim ObjIndex  As Integer

    Dim errordesc As String
    
    Dim Ropaje    As Integer
        
    ObjIndex = UserList(UserIndex).invent.Object(Slot).ObjIndex
    obj = ObjData(ObjIndex)
        
    If PuedeUsarObjeto(UserIndex, ObjIndex, True) > 0 Then

        Exit Sub

    End If

    With UserList(UserIndex)

        If .flags.Muerto = 1 Then
            'Msg77=¡¡Estás muerto!!.
            Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)

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
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)
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
            
                If obj.DosManos = 1 Then
                    If .invent.EquippedShieldObjIndex > 0 Then
                        Call Desequipar(UserIndex, .invent.EquippedShieldSlot)
                        ' Msg674=No puedes usar armas dos manos si tienes un escudo equipado. Tu escudo fue desequipado.
                        Call WriteLocaleMsg(UserIndex, "674", e_FontTypeNames.FONTTYPE_INFOIAO)
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
                        .Char.WeaponAnim = obj.WeaponAnim
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)

                    End If

                End If

            Case e_OBJType.otBackpack
                errordesc = "Backpack"
                
                If .invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    .Char.BackpackAnim = NoBackPack

                    If .flags.Montado = 0 And .flags.Navegando = 0 Then
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                    End If

                    Exit Sub

                End If

                If .invent.EquippedBackpackObjIndex > 0 Then
                    Call Desequipar(UserIndex, .invent.EquippedBackpackSlot)
                End If

                Ropaje = ObtenerRopaje(UserIndex, obj)

                If Ropaje = 0 Then
                    ' Msg676=Hay un error con este objeto. Infórmale a un administrador.
                    Call WriteLocaleMsg(UserIndex, "676", e_FontTypeNames.FONTTYPE_INFO)

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
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)
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
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)

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

                    Case e_MagicItemEffect.eModifyAttributes 'Modif la fuerza, agilidad, carisma, etc
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
                            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)
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
                
                'Quitamos el elemento anterior
                If .invent.EquippedMunitionObjIndex > 0 Then
                    Call Desequipar(UserIndex, .invent.EquippedMunitionSlot)
                End If
        
                .invent.Object(Slot).Equipped = 1
                .invent.EquippedMunitionObjIndex = .invent.Object(Slot).ObjIndex
                .invent.EquippedMunitionSlot = Slot

            Case e_OBJType.otArmor

                If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eArmor) Then
                    Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)

                    Exit Sub

                End If

                Ropaje = ObtenerRopaje(UserIndex, obj)

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
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)
                    Else
                        .flags.Desnudo = 1
                    End If

                    Exit Sub

                End If

                'Quita el anterior
                If .invent.EquippedArmorObjIndex > 0 Then
                    errordesc = "Armadura 2"
                    Call Desequipar(UserIndex, .invent.EquippedArmorSlot)
                    errordesc = "Armadura 3"

                End If
  
                'Si esta equipando armadura faccionaria fuera de zona segura o fuera de trigger seguro
                If Not UserIsLoggingIn Then
                    If obj.Real > 0 Or obj.Caos > 0 Then
                        If Not MapData(.pos.Map, .pos.x, .pos.y).trigger = e_Trigger.ZonaSegura And Not MapInfo(.pos.Map).Seguro = 1 Then
                            Call WriteLocaleMsg(UserIndex, "2091", e_FontTypeNames.FONTTYPE_INFO)

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

                If .flags.Montado = 0 And .flags.Navegando = 0 Then
                    .Char.body = Ropaje

                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)
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

                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)

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

                    Dim nuevoHead  As Integer

                    Dim nuevoCasco As Integer
                    
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
                    .Char.CascoAnim = nuevoCasco
                    
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
                        Call WriteLocaleMsg(UserIndex, "677", e_FontTypeNames.FONTTYPE_INFOIAO)
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
                    .Char.ShieldAnim = obj.ShieldAnim
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                End If

                If obj.ResistenciaMagica > 0 Then
                    Call WriteUpdateRM(UserIndex)
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
    Dim obj As t_ObjData
    For Index = 1 To UBound(inventory.Object)
        If Index <> Slot And inventory.Object(Index).Equipped Then
            If inventory.Object(Index).objIndex > 0 Then
                If inventory.Object(Index).objIndex > 0 Then
                    obj = ObjData(inventory.Object(Index).objIndex)
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

Dim obj                         As t_ObjData
Dim ObjIndex                    As Integer
Dim TargObj                     As t_ObjData
Dim MiObj                       As t_Obj

10  On Error GoTo hErr

    ' Agrego el Cuerno de la Armada y la Legión.
    'Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.

20  With UserList(UserIndex)

30      If .invent.Object(Slot).amount = 0 Then Exit Sub

40      If Not CanUseItem(.flags, .Counters) Then
50          Call WriteLocaleMsg(UserIndex, 395, e_FontTypeNames.FONTTYPE_INFO)
60          Exit Sub
70      End If

80      If PuedeUsarObjeto(UserIndex, .invent.Object(Slot).ObjIndex, True) > 0 Then
90          Exit Sub
100     End If

110     obj = ObjData(.invent.Object(Slot).ObjIndex)
120     Dim TimeSinceLastUse    As Long: TimeSinceLastUse = GetTickCount() - .CdTimes(obj.cdType)
130     If TimeSinceLastUse < obj.Cooldown Then Exit Sub

140     If IsSet(obj.ObjFlags, e_ObjFlags.e_UseOnSafeAreaOnly) Then
150         If MapInfo(.pos.Map).Seguro = 0 Then
                ' Msg678=Solo podes usar este objeto en mapas seguros.
160             Call WriteLocaleMsg(UserIndex, "678", e_FontTypeNames.FONTTYPE_INFO)
170             Exit Sub
180         End If
190     End If

200     If obj.OBJType = e_OBJType.otWeapon Then
210         If obj.Proyectil = 1 Then
                'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
220             If ByClick <> 0 Then
230                 If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
240             Else
250                 If Not IntervaloPermiteUsarClick(UserIndex) Then Exit Sub
260             End If
270         Else
                'dagas
280             If ByClick <> 0 Then
290                 If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
300             Else
310                 If Not IntervaloPermiteUsarClick(UserIndex) Then Exit Sub
320             End If
330         End If

340     Else
350         If ByClick <> 0 Then
360             If Not IntervaloPermiteUsarClick(UserIndex) Then Exit Sub
370         Else
380             If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
390         End If
400     End If

410     If .flags.Meditando Then
420         .flags.Meditando = False
430         .Char.FX = 0
440         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
450     End If

460     If obj.Newbie = 1 And Not EsNewbie(UserIndex) And Not EsGM(UserIndex) Then
            ' Msg679=Solo los newbies pueden usar estos objetos.
470         Call WriteLocaleMsg(UserIndex, "679", e_FontTypeNames.FONTTYPE_INFO)
480         Exit Sub
490     End If

500     If .Stats.ELV < obj.MinELV Then
510         Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1926, obj.MinELV, e_FontTypeNames.FONTTYPE_INFO))    ' Msg1926=Necesitas ser nivel ¬1 para usar este item.
520         Exit Sub
530     End If

540     If .Stats.ELV > obj.MaxLEV And obj.MaxLEV > 0 Then
550         Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1982, obj.MaxLEV, e_FontTypeNames.FONTTYPE_INFO))    ' Msg1982=Este objeto no puede ser utilizado por personajes de nivel ¬1 o superior.
560         Exit Sub
570     End If

580     ObjIndex = .invent.Object(Slot).ObjIndex
590     .flags.TargetObjInvIndex = ObjIndex
600     .flags.TargetObjInvSlot = Slot

610     Select Case obj.OBJType

            Case e_OBJType.otUseOnce
620             If .flags.Muerto = 1 Then
630                 Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
640                 Exit Sub
650             End If

                'Usa el item
660             .Stats.MinHam = .Stats.MinHam + obj.MinHam

670             If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
680             Call WriteUpdateHungerAndThirst(UserIndex)

690             If obj.JineteLevel > 0 Then
700                 If .Stats.JineteLevel < obj.JineteLevel Then
710                     .Stats.JineteLevel = obj.JineteLevel
720                 Else
                        'Msg2080 = No puedes consumir un nivel de jinete menor al que posees actualmente
730                     Call WriteLocaleMsg(UserIndex, 2079, e_FontTypeNames.FONTTYPE_INFO)
740                     Exit Sub
750                 End If
760             End If

                'Sonido
770             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.SOUND_COMIDA, .pos.x, .pos.y))

                'Quitamos del inv el item
780             Call QuitarUserInvItem(UserIndex, Slot, 1)
790             Call UpdateUserInv(False, UserIndex, Slot)
800             .flags.ModificoInventario = True

810         Case e_OBJType.otGoldCoin

820             If .flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
830                 Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
840                 Exit Sub

850             End If

860             .Stats.GLD = .Stats.GLD + .invent.Object(Slot).amount
870             .invent.Object(Slot).amount = 0
880             .invent.Object(Slot).ObjIndex = 0
890             .invent.NroItems = .invent.NroItems - 1
900             .flags.ModificoInventario = True
910             Call UpdateUserInv(False, UserIndex, Slot)
920             Call WriteUpdateGold(UserIndex)

930         Case e_OBJType.otWeapon

940             If .flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
950                 Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
960                 Exit Sub
970             End If

980             If Not .Stats.MinSta > 0 Then
990                 Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
1000                Exit Sub
1010            End If

1020            If ObjData(ObjIndex).Proyectil = 1 Then
1030                If IsSet(.flags.StatusMask, e_StatusMask.eTransformed) Then
1040                    Call WriteLocaleMsg(UserIndex, MsgCantUseBowTransformed, e_FontTypeNames.FONTTYPE_INFO)
1050                    Exit Sub
1060                End If
1070                Call WriteWorkRequestTarget(UserIndex, Proyectiles)
1080            Else
1090                If .flags.TargetObj = Wood Then
1100                    If .invent.Object(Slot).ObjIndex = DAGA Then
1110                        Call TratarDeHacerFogata(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY, UserIndex)
1120                    End If
1130                End If
1140            End If

1150            If .invent.Object(Slot).Equipped = 0 Then
1160                Exit Sub
1170            End If

1180        Case e_OBJType.otWorkingTools

1190            If .flags.Muerto = 1 Then
1200                Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
1210                Exit Sub
1220            End If

1230            If Not .Stats.MinSta > 0 Then
1240                Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
1250                Exit Sub
1260            End If

                'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
1270            If .invent.Object(Slot).Equipped = 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", e_FontTypeNames.FONTTYPE_INFO)
1280                Call WriteLocaleMsg(UserIndex, 376, e_FontTypeNames.FONTTYPE_INFO)
1290                Exit Sub
1300            End If

1310            Select Case obj.Subtipo
                    Case e_SubObjType.FishingRod, e_SubObjType.FishingNet    ' Herramientas del Pescador - Caña y Red
1320                    Call WriteWorkRequestTarget(UserIndex, e_Skill.Pescar)

1330                Case e_SubObjType.AlchemyScissors    ' Herramientas de Alquimia - Tijeras
1340                    Call WriteWorkRequestTarget(UserIndex, e_Skill.Alquimia)

1350                Case e_SubObjType.AlchemyPot    ' Herramientas de Alquimia - Olla
1360                    Call EnivarObjConstruiblesAlquimia(UserIndex)
1370                    Call WriteShowAlquimiaForm(UserIndex)

1380                Case e_SubObjType.CarpentrySaw    ' Herramientas de Carpinteria - Serrucho
1390                    Call EnivarObjConstruibles(UserIndex)
1400                    Call WriteShowCarpenterForm(UserIndex)

1410                Case e_SubObjType.LumberjackAxe    ' Herramientas de Tala - Hacha
1420                    Call WriteWorkRequestTarget(UserIndex, e_Skill.Talar)

1430                Case e_SubObjType.BlacksmithHammer    ' 7     ' Herramientas de Herrero - Martillo
                        ' Msg680=Debes hacer click derecho sobre el yunque.
1440                    Call WriteLocaleMsg(UserIndex, 680, e_FontTypeNames.FONTTYPE_INFOIAO)

1450                Case e_SubObjType.MinerPicket    ' 8     ' Herramientas de Mineria - Piquete
1460                    Call WriteWorkRequestTarget(UserIndex, e_Skill.Mineria)

1470                Case e_SubObjType.TailoringSewing    ' Herramientas de Sastreria - Costurero
1480                    Call EnivarObjConstruiblesSastre(UserIndex)
1490                    Call WriteShowSastreForm(UserIndex)
1500            End Select

1510        Case e_OBJType.otPotions

1520            If .flags.Muerto = 1 Then
1530                Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
1540                Exit Sub
1550            End If

1560            If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
                    ' Msg681=¡¡Debes esperar unos momentos para tomar otra poción!!
1570                Call WriteLocaleMsg(UserIndex, "681", e_FontTypeNames.FONTTYPE_INFO)
1580                Exit Sub
1590            End If

1600            .flags.TomoPocion = True
1610            .flags.TipoPocion = obj.TipoPocion

                Dim CabezaFinal As Integer
                Dim CabezaActual As Integer
                ' Esta en Zona de Pelea?
                Dim triggerStatus As e_Trigger6

1620            triggerStatus = TriggerZonaPelea(UserIndex, UserIndex)

1630            Select Case .flags.TipoPocion

                    Case e_TipoPocion.Agility    'Modif la agilidad
1640                    .flags.DuracionEfecto = obj.DuracionEfecto

                        'Usa el item
1650                    .Stats.UserAtributos(e_Atributos.Agilidad) = MinimoInt(.Stats.UserAtributos(e_Atributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador), .Stats.UserAtributosBackUP(e_Atributos.Agilidad) * 2)

1660                    Call WriteFYA(UserIndex)

                        ' Consumir pocion solo si el usuario no esta en zona de uso libre
1670                    If Not IsPotionFreeZone(UserIndex, triggerStatus) Then
                            ' Quitamos el ítem del inventario
1680                        Call QuitarUserInvItem(UserIndex, Slot, 1)
1690                    End If

1700                    If obj.Snd1 <> 0 Then
1710                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
1720                    Else
1730                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
1740                    End If

1750                Case e_TipoPocion.Strength    '2     'Modif la fuerza
1760                    .flags.DuracionEfecto = obj.DuracionEfecto

                        'Usa el item
1770                    .Stats.UserAtributos(e_Atributos.Fuerza) = MinimoInt(.Stats.UserAtributos(e_Atributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador), .Stats.UserAtributosBackUP(e_Atributos.Fuerza) * 2)

                        ' Consumir pocion solo si el usuario no esta en zona de uso libre
1780                    If Not IsPotionFreeZone(UserIndex, triggerStatus) Then
                            ' Quitamos el ítem del inventario
1790                        Call QuitarUserInvItem(UserIndex, Slot, 1)
1800                    End If

1810                    If obj.Snd1 <> 0 Then
1820                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
1830                    Else
1840                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
1850                    End If

1860                    Call WriteFYA(UserIndex)

1870                Case e_TipoPocion.Hp    'Poción roja, restaura HP
                        ' Usa el ítem
                        Dim HealingAmount As Long
                        Dim Source As Integer

                        ' Calcula la cantidad de curación
1880                    HealingAmount = RandomNumber(obj.MinModificador, obj.MaxModificador) * UserMod.GetSelfHealingBonus(UserList(UserIndex))

                        ' Modifica la salud del jugador
1890                    Call UserMod.ModifyHealth(UserIndex, HealingAmount)

                        ' Consumir pocion solo si el usuario no esta en zona de uso libre
1900                    If Not IsPotionFreeZone(UserIndex, triggerStatus) Then
                            ' Quitamos el ítem del inventario
1910                        Call QuitarUserInvItem(UserIndex, Slot, 1)
1920                    End If

                        ' Reproduce sonido al usar la poción
1930                    If obj.Snd1 <> 0 Then
1940                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
1950                    Else
1960                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
1970                    End If

1980                Case e_TipoPocion.Mp    'Poción azul, restaura MANA
                        Dim porcentajeRec As Byte
1990                    porcentajeRec = obj.Porcentaje

                        ' Usa el ítem: restaura el MANA
2000                    .Stats.MinMAN = IIf(.Stats.MinMAN > 20000, 20000, .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, porcentajeRec))
2010                    If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN

                        ' Consumir pocion solo si el usuario no esta en zona de uso libre
2020                    If Not IsPotionFreeZone(UserIndex, triggerStatus) Then
                            ' Quitamos el ítem del inventario
2030                        Call QuitarUserInvItem(UserIndex, Slot, 1)
2040                    End If

                        ' Reproduce sonido al usar la poción
2050                    If obj.Snd1 <> 0 Then
2060                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
2070                    Else
2080                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
2090                    End If

2100                Case e_TipoPocion.Poison    ' Pocion violeta

2110                    If .flags.Envenenado > 0 Then
2120                        .flags.Envenenado = 0
                            ' Msg682=Te has curado del envenenamiento.
2130                        Call WriteLocaleMsg(UserIndex, "682", e_FontTypeNames.FONTTYPE_INFO)
                            'Quitamos del inv el item
2140                        Call QuitarUserInvItem(UserIndex, Slot, 1)

2150                        If obj.Snd1 <> 0 Then
2160                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
2170                        Else
2180                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
2190                        End If

2200                    Else
                            ' Msg683=¡No te encuentras envenenado!
2210                        Call WriteLocaleMsg(UserIndex, "683", e_FontTypeNames.FONTTYPE_INFO)
2220                    End If

2230                Case e_TipoPocion.RemoveParalisis    ' Remueve Parálisis

2240                    If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
2250                        If .flags.Paralizado = 1 Then
2260                            .flags.Paralizado = 0
2270                            Call WriteParalizeOK(UserIndex)
2280                        End If

2290                        If .flags.Inmovilizado = 1 Then
2300                            .Counters.Inmovilizado = 0
2310                            .flags.Inmovilizado = 0
2320                            Call WriteInmovilizaOK(UserIndex)
2330                        End If

2340                        Call QuitarUserInvItem(UserIndex, Slot, 1)

2350                        If obj.Snd1 <> 0 Then
2360                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
2370                        Else
2380                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(255, .pos.x, .pos.y))
2390                        End If

                            ' Msg684=Te has removido la paralizis.
2400                        Call WriteLocaleMsg(UserIndex, "684", e_FontTypeNames.FONTTYPE_INFOIAO)
2410                    Else
                            ' Msg685=No estas paralizado.
2420                        Call WriteLocaleMsg(UserIndex, "685", e_FontTypeNames.FONTTYPE_INFOIAO)

2430                    End If

2440                Case e_TipoPocion.Stamina    ' Pocion Naranja
2450                    .Stats.MinSta = .Stats.MinSta + RandomNumber(obj.MinModificador, obj.MaxModificador)

2460                    If .Stats.MinSta > .Stats.MaxSta Then .Stats.MinSta = .Stats.MaxSta

                        'Quitamos del inv el item
2470                    Call QuitarUserInvItem(UserIndex, Slot, 1)

2480                    If obj.Snd1 <> 0 Then
2490                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
2500                    Else
2510                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
2520                    End If

2530                Case e_TipoPocion.ChangeHead    ' Pocion cambio cara

2540                    Select Case .genero
                            Case e_Genero.Hombre
2550                            Select Case .raza
                                    Case e_Raza.Humano
2560                                    CabezaFinal = RandomNumber(1, 40)

2570                                Case e_Raza.Elfo
2580                                    CabezaFinal = RandomNumber(101, 132)

2590                                Case e_Raza.Drow
2600                                    CabezaFinal = RandomNumber(201, 229)

2610                                Case e_Raza.Enano
2620                                    CabezaFinal = RandomNumber(301, 329)

2630                                Case e_Raza.Gnomo
2640                                    CabezaFinal = RandomNumber(401, 429)

2650                                Case e_Raza.Orco
2660                                    CabezaFinal = RandomNumber(501, 529)
2670                            End Select

2680                        Case e_Genero.Mujer
2690                            Select Case .raza
                                    Case e_Raza.Humano
2700                                    CabezaFinal = RandomNumber(50, 80)

2710                                Case e_Raza.Elfo
2720                                    CabezaFinal = RandomNumber(150, 179)

2730                                Case e_Raza.Drow
2740                                    CabezaFinal = RandomNumber(250, 279)

2750                                Case e_Raza.Gnomo
2760                                    CabezaFinal = RandomNumber(350, 379)

2770                                Case e_Raza.Enano
2780                                    CabezaFinal = RandomNumber(450, 479)

2790                                Case e_Raza.Orco
2800                                    CabezaFinal = RandomNumber(550, 579)
2810                            End Select
2820                    End Select

2830                    .Char.head = CabezaFinal
2840                    .OrigChar.head = CabezaFinal
2850                    .OrigChar.originalhead = CabezaFinal    'cabeza final
2860                    Call ChangeUserChar(UserIndex, .Char.body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                        'Quitamos del inv el item

2870                    .Counters.timeFx = 3
2880                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 102, 0, .pos.x, .pos.y))

2890                    If CabezaActual <> CabezaFinal Then
2900                        Call QuitarUserInvItem(UserIndex, Slot, 1)
2910                    Else
                            ' Msg686=¡Rayos! Te tocó la misma cabeza, item no consumido. Tienes otra oportunidad.
2920                        Call WriteLocaleMsg(UserIndex, "686", e_FontTypeNames.FONTTYPE_INFOIAO)
2930                    End If

2940                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))

2950                Case e_TipoPocion.ChangeSex     ' Pocion sexo

2960                    Select Case .genero
                            Case e_Genero.Hombre
2970                            .genero = e_Genero.Mujer

2980                        Case e_Genero.Mujer
2990                            .genero = e_Genero.Hombre
3000                    End Select

3010                    Select Case .genero

                            Case e_Genero.Hombre
3020                            Select Case .raza
                                    Case e_Raza.Humano
3030                                    CabezaFinal = RandomNumber(1, 40)

3040                                Case e_Raza.Elfo
3050                                    CabezaFinal = RandomNumber(101, 132)

3060                                Case e_Raza.Drow
3070                                    CabezaFinal = RandomNumber(201, 229)

3080                                Case e_Raza.Enano
3090                                    CabezaFinal = RandomNumber(301, 329)

3100                                Case e_Raza.Gnomo
3110                                    CabezaFinal = RandomNumber(401, 429)

3120                                Case e_Raza.Orco
3130                                    CabezaFinal = RandomNumber(501, 529)
3140                            End Select

3150                        Case e_Genero.Mujer

3160                            Select Case .raza

                                    Case e_Raza.Humano
3170                                    CabezaFinal = RandomNumber(50, 80)

3180                                Case e_Raza.Elfo
3190                                    CabezaFinal = RandomNumber(150, 179)

3200                                Case e_Raza.Drow
3210                                    CabezaFinal = RandomNumber(250, 279)

3220                                Case e_Raza.Gnomo
3230                                    CabezaFinal = RandomNumber(350, 379)

3240                                Case e_Raza.Enano
3250                                    CabezaFinal = RandomNumber(450, 479)

3260                                Case e_Raza.Orco
3270                                    CabezaFinal = RandomNumber(550, 579)
3280                            End Select
3290                    End Select

3300                    .Char.head = CabezaFinal
3310                    .OrigChar.head = CabezaFinal
3320                    Call ChangeUserChar(UserIndex, .Char.body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
                        'Quitamos del inv el item
3330                    .Counters.timeFx = 3
3340                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 102, 0, .pos.x, .pos.y))
3350                    Call QuitarUserInvItem(UserIndex, Slot, 1)

3360                    If obj.Snd1 <> 0 Then
3370                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
3380                    Else
3390                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
3400                    End If

3410                Case e_TipoPocion.Invisibility    ' Invisibilidad

3420                    If .flags.invisible = 0 And .Counters.DisabledInvisibility = 0 Then
3430                        If IsSet(.flags.StatusMask, eTaunting) Then
                                ' Msg687=No tiene efecto.
3440                            Call WriteLocaleMsg(UserIndex, "687", e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
3450                            Exit Sub
3460                        End If

3470                        .flags.invisible = 1
3480                        .Counters.Invisibilidad = obj.DuracionEfecto
3490                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True, .pos.x, .pos.y))
3500                        Call WriteContadores(UserIndex)
3510                        Call QuitarUserInvItem(UserIndex, Slot, 1)

3520                        If obj.Snd1 <> 0 Then
3530                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))

3540                        Else
3550                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave("123", .pos.x, .pos.y))
3560                        End If

                            ' Msg688=Te has escondido entre las sombras...
3570                        Call WriteLocaleMsg(UserIndex, "688", e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)

3580                    Else
                            ' Msg689=Ya estás invisible.
3590                        Call WriteLocaleMsg(UserIndex, "689", e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
3600                        Exit Sub
3610                    End If

                        ' Poción que limpia todo
3620                Case e_TipoPocion.CleanEverything

3630                    Call QuitarUserInvItem(UserIndex, Slot, 1)
3640                    .flags.Envenenado = 0
3650                    .flags.Incinerado = 0

3660                    If .flags.Inmovilizado = 1 Then
3670                        .Counters.Inmovilizado = 0
3680                        .flags.Inmovilizado = 0
3690                        Call WriteInmovilizaOK(UserIndex)
3700                    End If

3710                    If .flags.Paralizado = 1 Then
3720                        .flags.Paralizado = 0
3730                        Call WriteParalizeOK(UserIndex)
3740                    End If

3750                    If .flags.Ceguera = 1 Then
3760                        .flags.Ceguera = 0
3770                        Call WriteBlindNoMore(UserIndex)
3780                    End If

3790                    If .flags.Maldicion = 1 Then
3800                        .flags.Maldicion = 0
3810                        .Counters.Maldicion = 0
3820                    End If

3830                    .Stats.MinSta = .Stats.MaxSta
3840                    .Stats.MinAGU = .Stats.MaxAGU
3850                    .Stats.MinMAN = .Stats.MaxMAN
3860                    .Stats.MinHp = .Stats.MaxHp
3870                    .Stats.MinHam = .Stats.MaxHam

3880                    Call WriteUpdateHungerAndThirst(UserIndex)
                        ' Msg690=Donador> Te sentís sano y lleno.
3890                    Call WriteLocaleMsg(UserIndex, "690", e_FontTypeNames.FONTTYPE_WARNING)

3900                    If obj.Snd1 <> 0 Then
3910                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))

3920                    Else
3930                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
3940                    End If

                        ' Poción runa
3950                Case e_TipoPocion.Rune

3960                    If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
                            ' Msg691=No podés usar la runa estando en la cárcel.
3970                        Call WriteLocaleMsg(UserIndex, "691", e_FontTypeNames.FONTTYPE_INFO)
3980                        Exit Sub
3990                    End If

                        Dim Map As Integer
                        Dim x   As Byte
                        Dim y   As Byte
                        Dim DeDonde As t_WorldPos

4000                    Call QuitarUserInvItem(UserIndex, Slot, 1)

4010                    Select Case .Hogar

                            Case e_Ciudad.cUllathorpe
4020                            DeDonde = Ullathorpe

4030                        Case e_Ciudad.cNix
4040                            DeDonde = Nix

4050                        Case e_Ciudad.cBanderbill
4060                            DeDonde = Banderbill

4070                        Case e_Ciudad.cLindos
4080                            DeDonde = Lindos

4090                        Case e_Ciudad.cArghal
4100                            DeDonde = Arghal

4110                        Case e_Ciudad.cArkhein
4120                            DeDonde = Arkhein

4130                        Case e_Ciudad.cForgat
4140                            DeDonde = Forgat

4150                        Case e_Ciudad.cEldoria
4160                            DeDonde = Eldoria

4170                        Case e_Ciudad.cPenthar
4180                            DeDonde = Penthar

4190                        Case Else
4200                            DeDonde = Ullathorpe

4210                    End Select

4220                    Map = DeDonde.Map
4230                    x = DeDonde.x
4240                    y = DeDonde.y

4250                    Call FindLegalPos(UserIndex, Map, x, y)
4260                    Call WarpUserChar(UserIndex, Map, x, y, True)
                        'Msg884= Ya estas a salvo...
4270                    Call WriteLocaleMsg(UserIndex, "884", e_FontTypeNames.FONTTYPE_WARNING)

4280                    If obj.Snd1 <> 0 Then
4290                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
4300                    Else
4310                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
4320                    End If

4330                Case e_TipoPocion.Divorce    ' Divorcio

4340                    If .flags.Casado = 1 Then
                            Dim tUser As t_UserReference

                            '.flags.Pareja
4350                        tUser = NameIndex(GetUserSpouse(.flags.SpouseId))

4360                        If Not IsValidUserRef(tUser) Then
                                'Msg885= Tu pareja deberás estar conectada para divorciarse.
4370                            Call WriteLocaleMsg(UserIndex, "885", e_FontTypeNames.FONTTYPE_INFOIAO)
4380                        Else
4390                            Call QuitarUserInvItem(UserIndex, Slot, 1)
4400                            UserList(tUser.ArrayIndex).flags.Casado = 0
4410                            UserList(tUser.ArrayIndex).flags.SpouseId = 0
4420                            .flags.Casado = 0
4430                            .flags.SpouseId = 0
                                'Msg886= Te has divorciado.
4440                            Call WriteLocaleMsg(UserIndex, "886", e_FontTypeNames.FONTTYPE_INFOIAO)

4450                            Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1983, .name, e_FontTypeNames.FONTTYPE_INFOIAO))    ' Msg1983=¬1 se ha divorciado de ti.

4460                            If obj.Snd1 <> 0 Then
4470                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
4480                            Else
4490                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
4500                            End If
4510                        End If
4530                    Else
                            'Msg887= No estas casado.
4540                        Call WriteLocaleMsg(UserIndex, "887", e_FontTypeNames.FONTTYPE_INFOIAO)

4550                    End If

4560                Case e_TipoPocion.Legendary    'Cara legendaria

4570                    Select Case .genero

                            Case e_Genero.Hombre

4580                            Select Case .raza

                                    Case e_Raza.Humano
4590                                    CabezaFinal = RandomNumber(684, 686)

4600                                Case e_Raza.Elfo
4610                                    CabezaFinal = RandomNumber(690, 692)

4620                                Case e_Raza.Drow
4630                                    CabezaFinal = RandomNumber(696, 698)

4640                                Case e_Raza.Enano
4650                                    CabezaFinal = RandomNumber(702, 704)

4660                                Case e_Raza.Gnomo
4670                                    CabezaFinal = RandomNumber(708, 710)

4680                                Case e_Raza.Orco
4690                                    CabezaFinal = RandomNumber(714, 716)

4700                            End Select

4710                        Case e_Genero.Mujer

4720                            Select Case .raza

                                    Case e_Raza.Humano
4730                                    CabezaFinal = RandomNumber(687, 689)

4740                                Case e_Raza.Elfo
4750                                    CabezaFinal = RandomNumber(693, 695)

4760                                Case e_Raza.Drow
4770                                    CabezaFinal = RandomNumber(699, 701)

4780                                Case e_Raza.Gnomo
4790                                    CabezaFinal = RandomNumber(705, 707)

4800                                Case e_Raza.Enano
4810                                    CabezaFinal = RandomNumber(711, 713)

4820                                Case e_Raza.Orco
4830                                    CabezaFinal = RandomNumber(717, 719)

4840                            End Select

4850                    End Select

4860                    CabezaActual = .OrigChar.head

4870                    .Char.head = CabezaFinal
4880                    .OrigChar.head = CabezaFinal
4890                    Call ChangeUserChar(UserIndex, .Char.body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)

                        'Quitamos del inv el item
4900                    If CabezaActual <> CabezaFinal Then
4910                        .Counters.timeFx = 3
4920                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 102, 0, .pos.x, .pos.y))
4930                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
4940                        Call QuitarUserInvItem(UserIndex, Slot, 1)
4950                    Else
                            'Msg888= ¡Rayos! No pude asignarte una cabeza nueva, item no consumido. ¡Proba de nuevo!
4960                        Call WriteLocaleMsg(UserIndex, "888", e_FontTypeNames.FONTTYPE_INFOIAO)

4970                    End If

4980                Case e_TipoPocion.Particle  '18 tan solo crea una particula por determinado tiempo

                        Dim Particula As Integer

                        Dim Tiempo As Long

                        Dim ParticulaPermanente As Byte

                        Dim sobrechar As Byte

4990                    If obj.CreaParticula <> "" Then
5000                        Particula = val(ReadField(1, obj.CreaParticula, Asc(":")))
5010                        Tiempo = val(ReadField(2, obj.CreaParticula, Asc(":")))
5020                        ParticulaPermanente = val(ReadField(3, obj.CreaParticula, Asc(":")))
5030                        sobrechar = val(ReadField(4, obj.CreaParticula, Asc(":")))

5040                        If ParticulaPermanente = 1 Then
5050                            .Char.ParticulaFx = Particula
5060                            .Char.loops = Tiempo

5070                        End If

5080                        If sobrechar = 1 Then
5090                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFXToFloor(.pos.x, .pos.y, Particula, Tiempo))
5100                        Else
5110                            .Counters.timeFx = 3
5120                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, Particula, Tiempo, False, , .pos.x, .pos.y))
5130                        End If

5140                    End If

5150                    If obj.CreaFX <> 0 Then
5160                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageFxPiso(obj.CreaFX, .pos.x, .pos.y))
5170                    End If

5180                    If obj.Snd1 <> 0 Then
5190                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))

5200                    End If

5210                    Call QuitarUserInvItem(UserIndex, Slot, 1)

5220                Case e_TipoPocion.ResetSkill '19    ' Reseteo de skill

                        Dim s   As Byte

5230                    If .Stats.UserSkills(e_Skill.liderazgo) >= 80 Then
                            'Msg889= Has fundado un clan, no podes resetar tus skills.
5240                        Call WriteLocaleMsg(UserIndex, "889", e_FontTypeNames.FONTTYPE_INFOIAO)
5250                        Exit Sub

5260                    End If

5270                    For s = 1 To NUMSKILLS
5280                        .Stats.UserSkills(s) = 0
5290                    Next s

                        Dim SkillLibres As Integer

5300                    SkillLibres = 5
5310                    SkillLibres = SkillLibres + (5 * .Stats.ELV)

5320                    .Stats.SkillPts = SkillLibres
5330                    Call WriteLevelUp(UserIndex, .Stats.SkillPts)
                        'Msg890= Tus skills han sido reseteados.
5340                    Call WriteLocaleMsg(UserIndex, "890", e_FontTypeNames.FONTTYPE_INFOIAO)
5350                    Call QuitarUserInvItem(UserIndex, Slot, 1)


                        ' Poción negra (suicidio)
5360                Case e_TipoPocion.Suicide '21
                        'Quitamos del inv el item
5370                    Call QuitarUserInvItem(UserIndex, Slot, 1)

5380                    If obj.Snd1 <> 0 Then
5390                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
5400                    Else
5410                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))
5420                    End If

                        'Msg893= Te has suicidado.
5430                    Call WriteLocaleMsg(UserIndex, "893", e_FontTypeNames.FONTTYPE_EJECUCION)
5440                    Call CustomScenarios.UserDie(UserIndex)
5450                    Call UserMod.UserDie(UserIndex)
                        'Poción de reset (resetea el personaje)
5460                Case e_TipoPocion.ResetCharacter '22
5470                    If GetTickCount - .Counters.LastResetTick > 3000 Then
5480                        Call writeAnswerReset(UserIndex)
5490                        .Counters.LastResetTick = GetTickCount
5500                    Else
                            'Msg894= Debes esperar unos momentos para tomar esta poción.
5510                        Call WriteLocaleMsg(UserIndex, "894", e_FontTypeNames.FONTTYPE_INFO)
5520                    End If

5530                Case e_TipoPocion.PoisonWeapon ' 23
5540                    If obj.ApplyEffectId > 0 Then
5550                        Call AddOrResetEffect(UserIndex, obj.ApplyEffectId)
5560                    End If
5570                    Call UpdateCd(UserIndex, ObjData(ObjIndex).cdType)
                        'Quitamos del inv el item
5580                    Call QuitarUserInvItem(UserIndex, Slot, 1)
5590                    Call UpdateUserInv(False, UserIndex, Slot)
5600                    Exit Sub
5610            End Select

5620            If obj.ApplyEffectId > 0 Then
5630                Call AddOrResetEffect(UserIndex, obj.ApplyEffectId)
5640            End If

5650            Call WriteUpdateUserStats(UserIndex)
5660            Call UpdateUserInv(False, UserIndex, Slot)

5670        Case e_OBJType.otDrinks

5680            If .flags.Muerto = 1 Then
5690                Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
5700                Exit Sub
5710            End If

5720            .Stats.MinAGU = .Stats.MinAGU + obj.MinSed

5730            If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
5740            Call WriteUpdateHungerAndThirst(UserIndex)

                'Quitamos del inv el item
5750            Call QuitarUserInvItem(UserIndex, Slot, 1)
5760            If obj.ApplyEffectId > 0 Then
5770                Call AddOrResetEffect(UserIndex, obj.ApplyEffectId)
5780            End If

5790            If obj.Snd1 <> 0 Then
5800                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))

5810            Else
5820                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .pos.x, .pos.y))

5830            End If

5840            Call UpdateUserInv(False, UserIndex, Slot)

5850        Case e_OBJType.otChest

5860            If .flags.Muerto = 1 Then
5870                Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
5880                Exit Sub

5890            End If

                'Quitamos del inv el item
5900            Call QuitarUserInvItem(UserIndex, Slot, 1)
5910            Call UpdateUserInv(False, UserIndex, Slot)


5920            Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1984, obj.name, e_FontTypeNames.FONTTYPE_New_DONADOR))    ' Msg1984=Has abierto un ¬1 y obtuviste...


5930            If obj.Snd1 <> 0 Then
5940                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
5950            End If

5960            If obj.CreaFX <> 0 Then
5970                .Counters.timeFx = 3
5980                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, obj.CreaFX, 0, .pos.x, .pos.y))
5990            End If

                Dim i           As Byte

6000            Select Case obj.Subtipo

                    Case 1

6010                    For i = 1 To obj.CantItem

6020                        If Not MeterItemEnInventario(UserIndex, obj.Item(i)) Then

6030                            If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Dios Or e_PlayerType.Admin)) Then
6040                                Call TirarItemAlPiso(.pos, obj.Item(i))
6050                            End If

6060                        End If

6070                        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).name & " (" & obj.Item(i).amount & ")", e_FontTypeNames.FONTTYPE_INFOBOLD))

6080                    Next i

6090                Case 2

6100                    For i = 1 To obj.CantEntrega

                            Dim indexobj As Byte
6110                        indexobj = RandomNumber(1, obj.CantItem)

                            Dim Index As t_Obj

6120                        Index.ObjIndex = obj.Item(indexobj).ObjIndex
6130                        Index.amount = obj.Item(indexobj).amount

6140                        If Not MeterItemEnInventario(UserIndex, Index) Then

6150                            If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Dios Or e_PlayerType.Admin)) Then
6160                                Call TirarItemAlPiso(.pos, Index)
6170                            End If

6180                        End If

6190                        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageConsoleMsg(ObjData(Index.ObjIndex).name & " (" & Index.amount & ")", e_FontTypeNames.FONTTYPE_INFOBOLD))
6200                    Next i

6210                Case 3

6220                    For i = 1 To obj.CantItem

6230                        If RandomNumber(1, obj.Item(i).data) = 1 Then

6240                            If Not MeterItemEnInventario(UserIndex, obj.Item(i)) Then

6250                                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Dios Or e_PlayerType.Admin)) Then
6260                                    Call TirarItemAlPiso(.pos, obj.Item(i))
6270                                End If

6280                            End If

6290                            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).name & " (" & obj.Item(i).amount & ")", e_FontTypeNames.FONTTYPE_INFOBOLD))

6300                        End If

6310                    Next i

6320            End Select

6330        Case e_OBJType.otKeys
6340            If .flags.Muerto = 1 Then
                    'Msg895= ¡¡Estas muerto!! Solo podes usar items cuando estas vivo.
6350                Call WriteLocaleMsg(UserIndex, "895", e_FontTypeNames.FONTTYPE_INFO)
6360                Exit Sub
6370            End If

6380            If .flags.TargetObj = 0 Then Exit Sub
6390            TargObj = ObjData(.flags.TargetObj)
                '¿El objeto clickeado es una puerta?
6400            If TargObj.OBJType = e_OBJType.otDoors Then
6410                If TargObj.clave < 1000 Then
                        'Msg896= Las llaves en el inventario están desactivadas. Sólo se permiten en el llavero.
6420                    Call WriteLocaleMsg(UserIndex, "896", e_FontTypeNames.FONTTYPE_INFO)
6430                    Exit Sub
6440                End If

                    '¿Esta cerrada?
6450                If TargObj.Cerrada = 1 Then
                        '¿Cerrada con llave?
6460                    If TargObj.Llave > 0 Then
                            Dim ClaveLlave As Integer

6470                        If TargObj.clave = obj.clave Then
6480                            MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                                        = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
6490                            .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                                'Msg897= Has abierto la puerta.
6500                            Call WriteLocaleMsg(UserIndex, "897", e_FontTypeNames.FONTTYPE_INFO)
6510                            ClaveLlave = obj.clave
6520                            Call EliminarLlaves(ClaveLlave, UserIndex)
6530                            Exit Sub
6540                        Else
                                'Msg898= La llave no sirve.
6550                            Call WriteLocaleMsg(UserIndex, "898", e_FontTypeNames.FONTTYPE_INFO)
6560                            Exit Sub
6570                        End If
6580                    Else
6590                        If TargObj.clave = obj.clave Then
6600                            MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                                        = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                                'Msg899= Has cerrado con llave la puerta.
6610                            Call WriteLocaleMsg(UserIndex, "899", e_FontTypeNames.FONTTYPE_INFO)
6620                            .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
6630                            Exit Sub
6640                        Else
                                'Msg900= La llave no sirve.
6650                            Call WriteLocaleMsg(UserIndex, "900", e_FontTypeNames.FONTTYPE_INFO)
6660                            Exit Sub
6670                        End If
6680                    End If
6690                Else
                        'Msg901= No esta cerrada.
6700                    Call WriteLocaleMsg(UserIndex, "901", e_FontTypeNames.FONTTYPE_INFO)
6710                    Exit Sub
6720                End If
6730            End If

6740        Case e_OBJType.otEmptyBottle

6750            If .flags.Muerto = 1 Then
6760                Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
6770                Exit Sub
6780            End If

6790            If Not InMapBounds(.flags.TargetMap, .flags.TargetX, .flags.TargetY) Then
6800                Exit Sub
6810            End If

6820            If (MapData(.pos.Map, .flags.TargetX, .flags.TargetY).Blocked And FLAG_AGUA) = 0 Then
                    'Msg902= No hay agua allí.
6830                Call WriteLocaleMsg(UserIndex, "902", e_FontTypeNames.FONTTYPE_INFO)
6840                Exit Sub
6850            End If

6860            If Distance(.pos.x, .pos.y, .flags.TargetX, .flags.TargetY) > 2 Then
                    'Msg903= Debes acercarte más al agua.
6870                Call WriteLocaleMsg(UserIndex, "903", e_FontTypeNames.FONTTYPE_INFO)
6880                Exit Sub
6890            End If

6900            MiObj.amount = 1
6910            MiObj.ObjIndex = ObjData(.invent.Object(Slot).ObjIndex).IndexAbierta

6920            Call QuitarUserInvItem(UserIndex, Slot, 1)

6930            If Not MeterItemEnInventario(UserIndex, MiObj) Then
6940                Call TirarItemAlPiso(.pos, MiObj)
6950            End If

6960            Call UpdateUserInv(False, UserIndex, Slot)

6970        Case e_OBJType.otFullBottle

6980            If .flags.Muerto = 1 Then
6990                Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
7000                Exit Sub

7010            End If

7020            .Stats.MinAGU = .Stats.MinAGU + obj.MinSed

7030            If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
7040            Call WriteUpdateHungerAndThirst(UserIndex)
7050            MiObj.amount = 1
7060            MiObj.ObjIndex = ObjData(.invent.Object(Slot).ObjIndex).IndexCerrada
7070            Call QuitarUserInvItem(UserIndex, Slot, 1)

7080            If Not MeterItemEnInventario(UserIndex, MiObj) Then
7090                Call TirarItemAlPiso(.pos, MiObj)

7100            End If

7110            Call UpdateUserInv(False, UserIndex, Slot)

7120        Case e_OBJType.otParchment

7130            If .flags.Muerto = 1 Then
7140                Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
7150                Exit Sub

7160            End If

                'Call LogError(.Name & " intento aprender el hechizo " & ObjData(.Invent.Object(slot).ObjIndex).HechizoIndex)

7170            If ClasePuedeUsarItem(UserIndex, .invent.Object(Slot).ObjIndex, Slot) And RazaPuedeUsarItem(UserIndex, .invent.Object(Slot).ObjIndex, Slot) Then

                    'If .Stats.MaxMAN > 0 Then
7180                If .Stats.MinHam > 0 And .Stats.MinAGU > 0 Then
7190                    Call AgregarHechizo(UserIndex, Slot)
7200                    Call UpdateUserInv(False, UserIndex, Slot)
                        ' Call LogError(.Name & " lo aprendio.")
7210                Else
                        'Msg904= Estas demasiado hambriento y sediento.
7220                    Call WriteLocaleMsg(UserIndex, "904", e_FontTypeNames.FONTTYPE_INFO)

7230                End If

                    ' Else
                    '    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", e_FontTypeNames.FONTTYPE_WARNING)
                    'End If
7240            Else

                    'Msg906= Por mas que lo intentas, no podés comprender el manuescrito.
7250                Call WriteLocaleMsg(UserIndex, "906", e_FontTypeNames.FONTTYPE_INFO)

7260            End If

7270        Case e_OBJType.otMinerals

7280            If .flags.Muerto = 1 Then
7290                Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
7300                Exit Sub

7310            End If

7320            Call WriteWorkRequestTarget(UserIndex, FundirMetal)

7330        Case e_OBJType.otMusicalInstruments

7340            If .flags.Muerto = 1 Then
7350                Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
7360                Exit Sub

7370            End If

7380            If obj.Real Then    '¿Es el Cuerno Real?
7390                If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
7400                    If MapInfo(.pos.Map).Seguro = 1 Then
                            'Msg907= No hay Peligro aquí. Es Zona Segura
7410                        Call WriteLocaleMsg(UserIndex, "907", e_FontTypeNames.FONTTYPE_INFO)
7420                        Exit Sub

7430                    End If

7440                    Call SendData(SendTarget.toMap, .pos.Map, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
7450                    Exit Sub
7460                Else
                        'Msg908= Solo Miembros de la Armada Real pueden usar este cuerno.
7470                    Call WriteLocaleMsg(UserIndex, "908", e_FontTypeNames.FONTTYPE_INFO)
7480                    Exit Sub

7490                End If

7500            ElseIf obj.Caos Then    '¿Es el Cuerno Legión?

7510                If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
7520                    If MapInfo(.pos.Map).Seguro = 1 Then
                            'Msg909= No hay Peligro aquí. Es Zona Segura
7530                        Call WriteLocaleMsg(UserIndex, "909", e_FontTypeNames.FONTTYPE_INFO)
7540                        Exit Sub

7550                    End If

7560                    Call SendData(SendTarget.toMap, .pos.Map, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))
7570                    Exit Sub
7580                Else
                        'Msg910= Solo Miembros de la Legión Oscura pueden usar este cuerno.
7590                    Call WriteLocaleMsg(UserIndex, "910", e_FontTypeNames.FONTTYPE_INFO)
7600                    Exit Sub

7610                End If

7620            End If

                'Si llega aca es porque es o Laud o Tambor o Flauta
7630            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .pos.x, .pos.y))

7640        Case e_OBJType.otShips

                ' Piratas y trabajadores navegan al nivel 23
7650            If .invent.Object(Slot).ObjIndex <> iObjTrajeAltoNw And .invent.Object(Slot).ObjIndex <> iObjTrajeBajoNw And .invent.Object(Slot).ObjIndex <> iObjTraje Then
7660                If .clase = e_Class.Trabajador Or .clase = e_Class.Pirat Then
7670                    If .Stats.ELV < 23 Then
                            'Msg911= Para recorrer los mares debes ser nivel 23 o superior.
7680                        Call WriteLocaleMsg(UserIndex, "911", e_FontTypeNames.FONTTYPE_INFO)
7690                        Exit Sub
7700                    End If
                        ' Nivel mínimo 25 para navegar, si no sos pirata ni trabajador
7710                ElseIf .Stats.ELV < 25 Then
                        'Msg912= Para recorrer los mares debes ser nivel 25 o superior.
7720                    Call WriteLocaleMsg(UserIndex, "912", e_FontTypeNames.FONTTYPE_INFO)
7730                    Exit Sub
7740                End If
7750            ElseIf .invent.Object(Slot).ObjIndex = iObjTrajeAltoNw Or .invent.Object(Slot).ObjIndex = iObjTrajeBajoNw Then
7760                If (.flags.Navegando = 0 Or (.invent.EquippedShipObjIndex <> iObjTrajeAltoNw And .invent.EquippedShipObjIndex <> iObjTrajeBajoNw)) And MapData(.pos.Map, .pos.x + 1, .pos.y).trigger <> e_Trigger.DETALLEAGUA And MapData(.pos.Map, .pos.x - 1, .pos.y).trigger <> e_Trigger.DETALLEAGUA And MapData(.pos.Map, .pos.x, .pos.y + 1).trigger <> e_Trigger.DETALLEAGUA And MapData(.pos.Map, .pos.x, .pos.y - 1).trigger <> e_Trigger.DETALLEAGUA Then
                        'Msg913= Este traje es para aguas contaminadas.
7770                    Call WriteLocaleMsg(UserIndex, "913", e_FontTypeNames.FONTTYPE_INFO)
7780                    Exit Sub
7790                End If
7800            ElseIf .invent.Object(Slot).ObjIndex = iObjTraje Then
7810                If (.flags.Navegando = 0 Or .invent.EquippedShipObjIndex <> iObjTraje) And MapData(.pos.Map, .pos.x + 1, .pos.y).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.x - 1, .pos.y).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.x, .pos.y + 1).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.x, .pos.y - 1).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.x + 1, .pos.y).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.x - 1, .pos.y).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.x, .pos.y + 1).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.x, .pos.y - 1).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.x + 1, .pos.y).trigger <> e_Trigger.NADOBAJOTECHO And MapData(.pos.Map, .pos.x - 1, .pos.y).trigger <> e_Trigger.NADOBAJOTECHO And MapData(.pos.Map, .pos.x, .pos.y + 1).trigger <> e_Trigger.NADOBAJOTECHO And MapData(.pos.Map, .pos.x, .pos.y - 1).trigger <> e_Trigger.NADOBAJOTECHO Then
                        'Msg914= Este traje es para zonas poco profundas.
7820                    Call WriteLocaleMsg(UserIndex, "914", e_FontTypeNames.FONTTYPE_INFO)
7830                    Exit Sub
7840                End If
7850            End If


7860            If .flags.Navegando = 0 Then
7870                If LegalWalk(.pos.Map, .pos.x - 1, .pos.y, e_Heading.WEST, True, False) Or LegalWalk(.pos.Map, .pos.x, .pos.y - 1, e_Heading.NORTH, True, False) Or LegalWalk(.pos.Map, .pos.x + 1, .pos.y, e_Heading.EAST, True, False) Or LegalWalk(.pos.Map, .pos.x, .pos.y + 1, e_Heading.SOUTH, True, False) Then
7880                    Call DoNavega(UserIndex, obj, Slot)
7890                Else
                        'Msg915= ¡Debes aproximarte al agua para usar el barco o traje de baño!
7900                    Call WriteLocaleMsg(UserIndex, "915", e_FontTypeNames.FONTTYPE_INFO)

7910                End If

7920            Else
7930                If .invent.EquippedShipObjIndex <> .invent.Object(Slot).ObjIndex Then
7940                    Call DoNavega(UserIndex, obj, Slot)
7950                Else
7960                    If LegalWalk(.pos.Map, .pos.x - 1, .pos.y, e_Heading.WEST, False, True) Or LegalWalk(.pos.Map, .pos.x, .pos.y - 1, e_Heading.NORTH, False, True) Or LegalWalk(.pos.Map, .pos.x + 1, .pos.y, e_Heading.EAST, False, True) Or LegalWalk(.pos.Map, .pos.x, .pos.y + 1, e_Heading.SOUTH, False, True) Then
7970                        Call DoNavega(UserIndex, obj, Slot)
7980                    Else
                            'Msg916= ¡Debes aproximarte a la costa para dejar la barca!
7990                        Call WriteLocaleMsg(UserIndex, "916", e_FontTypeNames.FONTTYPE_INFO)

8000                    End If
8010                End If
8020            End If

8030        Case e_OBJType.otSaddles
                'Verifica todo lo que requiere la montura

8040            If .flags.Muerto = 1 Then
8050                Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                    'Msg77=¡¡Estás muerto!!.
8060                Exit Sub

8070            End If

8080            If .flags.Navegando = 1 Then
                    'Msg917= Debes dejar de navegar para poder cabalgar.
8090                Call WriteLocaleMsg(UserIndex, "917", e_FontTypeNames.FONTTYPE_INFO)
8100                Exit Sub

8110            End If

8120            If MapInfo(.pos.Map).zone = "DUNGEON" Then
                    'Msg918= No podes cabalgar dentro de un dungeon.
8130                Call WriteLocaleMsg(UserIndex, "918", e_FontTypeNames.FONTTYPE_INFO)
8140                Exit Sub

8150            End If

8160            Call DoMontar(UserIndex, obj, Slot)

8170        Case e_OBJType.otDonator
8180            Select Case obj.Subtipo
                    Case 1
8190                    If .Counters.Pena <> 0 Then
                            ' Msg691=No podés usar la runa estando en la cárcel.
8200                        Call WriteLocaleMsg(UserIndex, "691", e_FontTypeNames.FONTTYPE_INFO)
8210                        Exit Sub
8220                    End If

8230                    If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
                            ' Msg691=No podés usar la runa estando en la cárcel.
8240                        Call WriteLocaleMsg(UserIndex, "691", e_FontTypeNames.FONTTYPE_INFO)
8250                        Exit Sub
8260                    End If

8270                    Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
                        'Msg919= Has viajado por el mundo.
8280                    Call WriteLocaleMsg(UserIndex, "919", e_FontTypeNames.FONTTYPE_WARNING)
8290                    Call QuitarUserInvItem(UserIndex, Slot, 1)
8300                    Call UpdateUserInv(False, UserIndex, Slot)

8310                Case 2
8320                    Exit Sub
8330                Case 3
8340                    Exit Sub
8350            End Select
8360        Case e_OBJType.otPassageTicket

8370            If .flags.Muerto = 1 Then
                    'Msg77=¡¡Estás muerto!!.
8380                Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
8390                Exit Sub
8400            End If

8410            If .flags.TargetNpcTipo <> Pirata Then
                    'Msg920= Primero debes hacer click sobre el pirata.
8420                Call WriteLocaleMsg(UserIndex, "920", e_FontTypeNames.FONTTYPE_INFO)
8430                Exit Sub
8440            End If

8450            If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).pos, .pos) > 3 Then
8460                Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
8470                Exit Sub
8480            End If

8490            If .pos.Map <> obj.DesdeMap Then
8500                Call WriteLocaleChatOverHead(UserIndex, "1354", "", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite)    ' Msg1354=El pasaje no lo compraste aquí! Largate!
8510                Exit Sub
8520            End If

8530            If Not MapaValido(obj.HastaMap) Then
8540                Call WriteLocaleChatOverHead(UserIndex, "1355", "", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite)    ' Msg1355=El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.
8550                Exit Sub
8560            End If

8570            If obj.NecesitaNave > 0 Then
8580                If .Stats.UserSkills(e_Skill.Navegacion) < 80 Then
8590                    Call WriteLocaleChatOverHead(UserIndex, "1356", "", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite)    ' Msg1356=Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.
8600                    Exit Sub
8610                End If
8620            End If

8630            Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
                'Msg921= Has viajado por varios días, te sientes exhausto!
8640            Call WriteLocaleMsg(UserIndex, "921", e_FontTypeNames.FONTTYPE_WARNING)
8650            .Stats.MinAGU = 0
8660            .Stats.MinHam = 0
8670            Call WriteUpdateHungerAndThirst(UserIndex)
8680            Call QuitarUserInvItem(UserIndex, Slot, 1)
8690            Call UpdateUserInv(False, UserIndex, Slot)

8700        Case e_OBJType.otRecallStones

8710            If .Counters.Pena <> 0 Then
                    ' Msg691=No podés usar la runa estando en la cárcel.
8720                Call WriteLocaleMsg(UserIndex, "691", e_FontTypeNames.FONTTYPE_INFO)
8730                Exit Sub

8740            End If

8750            If MapData(.pos.Map, .pos.x, .pos.y).trigger = CARCEL Then
                    ' Msg691=No podés usar la runa estando en la cárcel.
8760                Call WriteLocaleMsg(UserIndex, "691", e_FontTypeNames.FONTTYPE_INFO)
8770                Exit Sub

8780            End If

8790            If MapInfo(.pos.Map).Seguro = 0 And .flags.Muerto = 0 Then
                    ' Msg692=Solo podes usar tu runa en zonas seguras.
8800                Call WriteLocaleMsg(UserIndex, "692", e_FontTypeNames.FONTTYPE_INFO)
8810                Exit Sub

8820            End If

8830            If .Accion.AccionPendiente Then
8840                Exit Sub

8850            End If

8860            Select Case ObjData(ObjIndex).TipoRuna

                    Case e_RuneType.ReturnHome
8870                    .Counters.TimerBarra = HomeTimer
8880                Case e_RuneType.Escape
8890                    .Counters.TimerBarra = HomeTimer
8900                Case e_RuneType.MesonSafePassage
8910                    .Counters.TimerBarra = 5
8920            End Select
8930            If Not EsGM(UserIndex) Then
8940                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, e_ParticulasIndex.Runa, 400, False))
8950                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageBarFx(.Char.charindex, 350, e_AccionBarra.Runa))
8960            Else
8970                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, e_ParticulasIndex.Runa, 50, False))
8980                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageBarFx(.Char.charindex, 100, e_AccionBarra.Runa))

8990            End If

9000            .Accion.Particula = e_ParticulasIndex.Runa
9010            .Accion.AccionPendiente = True
9020            .Accion.TipoAccion = e_AccionBarra.Runa
9030            .Accion.RunaObj = ObjIndex
9040            .Accion.ObjSlot = Slot


9050        Case e_OBJType.otMap
9060            Call WriteShowFrmMapa(UserIndex)
9070        Case e_OBJType.OtQuest
9080            If obj.QuestId > 0 Then Call WriteObjQuestSend(UserIndex, obj.QuestId, Slot)
9090        Case e_OBJType.otAmulets
9100            Select Case ObjData(ObjIndex).Subtipo
                    Case e_MagicItemSubType.TargetUsable
9110                    Call WriteWorkRequestTarget(UserIndex, e_Skill.TargetableItem)
9120            End Select
9130            Select Case ObjData(ObjIndex).EfectoMagico
                    Case e_MagicItemEffect.eProtectedResources
9140                    If ObjData(ObjIndex).ApplyEffectId <= 0 Then
9150                        Exit Sub
9160                    End If
9170                    Call UpdateCd(UserIndex, ObjData(ObjIndex).cdType)
9180                    Call AddOrResetEffect(UserIndex, ObjData(ObjIndex).ApplyEffectId)
9190            End Select
9200        Case e_OBJType.otUsableOntarget
9210            .flags.UsingItemSlot = .flags.TargetObjInvSlot
9220            Call WriteWorkRequestTarget(UserIndex, e_Skill.TargetableItem)
9230    End Select
9240 End With

9250 Exit Sub

hErr:
9260 LogError "Error en useinvitem Usuario: " & UserList(UserIndex).name & " item:" & obj.name & " index: " & UserList(UserIndex).invent.Object(Slot).ObjIndex

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
Private Function IsPotionFreeZone(ByVal UserIndex As Integer, ByVal triggerStatus As e_Trigger6) As Boolean
    Dim currentMap As Integer
    Dim isTriggerZone As Boolean
    Dim isTierUser As Boolean
    Dim isHouseZone As Boolean
    Dim isSpecialZone As Boolean
    Dim isTrainingZone As Boolean
    Dim isArena As Boolean

    ' Obtener el mapa actual del usuario
    currentMap = UserList(UserIndex).pos.Map

    ' Verificar si está en zona con trigger activo
    isTriggerZone = (triggerStatus = e_Trigger6.TRIGGER6_PERMITE)

    ' Verificar si es un usuario con tier de suscripción
    isTierUser = (UserList(UserIndex).Stats.tipoUsuario = tAventurero Or _
                  UserList(UserIndex).Stats.tipoUsuario = tHeroe Or _
                  UserList(UserIndex).Stats.tipoUsuario = tLeyenda)

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
    IsPotionFreeZone = (isHouseZone Or isSpecialZone Or isTrainingZone Or isArena)
End Function

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

Sub SendCraftableElementRunes(ByVal UserIndex As Integer)
        
        On Error GoTo SendCraftableElementRunes_Err
        

100     Call WriteBlacksmithElementalRunes(UserIndex)

        
        Exit Sub

SendCraftableElementRunes_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.SendCraftableElementRunes", Erl)

        
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
        

100     ItemSeCae = (ObjData(Index).Real <> 1 Or ObjData(Index).NoSeCae = 0) And (ObjData(Index).Caos <> 1 Or ObjData(Index).NoSeCae = 0) And ObjData(Index).OBJType <> e_OBJType.otKeys And ObjData(Index).OBJType <> e_OBJType.otShips And ObjData(Index).OBJType <> e_OBJType.otSaddles And ObjData(Index).NoSeCae = 0 And Not ObjData(Index).Intirable = 1 And Not ObjData(Index).Destruye = 1 And Not ObjData(Index).Instransferible = 1

        
        Exit Function

ItemSeCae_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.ItemSeCae", Erl)

        
End Function

Public Function PirataCaeItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

        On Error GoTo PirataCaeItem_Err

100     With UserList(UserIndex)

102         If .clase = e_Class.Pirat And .Stats.ELV >= 37 And .flags.Navegando = 1 Then

                ' Si no está navegando, se caen los items
104             If .invent.EquippedShipObjIndex > 0 Then

                    ' Con galeón cada item tiene una probabilidad de caerse del 67%
106                 If ObjData(.invent.EquippedShipObjIndex).Ropaje = iGaleon Then

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

            
            If ((.Pos.map = 58 Or .Pos.map = 59 Or .Pos.map = 60 Or .Pos.map = 61) And EnEventoFaccionario) Then Exit Sub
            ' Tambien se cae el oro de la billetera
            Dim GoldToDrop As Long
            GoldToDrop = .Stats.GLD - (SvrConfig.GetValue("OroPorNivelBilletera") * .Stats.ELV)
102         If GoldToDrop > 0 And Not EsGM(UserIndex) Then
104             Call TirarOro(GoldToDrop, UserIndex)
            End If
            
106         For i = 1 To .CurrentInventorySlots
    
108             ItemIndex = .Invent.Object(i).ObjIndex

110             If ItemIndex > 0 Then

112                 If ItemSeCae(ItemIndex) And PirataCaeItem(UserIndex, i) And (Not EsNewbie(UserIndex) Or Not ItemNewbie(ItemIndex)) Then
114                     NuevaPos.X = 0
116                     NuevaPos.Y = 0
                    
118                     MiObj.amount = DropAmmount(.invent, i)
120                     MiObj.ObjIndex = ItemIndex
                        MiObj.ElementalTags = .invent.Object(i).ElementalTags
                        
                        If .flags.Navegando Then
128                         Call Tilelibre(.Pos, NuevaPos, MiObj, True, True)
                        Else
129                         Call Tilelibre(.Pos, NuevaPos, MiObj, .flags.Navegando = True, (Not .flags.Navegando) = True)
                            Call ClosestLegalPos(.Pos, NuevaPos, .flags.Navegando, Not .flags.Navegando)
                        End If
130                     If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
132                         Call DropObj(UserIndex, i, MiObj.amount, NuevaPos.Map, NuevaPos.x, NuevaPos.y)
                        
                        '  Si no hay lugar, quemamos el item del inventario (nada de mochilas gratis)
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

Function DropAmmount(ByRef invent As t_Inventario, ByVal objectIndex As Integer) As Integer
100 DropAmmount = invent.Object(objectIndex).amount
102 If invent.EquippedAmuletAccesoryObjIndex > 0 Then
        With ObjData(invent.EquippedAmuletAccesoryObjIndex)
104     If .EfectoMagico = 12 Then
            Dim unprotected As Single
            unprotected = 1
106         If invent.Object(objectIndex).ObjIndex = ORO_MINA Then 'ore types
108             unprotected = CSng(1) - (CSng(.LingO) / 100)
110         ElseIf invent.Object(objectIndex).ObjIndex = PLATA_MINA Then
112             unprotected = CSng(1) - (CSng(.LingP) / 100)
114         ElseIf invent.Object(objectIndex).ObjIndex = HIERRO_MINA Then
116             unprotected = CSng(1) - (CSng(.LingH) / 100)
118         ElseIf invent.Object(objectIndex).ObjIndex = Wood Then ' wood types
120             unprotected = CSng(1) - (CSng(.Madera) / 100)
122         ElseIf invent.Object(objectIndex).ObjIndex = ElvenWood Then
124             unprotected = CSng(1) - (CSng(.MaderaElfica) / 100)
129         ElseIf invent.Object(objectIndex).objIndex = PinoWood Then
130             unprotected = CSng(1) - (CSng(.MaderaPino) / 100)
131         ElseIf invent.Object(objectIndex).objIndex = BLODIUM_MINA Then
132             unprotected = CSng(1) - (CSng(.Blodium) / 100)
            ElseIf invent.Object(objectIndex).ObjIndex > 0 Then 'fish types
                If ObjData(invent.Object(objectIndex).ObjIndex).OBJType = otUseOnce And _
                   ObjData(invent.Object(objectIndex).ObjIndex).Subtipo = e_UseOnceSubType.eFish Then
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

Public Function IsItemInCooldown(ByRef User As t_User, ByRef obj As t_UserOBJ) As Boolean
    Dim elapsedTime As Long
    ElapsedTime = GetTickCount() - User.CdTimes(ObjData(obj.objIndex).cdType)
    IsItemInCooldown = ElapsedTime < ObjData(obj.objIndex).Cooldown
End Function

Public Sub UserTargetableItem(ByVal UserIndex As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
On Error GoTo UserTargetableItem_Err
    With UserList(UserIndex)
        If IsItemInCooldown(UserList(UserIndex), .invent.Object(.flags.UsingItemSlot)) Then
            Exit Sub
        End If
        If .flags.UsingItemSlot = 0 Then Exit Sub
        Dim objIndex As Integer
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
100     Dim CanHelpResult As e_InteractionResult
102     If Not IsValidUserRef(.flags.TargetUser) Then
104         Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.TargetUser.ArrayIndex = UserIndex Then
            Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
114     Dim TargetUser As Integer
116     TargetUser = .flags.TargetUser.ArrayIndex
        If UserList(TargetUser).flags.Muerto = 0 Then
            Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
106     CanHelpResult = UserMod.CanHelpUser(UserIndex, targetUser)
        If UserList(TargetUser).flags.SeguroResu Then
            ' Msg693=El usuario tiene el seguro de resurrección activado.
            Call WriteLocaleMsg(UserIndex, "693", e_FontTypeNames.FONTTYPE_INFO)

            Call WriteConsoleMsg(TargetUser, PrepareMessageLocaleMsg(1985, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1985=¬1 está intentando revivirte. Desactiva el seguro de resurrección para permitirle hacerlo.

            Exit Sub
        End If
        If CanHelpResult <> eInteractionOk Then
            Call SendHelpInteractionMessage(UserIndex, CanHelpResult)
        End If
        
118     Dim costoVidaResu As Long
120     costoVidaResu = UserList(TargetUser).Stats.ELV * 1.5 + .Stats.MinHp * 0.5
    
122     Call UserMod.ModifyHealth(UserIndex, -costoVidaResu, 1)
124     Call ModifyStamina(UserIndex, -UserList(UserIndex).Stats.MinSta, False, 0)
        Dim objIndex As Integer
126     ObjIndex = .invent.Object(.flags.UsingItemSlot).ObjIndex
128     Call UpdateCd(UserIndex, ObjData(objIndex).cdType)
192     Call RemoveItemFromInventory(UserIndex, UserList(UserIndex).flags.UsingItemSlot)
196     Call ResurrectUser(TargetUser)
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
        If Not CanAddTrapAt(.pos.map, TileX, TileY) Then
            Call WriteLocaleMsg(UserIndex, MsgInvalidTile, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim i As Integer
        Dim OlderTrapTime As Long
        Dim OlderTrapIndex As Integer
        OlderTrapTime = 0
        Dim TrapCount As Integer
        Dim Trap As clsTrap
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
        Dim objIndex As Integer
        ObjIndex = UserList(UserIndex).invent.Object(UserList(UserIndex).flags.UsingItemSlot).ObjIndex
        Call UpdateCd(UserIndex, ObjData(objIndex).cdType)
        Call EffectsOverTime.CreateTrap(UserIndex, eUser, .pos.map, TileX, TileY, ObjData(objIndex).EfectoMagico)
        Call RemoveItemFromInventory(UserIndex, UserList(UserIndex).flags.UsingItemSlot)
    End With
End Sub

Public Sub UseArpon(ByVal UserIndex As Integer)
    With UserList(UserIndex)
100     Dim CanAttackResult As e_AttackInteractionResult
        Dim TargetRef As t_AnyReference
        If IsValidUserRef(.flags.targetUser) Then
            Call CastUserToAnyRef(.flags.targetUser, TargetRef)
        Else
            Call CastNpcToAnyRef(.flags.TargetNPC, TargetRef)
        End If
102     If Not IsValidRef(TargetRef) Then
104         Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
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
            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim, .Char.BackpackAnim)
            Exit Sub
        End If
        If .flags.Muerto = 1 Then
            .Char.body = iCuerpoMuerto
204         .Char.head = 0
206         .Char.ShieldAnim = NingunEscudo
208         .Char.WeaponAnim = NingunArma
210         .Char.CascoAnim = NingunCasco
211         .Char.CartAnim = NoCart
            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim, .Char.BackpackAnim)
        Exit Sub
        End If
        
        .Char.head = .OrigChar.head
        If .invent.EquippedWeaponObjIndex > 0 Then
            .Char.WeaponAnim = ObjData(.invent.EquippedWeaponObjIndex).WeaponAnim
        ElseIf .invent.EquippedWorkingToolObjIndex > 0 Then
            .Char.WeaponAnim = ObjData(.invent.EquippedWorkingToolObjIndex).WeaponAnim
        Else
            .Char.WeaponAnim = 0
        End If
        If .invent.EquippedArmorObjIndex > 0 Then
            .Char.body = ObtenerRopaje(UserIndex, ObjData(.invent.EquippedArmorObjIndex))
        Else
            Call SetNakedBody(UserList(UserIndex))
        End If
        If .invent.EquippedHelmetObjIndex > 0 Then
            .Char.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
        Else
            .Char.CascoAnim = 0
        End If
        If .invent.EquippedAmuletAccesoryObjIndex > 0 Then
            .Char.CartAnim = ObjData(.invent.EquippedAmuletAccesoryObjIndex).Ropaje
        Else
            .Char.CartAnim = 0
        End If
        If .invent.EquippedShieldObjIndex > 0 Then
            .Char.ShieldAnim = ObjData(.invent.EquippedShieldObjIndex).ShieldAnim
        Else
            .Char.ShieldAnim = 0
        End If
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim, .Char.BackpackAnim)
    End With
End Sub

Function RemoveGold(ByVal UserIndex As Integer, ByVal Amount As Long) As Boolean
    With UserList(UserIndex)
        If .Stats.GLD < Amount Then Exit Function
        .Stats.GLD = .Stats.GLD - Amount
        Call WriteUpdateGold(UserIndex)
        RemoveGold = True
    End With
End Function

Sub AddGold(ByVal UserIndex As Integer, ByVal Amount As Long)
    With UserList(UserIndex)
        .Stats.GLD = .Stats.GLD + Amount
        Call WriteUpdateGold(UserIndex)
    End With
End Sub

Function ObtenerRopaje(ByVal UserIndex As Integer, ByRef Obj As t_ObjData) As Integer
    Dim Race As e_Raza
    Race = UserList(UserIndex).raza
    
    Dim EsMujer As Boolean
    EsMujer = UserList(UserIndex).genero = e_Genero.Mujer

    Select Case Race
        Case e_Raza.Humano
            If EsMujer And Obj.RopajeHumana > 0 Then
                ObtenerRopaje = Obj.RopajeHumana
                Exit Function
            ElseIf Obj.RopajeHumano > 0 Then
                ObtenerRopaje = Obj.RopajeHumano
                Exit Function
            End If
        Case e_Raza.Elfo
            If EsMujer And Obj.RopajeElfa > 0 Then
                ObtenerRopaje = Obj.RopajeElfa
                Exit Function
            ElseIf Obj.RopajeElfo > 0 Then
                ObtenerRopaje = Obj.RopajeElfo
                Exit Function
            End If
        Case e_Raza.Drow
            If EsMujer And Obj.RopajeElfaOscura > 0 Then
                ObtenerRopaje = Obj.RopajeElfaOscura
                Exit Function
            ElseIf Obj.RopajeElfoOscuro > 0 Then
                ObtenerRopaje = Obj.RopajeElfoOscuro
                Exit Function
            End If
        Case e_Raza.Orco
            If EsMujer And Obj.RopajeOrca > 0 Then
                ObtenerRopaje = Obj.RopajeOrca
                Exit Function
            ElseIf Obj.RopajeOrco > 0 Then
                ObtenerRopaje = Obj.RopajeOrco
                Exit Function
            End If
        Case e_Raza.Enano
            If EsMujer And Obj.RopajeEnana > 0 Then
                ObtenerRopaje = Obj.RopajeEnana
                Exit Function
            ElseIf Obj.RopajeEnano > 0 Then
                ObtenerRopaje = Obj.RopajeEnano
                Exit Function
            End If
        Case e_Raza.Gnomo
            If EsMujer And Obj.RopajeGnoma > 0 Then
                ObtenerRopaje = Obj.RopajeGnoma
                Exit Function
            ElseIf Obj.RopajeGnomo > 0 Then
                ObtenerRopaje = Obj.RopajeGnomo
                Exit Function
            End If
    End Select
    
    ObtenerRopaje = Obj.Ropaje
End Function
Sub EliminarLlaves(ByVal ClaveLlave As Integer, ByVal UserIndex As Integer)
    ' Abrir el archivo "Eliminarllaves.dat" para lectura
    Open "Eliminarllaves.dat" For Input As #1

    ' Variables para el almacenamiento temporal de datos
    Dim Linea As String
    Dim Clave As Integer
    Dim Objeto As Integer
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
        Call WriteLocaleMsg(UserIndex, "2087", e_FontTypeNames.FONTTYPE_INFOIAO)
        Exit Function
    End If
    
    If UserList(UserIndex).invent.Object(TargetSlot).ElementalTags <> e_ElementalTags.Normal Then
        Call WriteLocaleMsg(UserIndex, "2087", e_FontTypeNames.FONTTYPE_INFOIAO)
        Exit Function
    End If
    
    If UserList(UserIndex).invent.Object(TargetSlot).amount > 1 Then
        Call WriteLocaleMsg(UserIndex, "2088", e_FontTypeNames.FONTTYPE_INFOIAO)
        Exit Function
    End If
    
    
    
    Select Case SourceObj.ElementalTags
        Case e_ElementalTags.Fire
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticulasIndex.Incinerar, 10, False))
        Case e_ElementalTags.Water
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticulasIndex.CurarCrimi, 10, False))
        Case e_ElementalTags.Earth
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticulasIndex.Envenena, 10, False))
        Case e_ElementalTags.Wind
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticulasIndex.Runa, 10, False))
        Case Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, e_ParticulasIndex.Curar, 10, False))
    End Select
    
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(e_FXSound.RUNE_SOUND, NO_3D_SOUND, NO_3D_SOUND))
    UserList(UserIndex).invent.Object(TargetSlot).ElementalTags = SourceObj.ElementalTags
    CanElementalTagBeApplied = True
    
End Function

