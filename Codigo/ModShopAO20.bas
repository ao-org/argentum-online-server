Attribute VB_Name = "ModShopAO20"
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

Public Sub init_transaction(ByVal obj_num As Long, ByVal UserIndex As Integer)
    On Error GoTo init_transaction_Err
    Dim obj As t_ObjData
    obj.ObjNum = obj_num
    With UserList(UserIndex)
        'Me fijo si es un item de shop
        If Not is_purchaseable_item(obj) Then
            'Msg1087= Error al realizar la transacción
            Call WriteLocaleMsg(UserIndex, 1087, e_FontTypeNames.FONTTYPE_INFO)
            Call LogShopErrors("El usuario " & .name & " intentó comprar un objeto que no es de shop (REVISAR) | " & obj.name)
            Exit Sub
        End If
        Call LoadPatronCreditsFromDB(UserIndex)
        If obj.Valor > .Stats.Creditos Then
            'Msg1088= Error al realizar la transacción.
            Call WriteLocaleMsg(UserIndex, 1088, e_FontTypeNames.FONTTYPE_INFO)
            Call LogShopErrors("El usuario " & .name & " intentó editar el valor del objeto (REVISAR) | " & obj.name)
            Exit Sub
        End If
        'Me fijo si tiene espacio en el inventario
        Dim objInventario As t_Obj
        objInventario.amount = 1
        objInventario.ObjIndex = obj.ObjNum
        If GetSlotForItemInInventory(UserIndex, objInventario) <= 0 Then
            'Msg1089= Asegurate de tener espacio suficiente en tu inventario.
            Call WriteLocaleMsg(UserIndex, 1089, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Descuento los créditos
        .Stats.Creditos = .Stats.Creditos - obj.Valor
        'Genero un log de los créditos que gastó y cuantos le quedan luego de la transacción.
        Call LogShopTransactions(.name & " | Compró -> " & ObjData(obj.ObjNum).name & " | Valor -> " & obj.Valor)
        Call Query("update account set offline_patron_credits = ? where id = ?;", .Stats.Creditos, .AccountID)
        Call writeUpdateShopClienteCredits(UserIndex)
        Call RegisterTransaction(.AccountID, .Id, obj.ObjNum, obj.Valor, .Stats.Creditos)
        Call MeterItemEnInventario(UserIndex, objInventario)
    End With
    Exit Sub
init_transaction_Err:
    Call TraceError(Err.Number, Err.Description, "ShopAo20.init_transaction", Erl)
End Sub

Private Function is_purchaseable_item(ByRef obj As t_ObjData) As Boolean
    Dim i As Long
    For i = 1 To UBound(ObjShop)
        If ObjShop(i).ObjNum = obj.ObjNum Then
            'Si es un item de shop, aparte le agrego el valor (por ref)
            obj.Valor = ObjShop(i).Valor
            is_purchaseable_item = True
            Exit Function
        End If
    Next i
    is_purchaseable_item = False
End Function

Private Sub RegisterTransaction(ByVal AccId As Long, ByVal CharId As Long, ByVal ItemId As Long, ByVal price As Long, ByVal CreditLeft As Long)
    On Error GoTo RegisterTransaction_Err
    Call Query("insert into patreon_shop_audit (acc_id, char_id, item_id, price, credit_left, time) VALUES (?,?,?,?,?, STRFTIME('%s'));", AccId, CharId, ItemId, price, CreditLeft)
    Exit Sub
RegisterTransaction_Err:
    Call TraceError(Err.Number, Err.Description, "ShopAo20.RegisterTransaction", Erl)
End Sub
