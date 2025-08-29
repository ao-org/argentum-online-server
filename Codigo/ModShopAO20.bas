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


Public Sub init_transaction(ByVal obj_num As Long, ByVal userindex As Integer)
    On Error Goto init_transaction_Err
On Error GoTo init_transaction_Err
    Dim obj As t_ObjData
    
100 obj.ObjNum = obj_num
102 With UserList(userIndex)
        
        'Me fijo si es un item de shop
104     If Not is_purchaseable_item(obj) Then
            'Msg1087= Error al realizar la transacción
            Call WriteLocaleMsg(UserIndex, "1087", e_FontTypeNames.FONTTYPE_INFO)
108         Call LogShopErrors("El usuario " & .name & " intentó comprar un objeto que no es de shop (REVISAR) | " & obj.name)
            Exit Sub
        End If
        Call LoadPatronCreditsFromDB(UserIndex)
110     If obj.Valor > .Stats.Creditos Then
            'Msg1088= Error al realizar la transacción.
            Call WriteLocaleMsg(UserIndex, "1088", e_FontTypeNames.FONTTYPE_INFO)
114         Call LogShopErrors("El usuario " & .name & " intentó editar el valor del objeto (REVISAR) | " & obj.name)
            Exit Sub
        End If
        
        'Me fijo si tiene espacio en el inventario
        Dim objInventario As t_Obj
        
116     objInventario.amount = 1
118     objInventario.objIndex = obj.ObjNum
        
        If GetSlotForItemInInventory(UserIndex, objInventario) <= 0 Then
            'Msg1089= Asegurate de tener espacio suficiente en tu inventario.
            Call WriteLocaleMsg(UserIndex, "1089", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
120     'Descuento los créditos
124     .Stats.Creditos = .Stats.Creditos - obj.Valor
          
        'Genero un log de los créditos que gastó y cuantos le quedan luego de la transacción.
126     Call LogShopTransactions(.name & " | Compró -> " & ObjData(obj.ObjNum).name & " | Valor -> " & obj.Valor)
128     Call Query("update account set offline_patron_credits = ? where id = ?;", .Stats.Creditos, .AccountID)
130     Call writeUpdateShopClienteCredits(UserIndex)
132     Call RegisterTransaction(.AccountID, .ID, obj.ObjNum, obj.Valor, .Stats.Creditos)
        Call MeterItemEnInventario(UserIndex, objInventario)
    End With
    Exit Sub
init_transaction_Err:
    Call TraceError(Err.Number, Err.Description, "ShopAo20.init_transaction", Erl)
    Exit Sub
init_transaction_Err:
    Call TraceError(Err.Number, Err.Description, "ModShopAO20.init_transaction", Erl)
End Sub

Private Function is_purchaseable_item(ByRef obj As t_ObjData) As Boolean
    On Error Goto is_purchaseable_item_Err
    Dim i As Long
    
    For i = 1 To UBound(ObjShop)
        If ObjShop(i).ObjNum = obj.ObjNum Then
            'Si es un item de shop, aparte le agrego el valor (por ref)
            obj.valor = ObjShop(i).valor
            is_purchaseable_item = True
            Exit Function
        End If
    Next i
    
    is_purchaseable_item = False
    
    Exit Function
is_purchaseable_item_Err:
    Call TraceError(Err.Number, Err.Description, "ModShopAO20.is_purchaseable_item", Erl)
End Function

Private Sub RegisterTransaction(ByVal AccId As Long, ByVal CharId As Long, ByVal itemId As Long, ByVal Price As Long, ByVal CreditLeft As Long)
    On Error Goto RegisterTransaction_Err
On Error GoTo RegisterTransaction_Err
100 Call Query("insert into patreon_shop_audit (acc_id, char_id, item_id, price, credit_left, time) VALUES (?,?,?,?,?, STRFTIME('%s'));", AccId, CharId, itemId, price, CreditLeft)
    Exit Sub
RegisterTransaction_Err:
    Call TraceError(Err.Number, Err.Description, "ShopAo20.RegisterTransaction", Erl)
    Exit Sub
RegisterTransaction_Err:
    Call TraceError(Err.Number, Err.Description, "ModShopAO20.RegisterTransaction", Erl)
End Sub
