Attribute VB_Name = "ModShopAO20"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit


Public Sub init_transaction(ByVal obj_num As Long, ByVal userindex As Integer)
On Error GoTo init_transaction_Err
    Dim obj As t_ObjData
    
100 obj.ObjNum = obj_num
102 With UserList(userIndex)
        
        'Me fijo si es un item de shop
104     If Not is_purchaseable_item(obj) Then
106         Call WriteConsoleMsg(userIndex, "Error al realizar la transacción", e_FontTypeNames.FONTTYPE_INFO)
108         Call LogShopErrors("El usuario " & .name & " intentó comprar un objeto que no es de shop (REVISAR) | " & obj.name)
            Exit Sub
        End If
        
110     If obj.Valor > .Stats.Creditos Then
112         Call WriteConsoleMsg(userIndex, "Error al realizar la transacción.", e_FontTypeNames.FONTTYPE_INFO)
114         Call LogShopErrors("El usuario " & .name & " intentó editar el valor del objeto (REVISAR) | " & obj.name)
            Exit Sub
        End If
        
        'Me fijo si tiene espacio en el inventario
        Dim objInventario As t_Obj
        
116     objInventario.amount = 1
118     objInventario.objIndex = obj.ObjNum
        
120     If Not MeterItemEnInventario(userIndex, objInventario) Then
122         Call WriteConsoleMsg(userIndex, "Asegurate de tener espacio suficiente en tu inventario.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            'Descuento los créditos
124         .Stats.Creditos = .Stats.Creditos - obj.Valor
            
            'Genero un log de los créditos que gastó y cuantos le quedan luego de la transacción.
126         Call LogShopTransactions(.name & " | Compró -> " & ObjData(obj.ObjNum).name & " | Valor -> " & obj.Valor)
128         Call Execute("update user set credits = ? where id = ?;", .Stats.Creditos, .ID)
130         Call writeUpdateShopClienteCredits(userIndex)
132         Call RegisterTransaction(.AccountID, .ID, obj.ObjNum, obj.Valor, .Stats.Creditos)
        End If
                
    End With
init_transaction_Err:
    Call TraceError(Err.Number, Err.Description, "ShopAo20.init_transaction", Erl)
End Sub

Private Function is_purchaseable_item(ByRef obj As t_ObjData) As Boolean
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
    
End Function

Private Sub RegisterTransaction(ByVal AccId As Long, ByVal CharId As Long, ByVal itemId As Long, ByVal Price As Long, ByVal CreditLeft As Long)
On Error GoTo RegisterTransaction_Err
100 Call Execute("insert into patreon_shop_audit (acc_id, char_id, item_id, price, credit_left, time) VALUES (?,?,?,?,?, STRFTIME('%s'));", AccId, CharId, itemId, Price, CreditLeft)
    Exit Sub
RegisterTransaction_Err:
    Call TraceError(Err.Number, Err.Description, "ShopAo20.RegisterTransaction", Erl)
End Sub
