Attribute VB_Name = "ModShopAO20"
Option Explicit


Public Sub init_transaction(ByVal obj_num As Long, ByVal userindex As Integer)
    
    Dim obj As t_ObjData
    
    obj.ObjNum = obj_num
    With UserList(userindex)
        
        'Me fijo si es un item de shop
        If Not is_purchaseable_item(obj) Then
            Call WriteConsoleMsg(userindex, "Error al realizar la transacción", e_FontTypeNames.FONTTYPE_INFO)
            Call LogShopErrors("El usuario " & .Name & " intentó comprar un objeto que no es de shop (REVISAR) | " & obj.Name)
            Exit Sub
        End If
        
        If obj.valor > .Stats.Creditos Then
            Call WriteConsoleMsg(userindex, "Error al realizar la transacción.", e_FontTypeNames.FONTTYPE_INFO)
            Call LogShopErrors("El usuario " & .Name & " intentó editar el valor del objeto (REVISAR) | " & obj.Name)
            Exit Sub
        End If
        
        'Me fijo si tiene espacio en el inventario
        Dim objInventario As t_Obj
        
        objInventario.amount = 1
        objInventario.ObjIndex = obj.ObjNum
        
        If Not MeterItemEnInventario(userindex, objInventario) Then
            Call WriteConsoleMsg(userindex, "Asegurate de tener espacio suficiente en tu inventario.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            'Descuento los créditos
            .Stats.Creditos = .Stats.Creditos - obj.valor
            
            'Genero un log de los créditos que gastó y cuantos le quedan luego de la transacción.
            Call LogShopTransactions(.Name & " | Compró -> " & ObjData(obj.ObjNum).Name & " | Valor -> " & obj.valor)
            Call Execute("update user set credits = ? where id = ?;", .Stats.Creditos, .ID)
            Call writeUpdateShopClienteCredits(userindex)
        End If
                
    End With
        
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
