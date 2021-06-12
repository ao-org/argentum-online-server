Attribute VB_Name = "modCrafteos"
Option Explicit

Public Const MAX_SLOTS_CRAFTEO = 5

Public Sub SortIntegerArray(Data() As Integer, ByVal First As Integer, ByVal Last As Integer)
    Dim Low As Integer, High As Integer
    Dim MidValue As Integer, Temp As Integer

    Low = First
    High = Last
    MidValue = Data((First + Last) \ 2)
    
    Do
        While Data(Low) < MidValue
            Low = Low + 1
        Wend

        While Data(High) > MidValue
            High = High - 1
        Wend

        If Low <= High Then
            Temp = Data(Low)
            Data(Low) = Data(High)
            Data(High) = Temp
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High

    If First < High Then Call SortIntegerArray(Data, First, High)
    If Low < Last Then Call SortIntegerArray(Data, Low, Last)
End Sub

Public Function GetRecipeKey(Data() As Integer) As String
    Dim i As Integer
    For i = LBound(Data) To UBound(Data)
        GetRecipeKey = GetRecipeKey & Data(i) & ":"
    Next
End Function

Public Sub ReturnCraftingItems(ByVal UserIndex As Integer)
    Dim i As Integer, TmpObj As obj

    With UserList(UserIndex)

        For i = 1 To UBound(.CraftInventory)
            If .CraftInventory(i) <> 0 Then
                TmpObj.ObjIndex = .CraftInventory(i)
                TmpObj.amount = 1
            
                If Not MeterItemEnInventario(UserIndex, TmpObj) Then
                    Call TirarItemAlPiso(.Pos, TmpObj)
                End If
            
                .CraftInventory(i) = 0
            End If
        Next
        
        If .CraftCatalyst.amount > 0 Then
            If Not MeterItemEnInventario(UserIndex, .CraftCatalyst) Then
                Call TirarItemAlPiso(.Pos, .CraftCatalyst)
            End If
        
            .CraftCatalyst.ObjIndex = 0
            .CraftCatalyst.amount = 0
        End If

        Set .CraftResult = Nothing

    End With
End Sub

Public Function CheckCraftingResult(ByVal UserIndex As Integer) As clsCrafteo

    With UserList(UserIndex)
        
        If Not Crafteos.Exists(.flags.Crafteando) Then Exit Function
        
        Dim CrafteosDeEsteTipo As Dictionary
        Set CrafteosDeEsteTipo = Crafteos.Item(.flags.Crafteando)

        Dim Key As String
        Key = GetRecipeKey(.CraftInventory)

        If Not CrafteosDeEsteTipo.Exists(Key) Then Exit Function

        Set CheckCraftingResult = CrafteosDeEsteTipo.Item(Key)

    End With

End Function

Public Sub DoCraftItem(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        If Not .CraftResult Is Nothing Then
            If .CraftResult.Precio > .Stats.GLD Then
                Call WriteConsoleMsg(UserIndex, "No tienes el oro suficiente.", FontTypeNames.FONTTYPE_INFO)
                ' TODO: Mensaje en la ventana de crafteo
                Exit Sub
            End If

            Dim Porcentaje As Byte
            Porcentaje = CalculateCraftProb(UserIndex, .CraftResult.Probabilidad)
            
            If RandomNumber(1, 100) <= Porcentaje Then
                Dim TmpObj As obj
                TmpObj.ObjIndex = .CraftResult.Resultado
                TmpObj.amount = 1
                
                If Not MeterItemEnInventario(UserIndex, TmpObj) Then
                    Call WriteConsoleMsg(UserIndex, "No tenés espacio suficiente en el inventario.", FontTypeNames.FONTTYPE_WARNING)
                    ' TODO: Mensaje en la ventana de crafteo
                    Exit Sub
                End If
            Else
                Call WriteConsoleMsg(UserIndex, "La combinación ha fallado.", FontTypeNames.FONTTYPE_FIGHT)
                ' TODO: Mensaje en la ventana de crafteo y sonido (?
            End If

            .Stats.GLD = .Stats.GLD - .CraftResult.Precio
            Call WriteUpdateGold(UserIndex)

            Dim i As Integer
            For i = 1 To UBound(.CraftInventory)
                .CraftInventory(i) = 0
                Call WriteCraftingItem(UserIndex, i, 0)
            Next
            
            .CraftCatalyst.amount = .CraftCatalyst.amount - 1
            If .CraftCatalyst.amount <= 0 Then .CraftCatalyst.ObjIndex = 0

            Call WriteCraftingCatalyst(UserIndex, .CraftCatalyst.ObjIndex, .CraftCatalyst.amount, 0)
            
            Set .CraftResult = Nothing
            Call WriteCraftingResult(UserIndex, 0)

            Call WriteConsoleMsg(UserIndex, "La combinación ha sido exitosa.", FontTypeNames.FONTTYPE_INFO)
            ' TODO: Mensaje en la ventana de crafteo y sonido (?
        End If

    End With

End Sub

Public Function CalculateCraftProb(ByVal UserIndex As Integer, ByVal PorcentajeBase As Byte) As Byte
    
    With UserList(UserIndex)

        If .CraftCatalyst.ObjIndex <> 0 Then
            If ObjData(.CraftCatalyst.ObjIndex).CatalizadorTipo = .flags.Crafteando Then
                CalculateCraftProb = Clamp(Fix(PorcentajeBase * (1 + ObjData(.CraftCatalyst.ObjIndex).CatalizadorAumento)), 0, 100)
                Exit Function
            End If
        End If

        CalculateCraftProb = PorcentajeBase
    
    End With

End Function
