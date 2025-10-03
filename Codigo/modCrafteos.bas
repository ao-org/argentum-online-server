Attribute VB_Name = "modCrafteos"
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
Public Const MAX_SLOTS_CRAFTEO = 5

Public Sub SortIntegerArray(data() As Integer, ByVal First As Integer, ByVal Last As Integer)
    On Error GoTo SortIntegerArray_Err:
    Dim Low      As Integer, High As Integer
    Dim MidValue As Integer, Temp As Integer
    Low = First
    High = Last
    MidValue = data((First + Last) \ 2)
    Do
        While data(Low) < MidValue
            Low = Low + 1
        Wend
        While data(High) > MidValue
            High = High - 1
        Wend
        If Low <= High Then
            Temp = data(Low)
            data(Low) = data(High)
            data(High) = Temp
            Low = Low + 1
            High = High - 1
        End If
    Loop While Low <= High
    If First < High Then Call SortIntegerArray(data, First, High)
    If Low < Last Then Call SortIntegerArray(data, Low, Last)
    Exit Sub
SortIntegerArray_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.SortIntegerArray", Erl)
End Sub

Public Function GetRecipeKey(data() As Integer) As String
    On Error GoTo GetRecipeKey_Err:
    Dim i As Integer
    For i = LBound(data) To UBound(data)
        GetRecipeKey = GetRecipeKey & data(i) & ":"
    Next
    Exit Function
GetRecipeKey_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.GetRecipeKey", Erl)
End Function

Public Sub ReturnCraftingItems(ByVal UserIndex As Integer)
    On Error GoTo ReturnCraftingItems_Err:
    Dim i As Integer, TmpObj As t_Obj
    With UserList(UserIndex)
        For i = 1 To UBound(.CraftInventory)
            If .CraftInventory(i) <> 0 Then
                TmpObj.ObjIndex = .CraftInventory(i)
                TmpObj.amount = 1
                If Not MeterItemEnInventario(UserIndex, TmpObj) Then
                    Call TirarItemAlPiso(.pos, TmpObj)
                End If
                .CraftInventory(i) = 0
            End If
        Next
        If .CraftCatalyst.amount > 0 Then
            If Not MeterItemEnInventario(UserIndex, .CraftCatalyst) Then
                Call TirarItemAlPiso(.pos, .CraftCatalyst)
            End If
            .CraftCatalyst.ObjIndex = 0
            .CraftCatalyst.amount = 0
        End If
        Set .CraftResult = Nothing
    End With
    Exit Sub
ReturnCraftingItems_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.ReturnCraftingItems", Erl)
End Sub

Public Function CheckCraftingResult(ByVal UserIndex As Integer) As clsCrafteo
    On Error GoTo CheckCraftingResult_Err:
    With UserList(UserIndex)
        If Not Crafteos.Exists(.flags.Crafteando) Then Exit Function
        Dim CrafteosDeEsteTipo As Dictionary
        Set CrafteosDeEsteTipo = Crafteos.Item(.flags.Crafteando)
        Dim key As String
        key = GetRecipeKey(.CraftInventory)
        If Not CrafteosDeEsteTipo.Exists(key) Then Exit Function
        Set CheckCraftingResult = CrafteosDeEsteTipo.Item(key)
    End With
    Exit Function
CheckCraftingResult_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.CheckCraftingResult", Erl)
End Function

Public Sub DoCraftItem(ByVal UserIndex As Integer)
    On Error GoTo DoCraftItem_Err:
    With UserList(UserIndex)
        If Not .CraftResult Is Nothing Then
            If .CraftResult.precio > .Stats.GLD Then
                ' Msg588=No tienes el oro suficiente.
                Call WriteLocaleMsg(UserIndex, 588, e_FontTypeNames.FONTTYPE_INFO)
                ' TODO: Mensaje en la ventana de crafteo
                Exit Sub
            End If
            Dim Porcentaje As Byte
            Porcentaje = CalculateCraftProb(UserIndex, .CraftResult.Probabilidad)
            If RandomNumber(1, 100) <= Porcentaje Then
                Dim TmpObj As t_Obj
                TmpObj.ObjIndex = .CraftResult.Resultado
                TmpObj.amount = 1
                If Not MeterItemEnInventario(UserIndex, TmpObj) Then
                    ' Msg589=No tenés espacio suficiente en el inventario.
                    Call WriteLocaleMsg(UserIndex, 589, e_FontTypeNames.FONTTYPE_WARNING)
                    ' TODO: Mensaje en la ventana de crafteo
                    Exit Sub
                End If
                ' Msg590=La combinación ha sido exitosa.
                Call WriteLocaleMsg(UserIndex, 590, e_FontTypeNames.FONTTYPE_INFO)
                ' TODO: Mensaje en la ventana de crafteo y sonido (?
            Else
                'Msg923= La combinación ha fallado.
                Call WriteLocaleMsg(UserIndex, 923, e_FontTypeNames.FONTTYPE_FIGHT)
                ' TODO: Mensaje en la ventana de crafteo y sonido (?
            End If
            .Stats.GLD = .Stats.GLD - .CraftResult.precio
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
        End If
    End With
    Exit Sub
DoCraftItem_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.DoCraftItem", Erl)
End Sub

Public Function CalculateCraftProb(ByVal UserIndex As Integer, ByVal PorcentajeBase As Byte) As Byte
    On Error GoTo CalculateCraftProb_Err:
    With UserList(UserIndex)
        If .CraftCatalyst.ObjIndex <> 0 Then
            If ObjData(.CraftCatalyst.ObjIndex).CatalizadorTipo = .flags.Crafteando Then
                CalculateCraftProb = Clamp(Fix(PorcentajeBase * (1 + ObjData(.CraftCatalyst.ObjIndex).CatalizadorAumento)), 0, 100)
                Exit Function
            End If
        End If
        CalculateCraftProb = PorcentajeBase
    End With
    Exit Function
CalculateCraftProb_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.CalculateCraftProb", Erl)
End Function
