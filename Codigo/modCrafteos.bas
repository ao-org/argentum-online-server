Attribute VB_Name = "modCrafteos"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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

Public Function FindCraftingRecipe(ByVal CraftingType As Byte, ByVal ResultObjIndex As Integer) As clsCrafteo
    On Error GoTo FindCraftingRecipe_Err
    If Not Crafteos.Exists(CraftingType) Then Exit Function
    Dim Recipes As Dictionary
    Set Recipes = Crafteos.Item(CraftingType)
    Dim Recipe As clsCrafteo
    Dim RecipeValue As Variant
    For Each RecipeValue In Recipes.Items
        Set Recipe = RecipeValue
        If Recipe.Resultado = ResultObjIndex Then
            Set FindCraftingRecipe = Recipe
            Exit Function
        End If
    Next RecipeValue
    Exit Function
FindCraftingRecipe_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.FindCraftingRecipe", Erl)
End Function

Private Function HasNpcCraftingMaterials(ByVal UserIndex As Integer, ByRef Recipe As clsCrafteo) As Boolean
    On Error GoTo HasNpcCraftingMaterials_Err
    Dim ItemIndex As Integer, Required As Integer, i As Integer
    i = 1
    Do While i <= Recipe.Tamaño
        ItemIndex = Recipe.Items(i)
        Required = 1
        Do While i + Required <= Recipe.Tamaño
            If Recipe.Items(i + Required) <> ItemIndex Then Exit Do
            Required = Required + 1
        Loop
        If ItemIndex <= 0 Or Not TieneObjetos(ItemIndex, Required, UserIndex) Then Exit Function
        i = i + Required
    Loop
    HasNpcCraftingMaterials = True
    Exit Function
HasNpcCraftingMaterials_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.HasNpcCraftingMaterials", Erl)
End Function

Private Sub ConsumeNpcCraftingMaterials(ByVal UserIndex As Integer, ByRef Recipe As clsCrafteo)
    On Error GoTo ConsumeNpcCraftingMaterials_Err
    Dim ItemIndex As Integer, Required As Integer, i As Integer
    i = 1
    Do While i <= Recipe.Tamaño
        ItemIndex = Recipe.Items(i)
        Required = 1
        Do While i + Required <= Recipe.Tamaño
            If Recipe.Items(i + Required) <> ItemIndex Then Exit Do
            Required = Required + 1
        Loop
        Call QuitarObjetos(ItemIndex, Required, UserIndex)
        i = i + Required
    Loop
    Exit Sub
ConsumeNpcCraftingMaterials_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.ConsumeNpcCraftingMaterials", Erl)
End Sub

Public Sub DoNpcCraftItem(ByVal UserIndex As Integer, ByVal ResultObjIndex As Integer)
    On Error GoTo DoNpcCraftItem_Err
    With UserList(UserIndex)
        If .flags.Crafteando = 0 Then Exit Sub
        If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
        Dim NpcIndex As Integer
        NpcIndex = .flags.TargetNPC.ArrayIndex
        If NpcList(NpcIndex).Craftea <> .flags.Crafteando Then Exit Sub
        If Distancia(.pos, NpcList(NpcIndex).pos) > 3 Then Exit Sub
        If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub
        Dim Recipe As clsCrafteo
        Set Recipe = FindCraftingRecipe(.flags.Crafteando, ResultObjIndex)
        If Recipe Is Nothing Then Exit Sub
        If Recipe.precio > .Stats.GLD Then
            Call WriteLocaleMsg(UserIndex, MSG_NO_TIENES_ORO_SUFICIENTE, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_New_Naranja)
            Exit Sub
        End If
        If Not HasNpcCraftingMaterials(UserIndex, Recipe) Then Exit Sub
        Call ConsumeNpcCraftingMaterials(UserIndex, Recipe)
        If RandomNumber(1, 100) <= Recipe.Probabilidad Then
            Dim CraftedObject As t_Obj
            CraftedObject.ObjIndex = Recipe.Resultado
            CraftedObject.amount = 1
            CraftedObject.ElementalTags = ObjData(Recipe.Resultado).ElementalTags
            If Not MeterItemEnInventario(UserIndex, CraftedObject) Then
                Call TirarItemAlPiso(.pos, CraftedObject)
            End If
            Call WriteLocaleMsg(UserIndex, MSG_COMBINACION_HA_SIDO_EXITOSA, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_SUBASTA)
        Else
            Call WriteLocaleMsg(UserIndex, MSG_COMBINACION_HA_FALLADO, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_SUBASTA)
        End If
        If Recipe.precio > 0 Then
            .Stats.GLD = .Stats.GLD - Recipe.precio
            Call WriteUpdateGold(UserIndex)
        End If
        Call UpdateUserInv(True, UserIndex, 0)
    End With
    Exit Sub
DoNpcCraftItem_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.DoNpcCraftItem", Erl)
End Sub

Public Sub DoCraftItem(ByVal UserIndex As Integer)
    On Error GoTo DoCraftItem_Err:
    With UserList(UserIndex)
        If Not .CraftResult Is Nothing Then
            If .CraftResult.precio > .Stats.GLD Then
                ' Msg588=No tienes el oro suficiente.
                Call WriteLocaleMsg(UserIndex, MSG_NO_TIENES_ORO_SUFICIENTE, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_New_Naranja)
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
                    Call WriteLocaleMsg(UserIndex, MSG_NO_TENES_ESPACIO_SUFICIENTE_INVENTARIO, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_New_Naranja)
                    ' TODO: Mensaje en la ventana de crafteo
                    Exit Sub
                End If
                ' Msg590=La combinación ha sido exitosa.
                Call WriteLocaleMsg(UserIndex, MSG_COMBINACION_HA_SIDO_EXITOSA, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_SUBASTA)
                ' TODO: Mensaje en la ventana de crafteo y sonido (?
            Else
                'Msg923= La combinación ha fallado.
                Call WriteLocaleMsg(UserIndex, MSG_COMBINACION_HA_FALLADO, e_TextChannel.TEXTCHANNEL_ECONOMY, e_FontTypeNames.FONTTYPE_SUBASTA)
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
