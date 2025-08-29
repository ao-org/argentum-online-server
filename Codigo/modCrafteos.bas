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

Public Sub SortIntegerArray(Data() As Integer, ByVal First As Integer, ByVal Last As Integer)
    On Error Goto SortIntegerArray_Err
    
        On Error GoTo SortIntegerArray_Err:
    
        Dim Low As Integer, High As Integer
        Dim MidValue As Integer, Temp As Integer

100     Low = First
102     High = Last
104     MidValue = data((First + Last) \ 2)
    
        Do
106         While data(Low) < MidValue
108             Low = Low + 1
            Wend

110         While data(High) > MidValue
112             High = High - 1
            Wend

114         If Low <= High Then
116             Temp = data(Low)
118             data(Low) = data(High)
120             data(High) = Temp
122             Low = Low + 1
124             High = High - 1
            End If
126     Loop While Low <= High

128     If First < High Then Call SortIntegerArray(data, First, High)
130     If Low < Last Then Call SortIntegerArray(data, Low, Last)
    
        Exit Sub
    
SortIntegerArray_Err:
132     Call TraceError(Err.Number, Err.Description, "modCrafteos.SortIntegerArray", Erl)
    
    Exit Sub
SortIntegerArray_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.SortIntegerArray", Erl)
End Sub

Public Function GetRecipeKey(Data() As Integer) As String
    On Error Goto GetRecipeKey_Err
    
        On Error GoTo GetRecipeKey_Err:
    
        Dim i As Integer
100     For i = LBound(data) To UBound(data)
102         GetRecipeKey = GetRecipeKey & data(i) & ":"
        Next
    
        Exit Function
    
GetRecipeKey_Err:
104     Call TraceError(Err.Number, Err.Description, "modCrafteos.GetRecipeKey", Erl)
    
    Exit Function
GetRecipeKey_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.GetRecipeKey", Erl)
End Function

Public Sub ReturnCraftingItems(ByVal UserIndex As Integer)
    On Error Goto ReturnCraftingItems_Err
    
        On Error GoTo ReturnCraftingItems_Err:
    
        Dim i As Integer, TmpObj As t_Obj

100     With UserList(UserIndex)

102         For i = 1 To UBound(.CraftInventory)
104             If .CraftInventory(i) <> 0 Then
106                 TmpObj.ObjIndex = .CraftInventory(i)
108                 TmpObj.amount = 1
            
110                 If Not MeterItemEnInventario(UserIndex, TmpObj) Then
112                     Call TirarItemAlPiso(.Pos, TmpObj)
                    End If
            
114                 .CraftInventory(i) = 0
                End If
            Next
        
116         If .CraftCatalyst.amount > 0 Then
118             If Not MeterItemEnInventario(UserIndex, .CraftCatalyst) Then
120                 Call TirarItemAlPiso(.Pos, .CraftCatalyst)
                End If
        
122             .CraftCatalyst.ObjIndex = 0
124             .CraftCatalyst.amount = 0
            End If

126         Set .CraftResult = Nothing

        End With
        
        Exit Sub
    
ReturnCraftingItems_Err:
128     Call TraceError(Err.Number, Err.Description, "modCrafteos.ReturnCraftingItems", Erl)
    
    Exit Sub
ReturnCraftingItems_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.ReturnCraftingItems", Erl)
End Sub

Public Function CheckCraftingResult(ByVal UserIndex As Integer) As clsCrafteo
    On Error Goto CheckCraftingResult_Err
    
        On Error GoTo CheckCraftingResult_Err:
    
100     With UserList(UserIndex)
        
102         If Not Crafteos.Exists(.flags.Crafteando) Then Exit Function
        
            Dim CrafteosDeEsteTipo As Dictionary
104         Set CrafteosDeEsteTipo = Crafteos.Item(.flags.Crafteando)

            Dim Key As String
106         Key = GetRecipeKey(.CraftInventory)

108         If Not CrafteosDeEsteTipo.Exists(Key) Then Exit Function

110         Set CheckCraftingResult = CrafteosDeEsteTipo.Item(Key)

        End With

        Exit Function
    
CheckCraftingResult_Err:
112     Call TraceError(Err.Number, Err.Description, "modCrafteos.CheckCraftingResult", Erl)
    
    Exit Function
CheckCraftingResult_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.CheckCraftingResult", Erl)
End Function

Public Sub DoCraftItem(ByVal UserIndex As Integer)
    On Error Goto DoCraftItem_Err
    
        On Error GoTo DoCraftItem_Err:
    
100     With UserList(UserIndex)

102         If Not .CraftResult Is Nothing Then
104             If .CraftResult.Precio > .Stats.GLD Then
106                 ' Msg588=No tienes el oro suficiente.
                    Call WriteLocaleMsg(UserIndex, "588", e_FontTypeNames.FONTTYPE_INFO)
                    ' TODO: Mensaje en la ventana de crafteo
                    Exit Sub
                End If

                Dim Porcentaje As Byte
108             Porcentaje = CalculateCraftProb(UserIndex, .CraftResult.Probabilidad)
            
110             If RandomNumber(1, 100) <= Porcentaje Then
                    Dim TmpObj As t_Obj
112                 TmpObj.ObjIndex = .CraftResult.Resultado
114                 TmpObj.amount = 1
                
116                 If Not MeterItemEnInventario(UserIndex, TmpObj) Then
118                     ' Msg589=No tenés espacio suficiente en el inventario.
                        Call WriteLocaleMsg(UserIndex, "589", e_FontTypeNames.FONTTYPE_WARNING)
                        ' TODO: Mensaje en la ventana de crafteo
                        Exit Sub
                    End If
                
120                 ' Msg590=La combinación ha sido exitosa.
                    Call WriteLocaleMsg(UserIndex, "590", e_FontTypeNames.FONTTYPE_INFO)
                    ' TODO: Mensaje en la ventana de crafteo y sonido (?
                Else
                    'Msg923= La combinación ha fallado.
                    Call WriteLocaleMsg(UserIndex, "923", e_FontTypeNames.FONTTYPE_FIGHT)
                    ' TODO: Mensaje en la ventana de crafteo y sonido (?
                End If

124             .Stats.GLD = .Stats.GLD - .CraftResult.Precio
126             Call WriteUpdateGold(UserIndex)

                Dim i As Integer
128             For i = 1 To UBound(.CraftInventory)
130                 .CraftInventory(i) = 0
132                 Call WriteCraftingItem(UserIndex, i, 0)
                Next
            
134             .CraftCatalyst.amount = .CraftCatalyst.amount - 1
136             If .CraftCatalyst.amount <= 0 Then .CraftCatalyst.ObjIndex = 0

138             Call WriteCraftingCatalyst(UserIndex, .CraftCatalyst.ObjIndex, .CraftCatalyst.amount, 0)
            
140             Set .CraftResult = Nothing
142             Call WriteCraftingResult(UserIndex, 0)
            End If

        End With

        Exit Sub
    
DoCraftItem_Err:
144     Call TraceError(Err.Number, Err.Description, "modCrafteos.DoCraftItem", Erl)
    
    Exit Sub
DoCraftItem_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.DoCraftItem", Erl)
End Sub

Public Function CalculateCraftProb(ByVal UserIndex As Integer, ByVal PorcentajeBase As Byte) As Byte
    On Error Goto CalculateCraftProb_Err
    
        On Error GoTo CalculateCraftProb_Err:
    
100     With UserList(UserIndex)

102         If .CraftCatalyst.ObjIndex <> 0 Then
104             If ObjData(.CraftCatalyst.ObjIndex).CatalizadorTipo = .flags.Crafteando Then
106                 CalculateCraftProb = Clamp(Fix(PorcentajeBase * (1 + ObjData(.CraftCatalyst.ObjIndex).CatalizadorAumento)), 0, 100)
                    Exit Function
                End If
            End If

108         CalculateCraftProb = PorcentajeBase
    
        End With

        Exit Function
    
CalculateCraftProb_Err:
110     Call TraceError(Err.Number, Err.Description, "modCrafteos.CalculateCraftProb", Erl)
    
    Exit Function
CalculateCraftProb_Err:
    Call TraceError(Err.Number, Err.Description, "modCrafteos.CalculateCraftProb", Erl)
End Function
