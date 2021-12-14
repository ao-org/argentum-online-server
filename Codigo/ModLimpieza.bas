Attribute VB_Name = "ModLimpieza"
Option Explicit

Private Const S As String * 1 = ","

Private Item_TimeClear As Long

Dim Item_List As Collection

Public Sub InicializarLimpieza()

    On Error GoTo ErrHandler

100     Set Item_List = New Collection
    
        'Tiempo que puede permanecer un objeto en el suelo (en milisegundos)
102     Item_TimeClear = (TimerLimpiarObjetos * 60000)
    
        Exit Sub
    
ErrHandler:
104     Call TraceError(Err.Number, Err.Description, "ModLimpieza.InicializarLimpieza", Erl)


End Sub

Public Sub LimpiarModuloLimpieza() ' Valga la redundancia
        
    On Error GoTo Class_Terminate_Err
        
100 Set Item_List = Nothing

    Exit Sub

Class_Terminate_Err:
102 Call TraceError(Err.Number, Err.Description, "ModLimpieza.LimpiarModuloLimpieza", Erl)

        
End Sub

Public Sub AgregarItemLimpieza(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal ResetTimer As Boolean = False)

        On Error GoTo hErr
        
        ' Mapas que ignoran limpieza
        Dim i As Integer
100     For i = 1 To UBound(MapasIgnoranLimpieza)
102         If Map = MapasIgnoranLimpieza(i) Then Exit Sub
        Next
    
        Dim Item As TLimpiezaItem

104     If ResetTimer Then
106         Set Item = Item_List.Item(GetIndiceByPos(Map, X, Y))
    
108         Item.Time = GetTickCount()
    
        Else
110         Set Item = New TLimpiezaItem
       
112         With Item
114             .Time = GetTickCount()
116             .Map = Map
118             .X = X
120             .Y = Y
            End With

122         Call Item_List.Add(Item, Item.Indice)
        End If
    
124     Set Item = Nothing
    
        Exit Sub
    
hErr:
126     Call TraceError(Err.Number, Err.Description, "ModLimpieza.AgregarItemLimpiza", Erl)


End Sub

Public Sub QuitarItemLimpieza(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)

    On Error GoTo hErr
    If Item_List.Count > 0 Then
100     Call Item_List.Remove(GetIndiceByPos(Map, X, Y))
    End If
    Exit Sub
    
hErr:
    ' No hace falta registrar el error.
    ' Si un item no existe en la colección, es porque el item era del mapa y alguien lo agarró.
End Sub

Public Sub LimpiarItemsViejos()

        On Error GoTo hErr
    
100     If Item_List Is Nothing Then Exit Sub
    
        Dim TimeClear As Long
102         TimeClear = GetTickCount()
    
        Dim Item As TLimpiezaItem

104     For Each Item In Item_List

106         With Item
108             If TimeClear - .Time >= Item_TimeClear Then
110                 If MapData(.Map, .X, .Y).ObjInfo.ObjIndex > 0 Then ' Por las dudas
112                     Call EraseObj(MAX_INVENTORY_OBJS, .Map, .X, .Y)
                    End If
                End If
            End With

114         Set Item = Nothing
        Next
    
        Exit Sub

hErr:
116     Call TraceError(Err.Number, Err.Description, "ModLimpieza.LimpiarItemsViejos", Erl)


End Sub

Public Sub LimpiezaForzada() ' Limpio todo, no importa el tiempo

        On Error GoTo hErr

        Dim Item As TLimpiezaItem

100     For Each Item In Item_List

102         With Item
104             If MapData(.Map, .X, .Y).ObjInfo.ObjIndex > 0 Then ' Por las dudas
106                 Call EraseObj(MAX_INVENTORY_OBJS, .Map, .X, .Y)
                End If
            End With

108         Set Item = Nothing
        Next
    
        Exit Sub
    
hErr:
110     Call TraceError(Err.Number, Err.Description, "ModLimpieza.LimpiezaForzada", Erl)


End Sub

Public Function GetIndiceByPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As String
        
        On Error GoTo GetIndiceByPos_Err
    
        
100     GetIndiceByPos = Map & S & X & S & Y
        
        Exit Function

GetIndiceByPos_Err:
102     Call TraceError(Err.Number, Err.Description, "ModLimpieza.GetIndiceByPos", Erl)

        
End Function
