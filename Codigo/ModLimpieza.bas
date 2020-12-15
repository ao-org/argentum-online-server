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
104     Call RegistrarError(Err.Number, Err.description, "ModLimpieza.InicializarLimpieza")
106     Resume Next

End Sub

Public Sub LimpiarModuloLimpieza() ' Valga la redundancia
        
    On Error GoTo Class_Terminate_Err
        
100 Set Item_List = Nothing

    Exit Sub

Class_Terminate_Err:
102 Call RegistrarError(Err.Number, Err.description, "ModLimpieza.LimpiarModuloLimpieza", Erl)
104 Resume Next
        
End Sub

Public Sub AgregarItemLimpieza(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal ResetTimer As Boolean = False)

        On Error GoTo hErr
    
        Dim Item As TLimpiezaItem

100     If ResetTimer Then
102         Set Item = Item_List.Item(GetIndiceByPos(Map, X, Y))
    
104         Item.Time = GetTickCount()
    
        Else
106         Set Item = New TLimpiezaItem
       
108         With Item
110             .Time = GetTickCount()
112             .Map = Map
114             .X = X
116             .Y = Y
            End With

118         Call Item_List.Add(Item, Item.Indice)
        End If
    
120     Set Item = Nothing
    
        Exit Sub
    
hErr:
122     Call RegistrarError(Err.Number, Err.description, "ModLimpieza.AgregarItemLimpiza")
124     Resume Next

End Sub

Public Sub QuitarItemLimpieza(ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte)

    On Error GoTo hErr

100 Call Item_List.Remove(GetIndiceByPos(Map, X, Y))
    
    Exit Sub
    
hErr:
    ' No hace falta registrar el error.
    ' Si un item no existe en la colección, es porque el item era del mapa y alguien lo agarró.
    'Call RegistrarError(Err.Number, Err.description, "ModLimpieza.QuitarItemLimpieza", Erl)
    'Resume Next

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
116     Call RegistrarError(Err.Number, Err.description, "ModLimpieza.LimpiarItemsViejos")
118     Resume Next

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
110     Call RegistrarError(Err.Number, Err.description, "ModLimpieza.LimpiezaForzada")
112     Resume Next

End Sub

Public Function GetIndiceByPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As String
100     GetIndiceByPos = Map & S & X & S & Y
End Function
