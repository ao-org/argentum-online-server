Attribute VB_Name = "ModLimpieza"
Option Explicit

Private Const S As String * 1 = ","

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Item_TimeClear As Long

Dim Item_List As Collection

Public Sub InicializarLimpieza()

On Error GoTo Errhandler

    Set Item_List = New Collection
    
    'Tiempo que puede permanecer un objeto en el suelo (en milisegundos)
    Item_TimeClear = (TimerLimpiarObjetos * 60000)
    
    Exit Sub
    
Errhandler:
    Call RegistrarError(Err.Number, Err.description, "ModLimpieza.InicializarLimpieza")
    Resume Next

End Sub

Public Sub LimpiarModuloLimpieza() ' Valga la redundancia
        
On Error GoTo Class_Terminate_Err
        
100 Set Item_List = Nothing

    Exit Sub

Class_Terminate_Err:
    Call RegistrarError(Err.Number, Err.description, "ModLimpieza.LimpiarModuloLimpieza", Erl)
    Resume Next
        
End Sub

Public Sub AgregarItemLimpiza(ByVal Map As Integer, ByVal x As Byte, ByVal Y As Byte, Optional ByVal ResetTimer As Boolean = False)

    On Error GoTo hErr
    
    Dim Item As TLimpiezaItem

    If ResetTimer Then
        Set Item = Item_List.Item(Item.Indice)
    
        Item.Time = GetTickCount And &H7FFFFFFF
    
    Else
        Set Item = New TLimpiezaItem
       
        With Item
            .Time = GetTickCount And &H7FFFFFFF
            .Map = Map
            .x = x
            .Y = Y
        End With

        Call Item_List.Add(Item, Item.Indice)
    End If
    
    Set Item = Nothing
    
    Exit Sub
    
hErr:
    Call RegistrarError(Err.Number, Err.description, "ModLimpieza.AgregarItemLimpiza")
    Resume Next

End Sub

Public Sub QuitarItemLimpieza(ByVal Map As Integer, ByVal x As Byte, ByVal Y As Byte)

    On Error GoTo hErr

101 Call Item_List.Remove(GetIndiceByPos(Map, x, Y))
    
    Exit Sub
    
hErr:
    ' No hace falta registrar el error.
    ' Si un item no existe en la colección, es porque el item era del mapa y alguien lo agarró.
    'Call RegistrarError(Err.Number, Err.description, "ModLimpieza.QuitarItemLimpieza", Erl)
    Resume Next

End Sub

Public Sub LimpiarItemsViejos()

    On Error GoTo hErr
    
    If Item_List Is Nothing Then Exit Sub
    
    Dim TimeClear As Long
        TimeClear = GetTickCount And &H7FFFFFFF
    
    Dim Item As TLimpiezaItem

    For Each Item In Item_List

        With Item
            If TimeClear - .Time >= Item_TimeClear Then
                If MapData(.Map, .x, .Y).ObjInfo.ObjIndex > 0 Then ' Por las dudas
                    Call EraseObj(MAX_INVENTORY_OBJS, .Map, .x, .Y)
                End If
            End If
        End With

        Set Item = Nothing
    Next
    
    Exit Sub

hErr:
    Call RegistrarError(Err.Number, Err.description, "ModLimpieza.LimpiarItemsViejos")
    Resume Next

End Sub

Public Sub LimpiezaForzada() ' Limpio todo, no importa el tiempo

    On Error GoTo hErr

    Dim Item As TLimpiezaItem

    For Each Item In Item_List

        With Item
            If MapData(.Map, .x, .Y).ObjInfo.ObjIndex > 0 Then ' Por las dudas
                Call EraseObj(MAX_INVENTORY_OBJS, .Map, .x, .Y)
            End If
        End With

        Set Item = Nothing
    Next
    
    Exit Sub
    
hErr:
    Call RegistrarError(Err.Number, Err.description, "ModLimpieza.LimpiezaForzada")
    Resume Next

End Sub

Public Function GetIndiceByPos(ByVal Map As Integer, ByVal x As Integer, ByVal Y As Integer) As String
    GetIndiceByPos = Map & S & x & S & Y
End Function
