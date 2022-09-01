Attribute VB_Name = "CustomMaps"
Option Explicit

Private CustomMapList As Dictionary

Public Sub InitializeCustomMaps()
    Set CustomMapList = New Dictionary
End Sub

Public Function GetMap(ByVal mapIndex As Integer) As IBaseMap
    Set GetMap = Nothing
    If CustomMapList.Exists(mapIndex) Then
        Set GetMap = CustomMapList.Item(mapIndex)
    End If
End Function

