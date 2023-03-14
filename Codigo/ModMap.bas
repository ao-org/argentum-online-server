Attribute VB_Name = "ModMap"
Option Explicit

Public Function CanAddTrapAt(ByVal mapIndex As Integer, ByVal posX As Integer, ByVal posY As Integer) As Boolean
    If Not MapData(mapIndex, posX, posY).Trap Is Nothing Then
        Exit Function
    End If
    If MapData(mapIndex, posX, posY).Blocked Then Exit Function
    If MapData(mapIndex, posX, posY).npcIndex > 0 Then Exit Function
    If MapData(mapIndex, posX, posY).UserIndex > 0 Then Exit Function
    If MapData(mapIndex, posX, posY).ObjInfo.objIndex > 0 Then Exit Function
    CanAddTrapAt = True
End Function

Public Sub ActivateTrap(ByVal TargetIndex, ByVal TargetType As e_ReferenceType, ByVal map As Integer, ByVal posX As Integer, ByVal posY As Integer)
    If MapData(map, posX, posY).Trap Is Nothing Then
     Exit Sub
    End If
    If Not MapData(map, posX, posY).Trap.CanAffectTarget(TargetIndex, TargetType) Then
        Exit Sub
    End If
    Call MapData(map, posX, posY).Trap.trigger(TargetIndex, TargetType)
End Sub
