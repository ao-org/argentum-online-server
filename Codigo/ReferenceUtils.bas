Attribute VB_Name = "ModReferenceUtils"
Option Explicit

Public Function GetPosition(ByRef Reference As t_AnyReference) As t_WorldPos
    If Not IsValidRef(Reference) Then
        Exit Function
    End If
    If Reference.RefType = eNpc Then
        GetPosition = NpcList(Reference.ArrayIndex).pos
    ElseIf Reference.RefType = eUser Then
        GetPosition = UserList(Reference.ArrayIndex).pos
    End If
End Function
