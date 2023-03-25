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

Public Sub SetTranslationState(ByRef Reference As t_AnyReference, ByVal NewState As Boolean)
    If Not IsValidRef(Reference) Then
        Exit Sub
    End If
    If Reference.RefType = eNpc Then
        NpcList(Reference.ArrayIndex).flags.TranslationActive = NewState
    ElseIf Reference.RefType = eUser Then
        UserList(Reference.ArrayIndex).flags.TranslationActive = NewState
    End If
End Sub
