Attribute VB_Name = "UserNames"
Option Explicit

Public Const MAX_ALIAS_LEN As Integer = 30

Public Function GetUserDisplayName(ByVal UserIndex As Integer) As String
    On Error GoTo GetUserDisplayName_Err
    If UserIndex < LBound(UserList) Or UserIndex > UBound(UserList) Then Exit Function
    If Not UserList(UserIndex).flags.UserLogged Then Exit Function
    If Not IsFeatureEnabled("EnablePatreonAlias") Then
        GetUserDisplayName = UserList(UserIndex).name
        Exit Function
    End If
    If IsPatreon(UserIndex) Then
        Dim aliasValue As String
        aliasValue = Trim$(UserList(UserIndex).Alias)
        If LenB(aliasValue) <> 0 Then
            GetUserDisplayName = aliasValue
            Exit Function
        End If
    End If
    GetUserDisplayName = UserList(UserIndex).name
    Exit Function
GetUserDisplayName_Err:
    Call TraceError(Err.Number, Err.Description, "UserNames.GetUserDisplayName", Erl)
End Function

Public Function GetUserRealName(ByVal UserIndex As Integer) As String
    On Error GoTo GetUserRealName_Err
    If UserIndex < LBound(UserList) Or UserIndex > UBound(UserList) Then Exit Function
    If Not UserList(UserIndex).flags.UserLogged Then Exit Function
    GetUserRealName = UserList(UserIndex).name
    Exit Function
GetUserRealName_Err:
    Call TraceError(Err.Number, Err.Description, "UserNames.GetUserRealName", Erl)
End Function


Public Function GetUserDisplayNameOrReal(ByVal UserIndex As Integer) As String
    Dim displayName As String
    displayName = GetUserDisplayName(UserIndex)
    If LenB(displayName) = 0 Then
        displayName = GetUserRealName(UserIndex)
    End If
    GetUserDisplayNameOrReal = displayName
End Function

Public Function GetUserGMName(ByVal UserIndex As Integer) As String
    On Error GoTo GetUserGMName_Err
    Dim displayName As String
    Dim realName As String
    displayName = GetUserDisplayName(UserIndex)
    realName = GetUserRealName(UserIndex)
    If LenB(displayName) <> 0 And LenB(realName) <> 0 And displayName <> realName Then
        GetUserGMName = displayName & " (" & realName & ")"
    Else
        GetUserGMName = realName
    End If
    Exit Function
GetUserGMName_Err:
    Call TraceError(Err.Number, Err.Description, "UserNames.GetUserGMName", Erl)
End Function

Public Function ValidateAlias(ByVal aliasValue As String, ByRef errorMessage As String) As Boolean
    Dim trimmed As String
    trimmed = Trim$(aliasValue)
    If LenB(trimmed) = 0 Then
        ValidateAlias = True
        Exit Function
    End If
    If Len(trimmed) > MAX_ALIAS_LEN Then
        errorMessage = "Alias demasiado largo."
        Exit Function
    End If
    Dim i As Long
    For i = 1 To Len(trimmed)
        Dim c As String
        Dim code As Long
        c = Mid$(trimmed, i, 1)
        code = AscW(c)
        If code < 32 Or code = 127 Then
            errorMessage = "Alias contiene caracteres invalidos."
            Exit Function
        End If
        If c = Chr$(255) Then
            errorMessage = "Alias contiene caracteres invalidos."
            Exit Function
        End If
    Next i
    ValidateAlias = True
End Function
