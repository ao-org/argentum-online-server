Attribute VB_Name = "ModCuentas"
Option Explicit

Public Function PasswordValida(Password As String, PasswordHash As String, Salt As String) As Boolean
        
        On Error GoTo PasswordValida_Err
        

        Dim oSHA256 As CSHA256

100     Set oSHA256 = New CSHA256

102     PasswordValida = (PasswordHash = oSHA256.SHA256(Password & Salt))
    
104     Set oSHA256 = Nothing

        
        Exit Function

PasswordValida_Err:
106     Call TraceError(Err.Number, Err.Description, "ModCuentas.PasswordValida", Erl)

        
End Function

Public Function GetUserGuildIndex(ByVal UserName As String) As Integer
        
        On Error GoTo GetUserGuildIndex_Err

100     If InStrB(UserName, "\") <> 0 Then
102         UserName = Replace(UserName, "\", vbNullString)
        End If

104     If InStrB(UserName, "/") <> 0 Then
106         UserName = Replace(UserName, "/", vbNullString)
        End If

108     If InStrB(UserName, ".") <> 0 Then
110         UserName = Replace(UserName, ".", vbNullString)
        End If

116     GetUserGuildIndex = GetUserGuildIndexDatabase(UserName)

        Exit Function

GetUserGuildIndex_Err:
118     Call TraceError(Err.Number, Err.Description, "ModCuentas.GetUserGuildIndex", Erl)

End Function

Public Function ObtenerCriminal(ByVal Name As String) As Byte

        On Error GoTo ErrorHandler
    
        Dim Criminal As Byte

102     Criminal = GetUserStatusDatabase(Name)

106     If EsRolesMaster(Name) Then
108         Criminal = 3
110     ElseIf EsConsejero(Name) Then
112         Criminal = 4
114     ElseIf EsSemiDios(Name) Then
116         Criminal = 5
118     ElseIf EsDios(Name) Then
120         Criminal = 6
122     ElseIf EsAdmin(Name) Then
124         Criminal = 7
        End If

126     ObtenerCriminal = Criminal

        Exit Function
ErrorHandler:
128     ObtenerCriminal = 1

End Function
