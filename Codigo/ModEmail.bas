Attribute VB_Name = "ModCuentas"
Option Explicit

Public Function ObtenerCodigo(ByVal name As String) As String
        
    ObtenerCodigo = GetCuentaValue(name, "validate_code")

End Function

Public Function ObtenerValidacion(ByVal name As String) As Boolean

    ObtenerValidacion = GetCuentaValue(name, "validated")

End Function

Public Function ObtenerEmail(ByVal name As String) As String

    ObtenerEmail = GetCuentaValue(name, "email")

End Function

Public Function ObtenerMacAdress(ByVal name As String) As String
        
    ObtenerMacAdress = GetCuentaValue(name, "mac_address")

End Function

Public Function ObtenerHDserial(ByVal name As String) As Long
        
    ObtenerHDserial = GetCuentaValue(name, "hd_serial")
        
End Function

Public Function CuentaExiste(ByVal CuentaEmail As String) As Boolean
        
    CuentaExiste = GetCuentaValue(CuentaEmail, "COUNT(*)") > 0
        
End Function

Public Sub SaveNewAccount(ByVal UserIndex As Integer, ByVal CuentaEmail As String, ByVal Password As String)

    Dim Salt As String * 10: Salt = RandomString(10) ' Alfanumerico
    
    Dim oSHA256 As CSHA256
    Set oSHA256 = New CSHA256

    Dim PasswordHash As String * 64: PasswordHash = oSHA256.SHA256(Password & Salt)
    Dim Codigo As String * 6: Codigo = RandomString(6, True) ' Letras mayusculas y numeros

    Call MakeQuery("INSERT INTO account SET email = ?, password = ?, salt = ?, validate_code = ?, date_created = NOW();", True, LCase$(CuentaEmail), PasswordHash, Salt, Codigo)
    
    Set oSHA256 = Nothing

End Sub

Public Function ObtenerCuenta(ByVal name As String) As String

    'Hacemos la query.
    Call MakeQuery("SELECT email FROM `account` INNER JOIN `user` ON user.account_id = account.id WHERE UPPER(user.name) = ?;", False, UCase$(name))
    
    'Verificamos que la query no devuelva un resultado vacio.
    If QueryData Is Nothing Then Exit Function
    
    'Obtenemos el nombre de la cuenta
    ObtenerCuenta = QueryData!name

End Function

Public Function PasswordValida(Password As String, PasswordHash As String, Salt As String) As Boolean

    Dim oSHA256 As CSHA256
    Set oSHA256 = New CSHA256

    PasswordValida = (PasswordHash = oSHA256.SHA256(Password & Salt))
    
    Set oSHA256 = Nothing

End Function

Public Function ObtenerBaneo(ByVal name As String) As Boolean
        
    ObtenerBaneo = CBool(GetCuentaValue(name, "is_banned"))
        
End Function

Public Function ObtenerMotivoBaneo(ByVal name As String) As String
        
    ObtenerMotivoBaneo = GetCuentaValue(name, "ban_reason")
    
End Function

Public Function ObtenerQuienBaneo(ByVal name As String) As String
        
    ObtenerQuienBaneo = GetCuentaValue(name, "banned_by")
        
End Function

Public Function ObtenerCantidadDePersonajes(ByVal name As String) As String
        
    Dim Id As Integer
        Id = GetDBValue("account", "id", "email", LCase$(name))
    
    ObtenerCantidadDePersonajes = GetPersonajesCountByID(Id)
        
End Function

Public Function GetPersonajesCountByID(ByVal AccountId As Long) As Byte

    Call MakeQuery("SELECT COUNT(*) FROM user WHERE deleted = FALSE AND account_id = ?;", False, AccountId)
    
    If QueryData Is Nothing Then Exit Function
    
    GetPersonajesCountByID = QueryData.Fields(0).Value

End Function

Public Function ObtenerCantidadDePersonajesByUserIndex(ByVal UserIndex As Integer) As Byte

    ObtenerCantidadDePersonajesByUserIndex = GetPersonajesCountByID(UserList(UserIndex).AccountId)

End Function

Public Function ObtenerLogeada(ByVal name As String) As Byte
        
    ObtenerLogeada = GetCuentaValue(name, "is_logged")
        
End Function

Sub BorrarCuenta(ByVal CuentaName As String)
        
    On Error GoTo ErrorHandler

    Dim Id As Integer
        Id = GetDBValue("account", "id", "email", LCase$(CuentaName))

    Call MakeQuery("UPDATE account SET email = CONCAT('DELETED_', email), deleted = TRUE WHERE email = ?;", True, LCase$(CuentaName))

    Call MakeQuery("UPDATE user SET name = CONCAT('DELETED_', name), deleted = TRUE WHERE account_id = ?;", True, Id)

    Exit Sub
    
ErrorHandler:
    Call LogDatabaseError("Error en BorrarCuentaDatabase borrando user de la Mysql Database: " & CuentaName & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function GetUserGuildIndex(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Juan Andres Dalmasso
    'Last Modification: 18/09/2018
    '18/09/2018 CHOTS: Checks database too
    '***************************************************
    If InStrB(UserName, "\") <> 0 Then
        UserName = Replace(UserName, "\", vbNullString)
    End If

    If InStrB(UserName, "/") <> 0 Then
        UserName = Replace(UserName, "/", vbNullString)
    End If

    If InStrB(UserName, ".") <> 0 Then
        UserName = Replace(UserName, ".", vbNullString)
    End If

    GetUserGuildIndex = SanitizeNullValue(GetUserValue(UserName, "guild_index"), 0)

End Function

Public Function ObtenerCriminal(ByVal name As String) As Byte

    On Error GoTo ErrorHandler
    
    Dim Criminal As Byte
        Criminal = GetUserValue(name, "status")

    If EsRolesMaster(name) Then
        Criminal = 3
    ElseIf EsConsejero(name) Then
        Criminal = 4
    ElseIf EsSemiDios(name) Then
        Criminal = 5
    ElseIf EsDios(name) Then
        Criminal = 6
    ElseIf EsAdmin(name) Then
        Criminal = 7

    End If

    ObtenerCriminal = Criminal

    Exit Function
ErrorHandler:
    ObtenerCriminal = 1

End Function

Public Sub SaveBanDatabase(UserName As String, Reason As String, BannedBy As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call MakeQuery("UPDATE user SET is_banned = TRUE WHERE UPPER(name) = ?;", True, UCase$(UserName))

    query = "INSERT INTO punishment SET "
    query = query & "user_id = (SELECT id from user WHERE UPPER(name) = ?), "
    query = query & "number = number + 1, "
    query = query & "reason = ?;"

    Call MakeQuery(query, True, UCase$(UserName), BannedBy & ": " & Reason & " " & Date & " " & Time)

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in SaveBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.Description)

End Sub
