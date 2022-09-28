Attribute VB_Name = "Database"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.ao20.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'

'Argentum Online Libre
'Database connection module
'Obtained from GS-Zone
'Adapted and modified by Juan Andres Dalmasso (CHOTS)
'September 2018
'Rewrited for Argentum20 by Alexis Caraballo (WyroX)
'October 2020



Public Sub Database_Connect_Async()
        On Error GoTo Database_Connect_AsyncErr
        
        Dim ConnectionID As String

        If Len(Database_Source) <> 0 Then
104         ConnectionID = "DATA SOURCE=" & Database_Source & ";"
        Else
106         ConnectionID = "DRIVER={SQLite3 ODBC Driver};" & "DATABASE=" & App.Path & "/Database.db"
        End If
                
        Dim i As Byte
        
        For i = 1 To MAX_ASYNC
            Set Connection_async(i) = New ADODB.Connection
110         Connection_async(i).CursorLocation = adUseClient
            Connection_async(i).ConnectionString = ConnectionID
112         Call Connection_async(i).Open(, , , adAsyncConnect)
        Next i

        Current_async = 1
        
113     Set Builder = New cStringBuilder

        Exit Sub
    
Database_Connect_AsyncErr:
116     Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description & " - Database_Connect_Async")
End Sub
Public Sub Database_Connect()
        On Error GoTo Database_Connect_Err
        
        Dim ConnectionID As String

        If Len(Database_Source) <> 0 Then
104         ConnectionID = "DATA SOURCE=" & Database_Source & ";"
        Else
106         ConnectionID = "DRIVER={SQLite3 ODBC Driver};" & "DATABASE=" & App.Path & "/Database.db"
        End If
                
        Set Connection = New ADODB.Connection
110     Connection.CursorLocation = adUseClient
        Connection.ConnectionString = ConnectionID

113     Set Builder = New cStringBuilder
        
112     Call Connection.Open

        Exit Sub
    
Database_Connect_Err:
116     Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description & " - Database_Connect")
End Sub

Public Sub Database_Close()
        On Error GoTo Database_Close_Err

        If Connection.State <> adStateClosed Then
            Call Connection.Close
        End If

        Set Connection = Nothing
        
        Exit Sub
     
Database_Close_Err:
        Call LogDatabaseError("Unable to close Mysql Database: " & Err.Number & " - " & Err.Description)
End Sub

Public Function Query(ByVal Text As String, ParamArray Arguments() As Variant) As ADODB.Recordset
    Dim Command  As New ADODB.Command
    Dim Argument As Variant
    
    Command.ActiveConnection = Connection
    Command.CommandText = Text
    Command.CommandType = adCmdText
    Command.Prepared = True
    
    For Each Argument In Arguments
        If (IsArray(Argument)) Then
            Dim Inner As Variant
            
            For Each Inner In Argument
                Command.Parameters.Append CreateParameter(Inner, adParamInput)
            Next Inner
        Else
            Command.Parameters.Append CreateParameter(Argument, adParamInput)
        End If
    Next Argument

    On Error GoTo Query_Err

    ' Statistics
    If frmMain.chkLogDbPerfomance.value = 1 Then
        Call GetElapsedTime
    End If
    
    Set Query = Command.Execute()
    
    ' Statistics
    If frmMain.chkLogDbPerfomance.value = 1 Then
        Call LogPerformance("Query: " & Text & vbNewLine & " - Tiempo transcurrido: " & Round(GetElapsedTime(), 1) & " ms" & vbNewLine)
    End If

    Exit Function
    
Query_Err:
    DBError = Err.Description
    Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description & " - " & vbCrLf & Text)
End Function

Public Function Execute(ByVal Text As String, ParamArray Arguments() As Variant) As Boolean
    Dim Command  As New ADODB.Command
    Dim Argument As Variant
    
    Command.ActiveConnection = Connection_async(Current_async)
    Command.CommandText = Text
    Command.CommandType = adCmdText
    Command.Prepared = True

    For Each Argument In Arguments
        If (IsArray(Argument)) Then
            Dim Inner As Variant
            
            For Each Inner In Argument
                Command.Parameters.Append CreateParameter(Inner, adParamInput)
            Next Inner
        Else
            Command.Parameters.Append CreateParameter(Argument, adParamInput)
        End If
    Next Argument
    
On Error GoTo Execute_Err
    
    ' Statistics
    If frmMain.chkLogDbPerfomance.value = 1 Then
        Call GetElapsedTime
    End If
    
    Call Command.Execute(, , adAsyncExecute)  ' @TODO: We want some operation to be async
    
    Current_async = Current_async + 1
    
    If Current_async = MAX_ASYNC Then
        Current_async = 1
    End If
    
    ' Statistics
    If frmMain.chkLogDbPerfomance.value = 1 Then
        Call LogPerformance("Execute: " & Text & vbNewLine & " - Tiempo transcurrido: " & Round(GetElapsedTime(), 1) & " ms" & vbNewLine)
    End If
    
    Execute = (Err.Number = 0)
    Exit Function
        
Execute_Err:
    
    If (Err.Number <> 0) Then
        Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description & " - " & vbCrLf & Text)
    End If
    
End Function

Public Function Invoke(ByVal Procedure As String, ParamArray Arguments() As Variant) As ADODB.Recordset
    Dim Command  As New ADODB.Command
    Dim Argument As Variant
    Dim Affected As Long
    
    Command.ActiveConnection = Connection
    Command.CommandText = Procedure
    Command.CommandType = adCmdStoredProc
    Command.Prepared = True

    For Each Argument In Arguments
        If (IsArray(Argument)) Then
            Dim Inner As Variant
            
            For Each Inner In Argument
                Command.Parameters.Append CreateParameter(Inner, adParamInput)
            Next Inner
        Else
            Command.Parameters.Append CreateParameter(Argument, adParamInput)
        End If
    Next Argument
    
On Error GoTo Execute_Err
    
    ' Statistics
    If frmMain.chkLogDbPerfomance.value = 1 Then
        Call GetElapsedTime
    End If
    
    Set Invoke = Command.Execute()
    
    If (Not Invoke Is Nothing And Invoke.EOF) Then
        Set Invoke = Nothing
    End If
    
    ' Statistics
    If frmMain.chkLogDbPerfomance.value = 1 Then
        Call LogPerformance("Invoke: " & Procedure & vbNewLine & " - Tiempo transcurrido: " & Round(GetElapsedTime(), 1) & " ms" & vbNewLine)
    End If
    
Execute_Err:
    If (Err.Number <> 0) Then
        Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.Description)
    End If
End Function

Private Function CreateParameter(ByVal value As Variant, ByVal Direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter
    Set CreateParameter = New ADODB.Parameter
    
    CreateParameter.Direction = Direction
    
    Select Case VarType(value)
        Case VbVarType.vbString
            CreateParameter.Type = adBSTR
            CreateParameter.size = Len(value)
            CreateParameter.value = CStr(value)
        Case VbVarType.vbDecimal
            CreateParameter.Type = adInteger
            CreateParameter.value = CLng(value)
        Case VbVarType.vbByte:
            CreateParameter.Type = adTinyInt
            CreateParameter.value = CByte(value)
        Case VbVarType.vbInteger
            CreateParameter.Type = adSmallInt
            CreateParameter.value = CInt(value)
        Case VbVarType.vbLong
            CreateParameter.Type = adInteger
            CreateParameter.value = CLng(value)
        Case VbVarType.vbBoolean
            CreateParameter.Type = adBoolean
            CreateParameter.value = CBool(value)
        Case VbVarType.vbSingle
            CreateParameter.Type = adSingle
            CreateParameter.value = CSng(value)
        Case VbVarType.vbDouble
            CreateParameter.Type = adDouble
            CreateParameter.value = CDbl(value)
    End Select
End Function

Public Function GetDBValue(Tabla As String, ColumnaGet As String, ColumnaTest As String, ValueTest As Variant) As Variant
        On Error GoTo ErrorHandler
    
100     Dim RS As ADODB.Recordset
        Set RS = Query("SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE LOWER(" & ColumnaTest & ") = ?;", ValueTest)

        'Revisamos si recibio un resultado
102     If RS Is Nothing Then Exit Function
        If RS.BOF Or RS.EOF Then Exit Function
        
        'Obtenemos la variable
104     GetDBValue = RS.Fields(ColumnaGet).value

        Exit Function
    
ErrorHandler:
106     Call LogDatabaseError("Error en GetDBValue: SELECT " & ColumnaGet & " FROM " & Tabla & " WHERE " & ColumnaTest & " = '" & ValueTest & "';" & ". " & Err.Number & " - " & Err.Description)
End Function

Public Function GetUserValue(CharName As String, Columna As String) As Variant
        On Error GoTo GetUserValue_Err
        
100     GetUserValue = GetDBValue("user", Columna, "name", CharName)
        
        Exit Function

GetUserValue_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetUserValue", Erl)
End Function

Public Sub SetDBValue(Tabla As String, ColumnaSet As String, ByVal ValueSet As Variant, ColumnaTest As String, ByVal ValueTest As Variant)
        On Error GoTo ErrorHandler

        Call Execute("UPDATE " & Tabla & " SET " & ColumnaSet & " = ? WHERE " & ColumnaTest & " = ?;", ValueSet, ValueTest)

        Exit Sub
    
ErrorHandler:
102     Call LogDatabaseError("Error en SetDBValue: UPDATE " & Tabla & " SET " & ColumnaSet & " = " & ValueSet & " WHERE " & ColumnaTest & " = " & ValueTest & ";" & ". " & Err.Number & " - " & Err.Description)
End Sub

Private Sub SetUserValue(CharName As String, Columna As String, value As Variant)
        On Error GoTo SetUserValue_Err
        
100     Call SetDBValue("user", Columna, value, "UPPER(name)", UCase(CharName))

        Exit Sub

SetUserValue_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SetUserValue", Erl)
End Sub

Private Sub SetUserValueByID(ByVal ID As Long, Columna As String, value As Variant)
        On Error GoTo SetUserValueByID_Err
        
100     Call SetDBValue("user", Columna, value, "id", ID)

        Exit Sub

SetUserValueByID_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SetUserValueByID", Erl)
End Sub

Public Function BANCheckDatabase(Name As String) As Boolean
        
        On Error GoTo BANCheckDatabase_Err
        
100     BANCheckDatabase = CBool(GetUserValue(LCase$(Name), "is_banned"))
  
        Exit Function

BANCheckDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.BANCheckDatabase", Erl)
End Function

Public Function GetUserStatusDatabase(Name As String) As Integer
        
        On Error GoTo GetUserStatusDatabase_Err
        
100     GetUserStatusDatabase = GetUserValue(LCase$(Name), "status")

        
        Exit Function

GetUserStatusDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.GetUserStatusDatabase", Erl)

        
End Function

Public Function GetAccountIDDatabase(Name As String) As Long
        
        On Error GoTo GetAccountIDDatabase_Err
        

        Dim Temp As Variant

100     Temp = GetUserValue(LCase$(Name), "account_id")
    
102     If VBA.IsEmpty(Temp) Then
104         GetAccountIDDatabase = -1
        Else
106         GetAccountIDDatabase = val(Temp)

        End If

        
        Exit Function

GetAccountIDDatabase_Err:
108     Call TraceError(Err.Number, Err.Description, "modDatabase.GetAccountIDDatabase", Erl)

        
End Function

Public Function GetPersonajesCountByIDDatabase(ByVal AccountID As Long) As Byte

        On Error GoTo ErrorHandler
    
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT COUNT(*) FROM user WHERE account_id = ?", AccountID)
    
102     If RS Is Nothing Then Exit Function
    
104     GetPersonajesCountByIDDatabase = RS.Fields(0).value
    
        Exit Function
    
ErrorHandler:
106     Call LogDatabaseError("Error in GetPersonajesCountByIDDatabase. AccountID: " & AccountID & ". " & Err.Number & " - " & Err.Description)
    
End Function

Public Function GetPersonajesCuentaDatabase(ByVal AccountID As Long, Personaje() As t_PersonajeCuenta) As Byte
        
        On Error GoTo GetPersonajesCuentaDatabase_Err
        
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT name, head_id, class_id, body_id, pos_map, pos_x, pos_y, level, status, helmet_id, shield_id, weapon_id, guild_index, is_dead, is_sailing FROM user WHERE account_id = ?;", AccountID)

102     If RS Is Nothing Then Exit Function
    
104     GetPersonajesCuentaDatabase = RS.RecordCount

        Dim i As Integer
        If GetPersonajesCuentaDatabase = 0 Then Exit Function
108     For i = 1 To GetPersonajesCuentaDatabase
110         Personaje(i).nombre = RS!Name
112         Personaje(i).Cabeza = RS!head_id
114         Personaje(i).clase = RS!class_id
116         Personaje(i).cuerpo = RS!body_id
118         Personaje(i).Mapa = RS!pos_map
120         Personaje(i).posX = RS!pos_X
122         Personaje(i).posY = RS!pos_Y
124         Personaje(i).nivel = RS!level
126         Personaje(i).Status = RS!Status
128         Personaje(i).Casco = RS!helmet_id
130         Personaje(i).Escudo = RS!shield_id
132         Personaje(i).Arma = RS!weapon_id
134         Personaje(i).ClanIndex = RS!Guild_Index
        
136         If EsRolesMaster(Personaje(i).nombre) Then
138             Personaje(i).Status = 3
140         ElseIf EsConsejero(Personaje(i).nombre) Then
142             Personaje(i).Status = 4
144         ElseIf EsSemiDios(Personaje(i).nombre) Then
146             Personaje(i).Status = 5
148         ElseIf EsDios(Personaje(i).nombre) Then
150             Personaje(i).Status = 6
152         ElseIf EsAdmin(Personaje(i).nombre) Then
154             Personaje(i).Status = 7

            End If

156         If val(RS!is_dead) = 1 Or val(RS!is_sailing) = 1 Then
158             Personaje(i).Cabeza = 0
            End If
        
160         RS.MoveNext
        Next

        Exit Function

GetPersonajesCuentaDatabase_Err:
162     Call TraceError(Err.Number, Err.Description, "modDatabase.GetPersonajesCuentaDatabase", Erl)

        
End Function


Public Sub SaveVotoDatabase(ByVal ID As Long, ByVal Encuestas As Integer)
        
        On Error GoTo SaveVotoDatabase_Err
        
100     Call SetUserValueByID(ID, "votes_amount", Encuestas)
        
        Exit Sub

SaveVotoDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SaveVotoDatabase", Erl)

        
End Sub

Public Sub SaveUserBodyDatabase(username As String, ByVal body As Integer)
        
        On Error GoTo SaveUserBodyDatabase_Err
        
100     Call SetUserValue(username, "body_id", body)

        
        Exit Sub

SaveUserBodyDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SaveUserBodyDatabase", Erl)

        
End Sub

Public Sub SaveUserHeadDatabase(username As String, ByVal head As Integer)
        
        On Error GoTo SaveUserHeadDatabase_Err
        
100     Call SetUserValue(username, "head_id", head)

        
        Exit Sub

SaveUserHeadDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SaveUserHeadDatabase", Erl)

        
End Sub

Public Sub SaveUserSkillDatabase(username As String, ByVal Skill As Integer, ByVal value As Integer)
        
        On Error GoTo SaveUserSkillDatabase_Err
        
        Call Execute("UPDATE skillpoints SET value = ? WHERE number = ? AND user_id = (SELECT id FROM user WHERE UPPER(name) = ?)", value, Skill, UCase$(username))
        
        Exit Sub

SaveUserSkillDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SaveUserSkillDatabase", Erl)

        
End Sub

Public Sub SaveUserSkillsLibres(username As String, ByVal SkillsLibres As Integer)
        
        On Error GoTo SaveUserHeadDatabase_Err
        
100     Call SetUserValue(username, "free_skillpoints", SkillsLibres)
        
        Exit Sub

SaveUserHeadDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SaveUserHeadDatabase", Erl)

        
End Sub


Public Sub SaveBanDatabase(username As String, Reason As String, BannedBy As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler
        
        Call Execute("UPDATE user SET is_banned = TRUE, banned_by = ?, ban_reason = ? WHERE UPPER(name) = ?;", BannedBy, Reason, UCase$(username))
        
102     Call SavePenaDatabase(username, "Baneado por: " & BannedBy & " debido a " & Reason)

        Exit Sub

ErrorHandler:
104     Call LogDatabaseError("Error in SaveBanDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveWarnDatabase(username As String, Reason As String, WarnedBy As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler
        
        Call Execute("UPDATE user SET warnings = warnings + 1 WHERE UPPER(name) = ?;", UCase$(username))
        
102     Call SavePenaDatabase(username, "Advertencia de: " & WarnedBy & " debido a " & Reason)
    
    Exit Sub

ErrorHandler:
104     Call LogDatabaseError("Error in SaveWarnDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SavePenaDatabase(username As String, Reason As String)

        On Error GoTo ErrorHandler

        Dim Query As String
100     Query = "INSERT INTO punishment(user_id, NUMBER, reason)"
102     Query = Query & " SELECT u.id, COUNT(p.number) + 1, ? FROM user u LEFT JOIN punishment p ON p.user_id = u.id WHERE UPPER(u.name) = ?;"
        
        Call Execute(Query, Reason, UCase$(username))

        Exit Sub

ErrorHandler:
106     Call LogDatabaseError("Error in SavePenaDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SilenciarUserDatabase(username As String, ByVal Tiempo As Integer)
    
        On Error GoTo ErrorHandler
        
        Call Execute("UPDATE user SET is_silenced = 1, silence_minutes_left = ?, silence_elapsed_seconds = 0 WHERE UPPER(name) = ?;", Tiempo, UCase$(username))
        
        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in SilenciarUserDatabase: " & username & ". " & Err.Number & " - " & Err.Description)
    
End Sub

Public Sub DesilenciarUserDatabase(username As String)

        On Error GoTo ErrorHandler

100     Call SetUserValue(username, "is_silenced", 0)

        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in DesilenciarUserDatabase: " & username & ". " & Err.Number & " - " & Err.Description)
    
End Sub

Public Sub UnBanDatabase(username As String)

        On Error GoTo ErrorHandler
        
        Call Execute("UPDATE user SET is_banned = FALSE, banned_by = '', ban_reason = '' WHERE UPPER(name) = ?;", UCase$(username))
        
        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in UnBanDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveBanCuentaDatabase(ByVal AccountID As Long, Reason As String, BannedBy As String)

        On Error GoTo ErrorHandler
        
        Call Execute("UPDATE account SET is_banned = TRUE, banned_by = ?, ban_reason = ? WHERE id = ?;", BannedBy, Reason, AccountID)

        Exit Sub

ErrorHandler:
102     Call LogDatabaseError("Error in SaveBanCuentaDatabase: AccountId=" & AccountID & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub EcharConsejoDatabase(username As String, ByVal Status As Integer)
        
        On Error GoTo EcharConsejoDatabase_Err
        
        Call Execute("UPDATE user SET status = ? WHERE UPPER(name) = ?;", Status, UCase$(username))
        
        Exit Sub

EcharConsejoDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.EcharConsejoDatabase", Erl)

        
End Sub

Public Sub EcharLegionDatabase(username As String)
        
        On Error GoTo EcharLegionDatabase_Err
        
        Call Execute("UPDATE user SET status = ? WHERE UPPER(name) = ?;", e_Facciones.Criminal, UCase$(username))
        
        Exit Sub

EcharLegionDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.EcharLegionDatabase", Erl)

        
End Sub

Public Sub EcharArmadaDatabase(username As String)
        
        On Error GoTo EcharArmadaDatabase_Err
        
        Call Execute("UPDATE user SET status = ? WHERE UPPER(name) = ?;", e_Facciones.Ciudadano, UCase$(username))

        Exit Sub

EcharArmadaDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.EcharArmadaDatabase", Erl)

        
End Sub

Public Sub CambiarPenaDatabase(username As String, ByVal Numero As Integer, Pena As String)
        
        On Error GoTo CambiarPenaDatabase_Err
        
        Call Execute("UPDATE punishment SET reason = ? WHERE number = ? AND user_id = (SELECT id from user WHERE UPPER(name) = ?);", Pena, Numero, UCase$(username))
        
        Exit Sub

CambiarPenaDatabase_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.CambiarPenaDatabase", Erl)

        
End Sub

Public Function GetUserAmountOfPunishmentsDatabase(ByVal username As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler
        
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT COUNT(*) as punishments FROM `punishment` INNER JOIN `user` ON punishment.user_id = user.id WHERE UPPER(user.name) = ?;", UCase$(username))

102     If RS Is Nothing Then Exit Function

104     GetUserAmountOfPunishmentsDatabase = RS!punishments

        Exit Function
ErrorHandler:
106     Call LogDatabaseError("Error in GetUserAmountOfPunishmentsDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Function

Public Sub SendUserPunishmentsDatabase(ByVal userIndex As Integer, ByVal username As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 10/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT user_id, number, reason FROM `punishment` INNER JOIN `user` ON punishment.user_id = user.id WHERE UPPER(user.name) = ?;", UCase$(username))
    
102     If RS Is Nothing Then Exit Sub

104     If Not RS.RecordCount = 0 Then

108         While Not RS.EOF
110             Call WriteConsoleMsg(userIndex, RS!Number & " - " & RS!Reason, e_FontTypeNames.FONTTYPE_INFO)
            
112             RS.MoveNext
            Wend

        End If

        Exit Sub
ErrorHandler:
114     Call LogDatabaseError("Error in SendUserPunishmentsDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function GetUserGuildIndexDatabase(username As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 09/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildIndexDatabase = SanitizeNullValue(GetUserValue(LCase$(username), "guild_index"), 0)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildIndexDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildMemberDatabase(username As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildMemberDatabase = SanitizeNullValue(GetUserValue(LCase$(username), "guild_member_history"), vbNullString)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildMemberDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildAspirantDatabase(username As String) As Integer

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildAspirantDatabase = SanitizeNullValue(GetUserValue(LCase$(username), "guild_aspirant_index"), 0)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildAspirantDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetUserGuildPedidosDatabase(username As String) As String

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     GetUserGuildPedidosDatabase = SanitizeNullValue(GetUserValue(LCase$(username), "guild_requests_history"), vbNullString)

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in GetUserGuildPedidosDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Function

Public Sub SaveUserGuildRejectionReasonDatabase(username As String, Reason As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(username, "guild_rejected_because", Reason)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildRejectionReasonDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildIndexDatabase(ByVal username As String, ByVal GuildIndex As Integer)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(username, "guild_index", GuildIndex)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildIndexDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildAspirantDatabase(ByVal username As String, ByVal AspirantIndex As Integer)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(username, "guild_aspirant_index", AspirantIndex)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildAspirantDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildMemberDatabase(ByVal username As String, ByVal guilds As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(username, "guild_member_history", guilds)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildMemberDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SaveUserGuildPedidosDatabase(ByVal username As String, ByVal Pedidos As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

100     Call SetUserValue(username, "guild_requests_history", Pedidos)

        Exit Sub
ErrorHandler:
102     Call LogDatabaseError("Error in SaveUserGuildPedidosDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Sub SendCharacterInfoDatabase(ByVal userIndex As Integer, ByVal username As String)

        '***************************************************
        'Author: Juan Andres Dalmasso (CHOTS)
        'Last Modification: 11/10/2018
        '***************************************************
        On Error GoTo ErrorHandler

        Dim gName       As String

        Dim Miembro     As String

        Dim GuildActual As Integer

        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT race_id, class_id, genre_id, level, gold, bank_gold, guild_requests_history, guild_index, guild_member_history, pertenece_real, pertenece_caos, ciudadanos_matados, criminales_matados FROM user WHERE UPPER(name) = ?;", UCase$(username))

102     If RS Is Nothing Then
104         Call WriteConsoleMsg(userIndex, "Pj Inexistente", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        ' Get the character's current guild
106     GuildActual = SanitizeNullValue(RS!Guild_Index, 0)

108     If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
110         gName = "<" & GuildName(GuildActual) & ">"
        Else
112         gName = "Ninguno"

        End If

        'Get previous guilds
114     Miembro = SanitizeNullValue(RS!guild_member_history, vbNullString)

116     If Len(Miembro) > 400 Then
118         Miembro = ".." & Right$(Miembro, 400)

        End If

120     Call WriteCharacterInfo(userIndex, username, RS!race_id, RS!class_id, RS!genre_id, RS!level, RS!gold, RS!bank_gold, SanitizeNullValue(RS!guild_requests_history, vbNullString), gName, Miembro, RS!pertenece_real, RS!pertenece_caos, RS!ciudadanos_matados, RS!criminales_matados)

        Exit Sub
ErrorHandler:
122     Call LogDatabaseError("Error in SendCharacterInfoDatabase: " & username & ". " & Err.Number & " - " & Err.Description)

End Sub

Public Function EnterAccountDatabase(ByVal userIndex As Integer, ByVal CuentaEmail As String) As Boolean

        On Error GoTo ErrorHandler
    
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT id from account WHERE email = ?", UCase$(CuentaEmail))
    
102     If Connection.State = adStateClosed Then
104         Call WriteShowMessageBox(userIndex, "Ha ocurrido un error interno en el servidor. ¡Estamos tratando de resolverlo!")
            Exit Function
        End If
    
122     UserList(userIndex).AccountID = RS!ID
124     UserList(userIndex).Cuenta = CuentaEmail
        UserList(userIndex).Email = CuentaEmail
    
128     EnterAccountDatabase = True
    
        Exit Function

ErrorHandler:
130     Call LogDatabaseError("Error in EnterAccountDatabase. UserCuenta: " & CuentaEmail & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function PersonajePerteneceID(ByVal username As String, ByVal AccountID As Long) As Boolean
    
        Dim RS As ADODB.Recordset
100     Set RS = Query("SELECT id FROM user WHERE name = ? AND account_id = ?;", username, AccountID)
    
102     If RS Is Nothing Then
104         PersonajePerteneceID = False
            Exit Function
        End If
    
106     PersonajePerteneceID = True
    
End Function

Public Function SetPositionDatabase(username As String, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
        On Error GoTo ErrorHandler

102     SetPositionDatabase = Execute("UPDATE user SET pos_map = ?, pos_x = ?, pos_y = ? WHERE UPPER(name) = ?;", map, x, y, UCase$(username))

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in SetPositionDatabase. UserName: " & username & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function GetMapDatabase(username As String) As Integer
        On Error GoTo ErrorHandler

100     GetMapDatabase = val(GetUserValue(LCase$(username), "pos_map"))

        Exit Function

ErrorHandler:
102     Call LogDatabaseError("Error in SetPositionDatabase. UserName: " & username & ". " & Err.Number & " - " & Err.Description)

End Function

Public Function AddOroBancoDatabase(username As String, ByVal OroGanado As Long) As Boolean
        On Error GoTo ErrorHandler

102     AddOroBancoDatabase = Execute("UPDATE user SET bank_gold = bank_gold + ? WHERE UPPER(name) = ?;", OroGanado, UCase$(username))

        Exit Function

ErrorHandler:
104     Call LogDatabaseError("Error in AddOroBancoDatabase. UserName: " & username & ". " & Err.Number & " - " & Err.Description)

End Function



Public Function SanitizeNullValue(ByVal value As Variant, ByVal defaultValue As Variant) As Variant
        
        On Error GoTo SanitizeNullValue_Err
        
100     SanitizeNullValue = IIf(IsNull(value), defaultValue, value)
        
        Exit Function

SanitizeNullValue_Err:
102     Call TraceError(Err.Number, Err.Description, "modDatabase.SanitizeNullValue", Erl)

        
End Function

Public Sub SetMessageInfoDatabase(ByVal Name As String, ByVal Message As String)
    Call Execute("update user set message_info = concat(message_info, ?) where upper(name) = ?;", Message, UCase$(Name))
End Sub

Public Sub ChangeNameDatabase(ByVal CurName As String, ByVal NewName As String)
    Call SetUserValue(CurName, "name", NewName)
End Sub

Public Sub ResetLastLogoutAndIsLogged()
    Call Execute("Update user set last_logout = 0, is_logged = 0")
End Sub
