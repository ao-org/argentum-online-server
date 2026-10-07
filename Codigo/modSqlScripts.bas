Attribute VB_Name = "modSqlScripts"
' Copyright (C) 2026 Noland Studios LTD
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit

' ODBC accepts one statement per Execute. Parse the entire script before executing
' any statement, preserving quoted semicolons and removing SQL comments safely.
Public Function SplitMigrationSql(ByVal script As String) As Collection
    Dim statements As Collection, statement As String
    Dim i As Long, state As Integer, ch As String, nextCh As String, closer As String
    Set statements = New Collection
    i = 1
    Do While i <= Len(script)
        ch = Mid$(script, i, 1)
        nextCh = Mid$(script, i + 1, 1)
        Select Case state
            Case 0
                If ch = "-" And nextCh = "-" Then
                    state = 5: statement = statement & " ": i = i + 1
                ElseIf ch = "/" And nextCh = "*" Then
                    state = 6: statement = statement & " ": i = i + 1
                ElseIf ch = ";" Then
                    Call AddSqlMigrationStatement(statements, statement)
                    statement = vbNullString
                Else
                    statement = statement & ch
                    Select Case ch
                        Case "'": state = 1
                        Case Chr$(34): state = 2
                        Case "`": state = 3
                        Case "[": state = 4
                    End Select
                End If
            Case 1, 2, 3, 4
                statement = statement & ch
                Select Case state
                    Case 1: closer = "'"
                    Case 2: closer = Chr$(34)
                    Case 3: closer = "`"
                    Case 4: closer = "]"
                End Select
                If ch = closer Then
                    If nextCh = closer Then
                        statement = statement & nextCh: i = i + 1
                    Else
                        state = 0
                    End If
                End If
            Case 5
                If ch = vbCr Or ch = vbLf Then
                    state = 0: statement = statement & ch
                End If
            Case 6
                If ch = "*" And nextCh = "/" Then
                    state = 0: i = i + 1
                End If
        End Select
        i = i + 1
    Loop
    If state <> 0 And state <> 5 Then Call Err.Raise(5, "SplitMigrationSql", "Unterminated quote or block comment in database migration")
    Call AddSqlMigrationStatement(statements, statement)
    Set SplitMigrationSql = statements
End Function

Private Sub AddSqlMigrationStatement(ByVal statements As Collection, ByVal statement As String)
    Dim normalized As String
    statement = Trim$(statement)
    normalized = UCase$(Replace(Replace(Replace(statement, vbCr, " "), vbLf, " "), vbTab, " "))
    Do While InStr(normalized, "  ") > 0
        normalized = Replace(normalized, "  ", " ")
    Loop
    normalized = Trim$(normalized)
    If Len(normalized) = 0 Then Exit Sub
    ' No existing ScriptsDB migration defines a trigger. Reject compound trigger
    ' bodies explicitly rather than submitting a partial BEGIN/END definition.
    If Left$(normalized, Len("CREATE TRIGGER ")) = "CREATE TRIGGER " Or Left$(normalized, Len("CREATE TEMP TRIGGER ")) = "CREATE TEMP TRIGGER " Or Left$(normalized, Len("CREATE TEMPORARY TRIGGER ")) = "CREATE TEMPORARY TRIGGER " Then
        Call Err.Raise(5, "SplitMigrationSql", "CREATE TRIGGER bodies require explicit migration-runner support")
    End If
    Call statements.Add(statement)
End Sub
