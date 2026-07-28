Attribute VB_Name = "modStringBuilderBenchmark"
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const BENCHMARK_OUTPUT_FILE As String = "StringBuilderBenchmark.txt"
Private Const CASE_SMALL_SQL_CSTRING As Long = 1
Private Const CASE_SMALL_SQL_FAST As Long = 2
Private Const CASE_SMALL_SQL_PLAIN As Long = 3
Private Const CASE_CHARACTER_SAVE_CSTRING As Long = 4
Private Const CASE_CHARACTER_SAVE_FAST As Long = 5
Private Const CASE_CHARACTER_SAVE_PLAIN As Long = 6
Private Const CASE_REUSED_CLEAR_CSTRING As Long = 7
Private Const CASE_REUSED_CLEAR_FAST As Long = 8
Private Const CASE_REUSED_CLEAR_PLAIN As Long = 9
Private Const CASE_JSON_CSTRING As Long = 10
Private Const CASE_JSON_FAST As Long = 11
Private Const CASE_JSON_PLAIN As Long = 12
Private Const CASE_MANY_SMALL_CSTRING As Long = 13
Private Const CASE_MANY_SMALL_FAST As Long = 14
Private Const CASE_MANY_SMALL_PLAIN As Long = 15
Private Const CASE_LARGE_APPENDS_CSTRING As Long = 16
Private Const CASE_LARGE_APPENDS_FAST As Long = 17
Private Const CASE_LARGE_APPENDS_PLAIN As Long = 18
Private Const CASE_REAL_SAVE_STRINGS_CSTRING As Long = 19
Private Const CASE_REAL_SAVE_STRINGS_FAST As Long = 20
Private Const CASE_REAL_SAVE_STRINGS_PLAIN As Long = 21
Private Const CASE_SIM_SAVE_TOTAL_CSTRING As Long = 22
Private Const CASE_SIM_SAVE_TOTAL_FAST As Long = 23
Private Const CASE_SIM_SAVE_TOTAL_PLAIN As Long = 24
Private Const CASE_GROWTH_CLEAR_CSTRING As Long = 25
Private Const CASE_GROWTH_CLEAR_FAST As Long = 26
Private Const CASE_GROWTH_CLEAR_PLAIN As Long = 27
Private Const CASE_GROWTH_MANY_FAST_FULL As Long = 28
Private Const CASE_GROWTH_MANY_FAST_USED As Long = 29

Private Const REAL_MAX_SPELLS As Long = 40
Private Const REAL_MAX_INVENTORY As Long = 42
Private Const REAL_MAX_BANK As Long = 42
Private Const REAL_NUM_SKILLS As Long = 24
Private Const REAL_MAX_QUESTS As Long = 5
Private Const REAL_MAX_PETS As Long = 3
Private Const REAL_MAX_SKINS As Long = 42

Private m_OutputFile As Integer
Private m_HasMismatch As Boolean
Private m_LastFastCapacity As Long
Private m_LastFastLength As Long

Private Type tBenchmarkResult
    CStringMs As Long
    FastMs As Long
    PlainMs As Long
    HasPlain As Boolean
End Type

Public Sub RunStringBuilderBenchmark()
    Dim outputPath As String

    On Error GoTo ErrHandler

    outputPath = App.Path & "\" & BENCHMARK_OUTPUT_FILE
    m_OutputFile = FreeFile
    Open outputPath For Output As #m_OutputFile

    m_HasMismatch = False

    PrintLine "String Builder Benchmark"
    PrintLine ""
    PrintLine "Build: compiled EXE or VB6 IDE process"
    PrintLine "Date/time: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    PrintLine "Machine: " & Environ$("COMPUTERNAME")
    PrintLine "Command: " & Command$
    PrintLine "Output file: " & outputPath
    PrintLine "Iteration counts: small SQL=20000, character save=5000, reused Clear=2000, JSON=3000, many small appends=20000, fewer large appends=10000"
    PrintLine "Second-stage counts: real save dynamic strings=10000, simulated full save prep=10000, growth clear=200, growth-copy comparison=100"
    PrintLine ""
    PrintLine "API review:"
    PrintLine "cStringBuilder has Append, AppendNL, Clear, ToString, Length, Capacity, plus ByteLength, ChunkSize, TheString, AppendByVal, Insert, InsertByVal, Remove, Find, and HeapMinimize."
    PrintLine "clsFastString only implements Append, AppendNL, Clear, ToString, Length, and Capacity in this benchmark, so it is not a full drop-in replacement."
    PrintLine "Existing production procedures such as CharacterPersistence quest-save helpers type parameters as cStringBuilder, so replacing callers with clsFastString would not compile without signature changes or an adapter."
    PrintLine "clsFastString.Clear keeps the allocated buffer; cStringBuilder.Clear releases it."
    PrintLine "clsFastString growth copies Len(sTemp), which is the full allocated buffer, not only the used length. That can copy large unused space after growth."
    PrintLine "Full DB save path: not run. This benchmark does not mutate the existing Database.db; full DB timings need a disposable fixture DB with schema and test character rows."
    PrintLine ""

    WarmUpBenchmarks

    RunCase "Small SQL query", 20000, True, CASE_SMALL_SQL_CSTRING, CASE_SMALL_SQL_FAST, CASE_SMALL_SQL_PLAIN
    RunCase "Character save query", 5000, True, CASE_CHARACTER_SAVE_CSTRING, CASE_CHARACTER_SAVE_FAST, CASE_CHARACTER_SAVE_PLAIN
    RunCase "Reused builder with Clear", 2000, True, CASE_REUSED_CLEAR_CSTRING, CASE_REUSED_CLEAR_FAST, CASE_REUSED_CLEAR_PLAIN
    RunCase "JSON encoding and records", 3000, True, CASE_JSON_CSTRING, CASE_JSON_FAST, CASE_JSON_PLAIN
    RunCase "Many small appends", 20000, True, CASE_MANY_SMALL_CSTRING, CASE_MANY_SMALL_FAST, CASE_MANY_SMALL_PLAIN
    RunCase "Fewer large appends", 10000, True, CASE_LARGE_APPENDS_CSTRING, CASE_LARGE_APPENDS_FAST, CASE_LARGE_APPENDS_PLAIN
    RunCase "Real save dynamic SQL strings", 10000, True, CASE_REAL_SAVE_STRINGS_CSTRING, CASE_REAL_SAVE_STRINGS_FAST, CASE_REAL_SAVE_STRINGS_PLAIN
    RunRealSaveShareCase 10000
    RunGrowthClearCase 200
    RunGrowthCopyComparison 100

    If m_HasMismatch Then
        PrintLine "ERROR: One or more benchmark cases produced mismatched output. Performance numbers are not valid."
    Else
        PrintLine "Correctness: all compared outputs matched."
    End If

    Close #m_OutputFile
    m_OutputFile = 0
    Exit Sub

ErrHandler:
    Debug.Print "String builder benchmark failed: " & Err.Number & " " & Err.Description
    If m_OutputFile <> 0 Then Close #m_OutputFile
End Sub

Private Sub WarmUpBenchmarks()
    Dim result As String
    Dim i      As Long

    For i = 1 To 100
        result = BuildSmallSqlCString()
        result = BuildSmallSqlFast()
        result = BuildSmallSqlPlain()
        result = BuildJsonCString()
        result = BuildJsonFast()
        result = BuildJsonPlain()
    Next i
End Sub

Private Sub RunCase(ByVal caseName As String, ByVal iterations As Long, ByVal includePlain As Boolean, ByVal cStringFn As Long, ByVal fastFn As Long, ByVal plainFn As Long)
    Dim expected As String
    Dim actual   As String
    Dim result   As tBenchmarkResult

    expected = InvokeStringBuilder(cStringFn)
    actual = InvokeStringBuilder(fastFn)

    If expected <> actual Then
        m_HasMismatch = True
        PrintLine "Case: " & caseName & ", " & CStr(iterations) & " iterations"
        PrintLine "ERROR: output mismatch between cStringBuilder and clsFastString"
        PrintLine "cStringBuilder length: " & CStr(Len(expected))
        PrintLine "clsFastString length:  " & CStr(Len(actual))
        PrintLine ""
        Exit Sub
    End If

    If includePlain Then
        actual = InvokeStringBuilder(plainFn)
        If expected <> actual Then
            m_HasMismatch = True
            PrintLine "Case: " & caseName & ", " & CStr(iterations) & " iterations"
            PrintLine "ERROR: output mismatch between cStringBuilder and plain concatenation"
            PrintLine "cStringBuilder length: " & CStr(Len(expected))
            PrintLine "Plain & length:        " & CStr(Len(actual))
            PrintLine ""
            Exit Sub
        End If
    End If

    result.CStringMs = Measure(iterations, cStringFn)
    result.FastMs = Measure(iterations, fastFn)
    result.HasPlain = includePlain

    If includePlain Then
        result.PlainMs = Measure(iterations, plainFn)
    End If

    PrintResult caseName, iterations, result
End Sub

Private Function Measure(ByVal iterations As Long, ByVal buildFn As Long) As Long
    Dim i          As Long
    Dim startTicks As Long
    Dim sink       As String

    startTicks = GetTickCount()
    For i = 1 To iterations
        sink = InvokeStringBuilder(buildFn)
    Next i
    Measure = TickDiff(startTicks, GetTickCount())

    If Len(sink) = -1 Then Debug.Print sink
End Function

Private Sub PrintResult(ByVal caseName As String, ByVal iterations As Long, ByRef result As tBenchmarkResult)
    PrintLine "Case: " & caseName & ", " & CStr(iterations) & " iterations"
    PrintLine "cStringBuilder: " & CStr(result.CStringMs) & " ms"
    PrintLine "clsFastString:  " & CStr(result.FastMs) & " ms"

    If result.HasPlain Then
        PrintLine "Plain &:        " & CStr(result.PlainMs) & " ms"
    End If

    PrintLine "Impact: " & ImpactText(result.CStringMs, result.FastMs)
    PrintLine ""
End Sub

Private Function ImpactText(ByVal baseMs As Long, ByVal candidateMs As Long) As String
    Dim pct As Double

    If baseMs <= 0 Then
        ImpactText = "not measurable because cStringBuilder time was 0 ms"
    ElseIf candidateMs < baseMs Then
        pct = ((CDbl(baseMs) - CDbl(candidateMs)) / CDbl(baseMs)) * 100#
        ImpactText = "clsFastString is " & Format$(pct, "0.0") & "% faster than cStringBuilder"
    ElseIf candidateMs > baseMs Then
        pct = ((CDbl(candidateMs) - CDbl(baseMs)) / CDbl(baseMs)) * 100#
        ImpactText = "clsFastString is " & Format$(pct, "0.0") & "% slower than cStringBuilder"
    Else
        ImpactText = "clsFastString matched cStringBuilder"
    End If
End Function

Private Function TickDiff(ByVal startTicks As Long, ByVal endTicks As Long) As Long
    If endTicks >= startTicks Then
        TickDiff = endTicks - startTicks
    Else
        TickDiff = (&H7FFFFFFF - startTicks) + (endTicks And &H7FFFFFFF) + 1
    End If
End Function

Private Sub PrintLine(ByVal text As String)
    Debug.Print text
    Print #m_OutputFile, text
End Sub

Private Function InvokeStringBuilder(ByVal fn As Long) As String
    Select Case fn
        Case 1: InvokeStringBuilder = BuildSmallSqlCString()
        Case 2: InvokeStringBuilder = BuildSmallSqlFast()
        Case 3: InvokeStringBuilder = BuildSmallSqlPlain()
        Case 4: InvokeStringBuilder = BuildCharacterSaveCString()
        Case 5: InvokeStringBuilder = BuildCharacterSaveFast()
        Case 6: InvokeStringBuilder = BuildCharacterSavePlain()
        Case 7: InvokeStringBuilder = BuildReusedClearCString()
        Case 8: InvokeStringBuilder = BuildReusedClearFast()
        Case 9: InvokeStringBuilder = BuildReusedClearPlain()
        Case 10: InvokeStringBuilder = BuildJsonCString()
        Case 11: InvokeStringBuilder = BuildJsonFast()
        Case 12: InvokeStringBuilder = BuildJsonPlain()
        Case 13: InvokeStringBuilder = BuildManySmallCString()
        Case 14: InvokeStringBuilder = BuildManySmallFast()
        Case 15: InvokeStringBuilder = BuildManySmallPlain()
        Case 16: InvokeStringBuilder = BuildLargeAppendsCString()
        Case 17: InvokeStringBuilder = BuildLargeAppendsFast()
        Case 18: InvokeStringBuilder = BuildLargeAppendsPlain()
        Case 19: InvokeStringBuilder = BuildRealSaveStringsCString()
        Case 20: InvokeStringBuilder = BuildRealSaveStringsFast(False)
        Case 21: InvokeStringBuilder = BuildRealSaveStringsPlain()
        Case 22: InvokeStringBuilder = BuildSimulatedSaveTotalCString()
        Case 23: InvokeStringBuilder = BuildSimulatedSaveTotalFast(False)
        Case 24: InvokeStringBuilder = BuildSimulatedSaveTotalPlain()
        Case 25: InvokeStringBuilder = BuildGrowthClearCString()
        Case 26: InvokeStringBuilder = BuildGrowthClearFast(False)
        Case 27: InvokeStringBuilder = BuildGrowthClearPlain()
        Case 28: InvokeStringBuilder = BuildGrowthManyFast(False)
        Case 29: InvokeStringBuilder = BuildGrowthManyFast(True)
    End Select
End Function

Private Function BuildSmallSqlCString() As String
    Dim sb As New cStringBuilder
    AppendSmallSqlCString sb
    BuildSmallSqlCString = sb.ToString
End Function

Private Function BuildSmallSqlFast() As String
    Dim sb As New clsFastString
    AppendSmallSqlFast sb
    BuildSmallSqlFast = sb.ToString
End Function

Private Function BuildSmallSqlPlain() As String
    BuildSmallSqlPlain = BuildSmallSqlPlainText()
End Function

Private Sub AppendSmallSqlCString(ByRef sb As cStringBuilder)
    Dim i As Long

    sb.Append "SELECT "
    For i = 1 To 60
        If i > 1 Then sb.Append ","
        sb.Append "col_"
        sb.Append CStr(i)
    Next i
    sb.Append " FROM user WHERE account_id = ? AND name = ? AND is_banned = 0 ORDER BY level DESC LIMIT 1"
End Sub

Private Sub AppendSmallSqlFast(ByRef sb As clsFastString)
    Dim i As Long

    sb.Append "SELECT "
    For i = 1 To 60
        If i > 1 Then sb.Append ","
        sb.Append "col_"
        sb.Append CStr(i)
    Next i
    sb.Append " FROM user WHERE account_id = ? AND name = ? AND is_banned = 0 ORDER BY level DESC LIMIT 1"
End Sub

Private Function BuildSmallSqlPlainText() As String
    Dim s As String
    Dim i As Long

    s = "SELECT "
    For i = 1 To 60
        If i > 1 Then s = s & ","
        s = s & "col_"
        s = s & CStr(i)
    Next i
    s = s & " FROM user WHERE account_id = ? AND name = ? AND is_banned = 0 ORDER BY level DESC LIMIT 1"
    BuildSmallSqlPlainText = s
End Function

Private Function BuildCharacterSaveCString() As String
    Dim sb As New cStringBuilder
    AppendCharacterSaveCString sb
    BuildCharacterSaveCString = sb.ToString
End Function

Private Function BuildCharacterSaveFast() As String
    Dim sb As New clsFastString
    AppendCharacterSaveFast sb
    BuildCharacterSaveFast = sb.ToString
End Function

Private Function BuildCharacterSavePlain() As String
    BuildCharacterSavePlain = BuildCharacterSavePlainText()
End Function

Private Sub AppendCharacterSaveCString(ByRef sb As cStringBuilder)
    Dim i As Long

    sb.Append "BEGIN TRANSACTION;"
    sb.Append "UPDATE user SET gold=?,bank_gold=?,level=?,exp=?,pos_map=?,pos_x=?,pos_y=? WHERE id=?;"
    sb.Append "DELETE FROM user_inventory WHERE user_id=?;"
    sb.Append "INSERT INTO user_inventory (user_id,slot,obj_id,amount,equipped) VALUES "
    For i = 1 To 30
        If i > 1 Then sb.Append ","
        sb.Append "(?,?,?,?,?)"
    Next i
    sb.Append ";DELETE FROM user_bank WHERE user_id=?;"
    sb.Append "INSERT INTO user_bank (user_id,slot,obj_id,amount) VALUES "
    For i = 1 To 40
        If i > 1 Then sb.Append ","
        sb.Append "(?,?,?,?)"
    Next i
    sb.Append ";REPLACE INTO user_spells (user_id,slot,spell_id) VALUES "
    For i = 1 To 35
        If i > 1 Then sb.Append ","
        sb.Append "(?,?,?)"
    Next i
    sb.Append ";REPLACE INTO user_skills (user_id,skill_id,value,elapsed) VALUES "
    For i = 1 To 21
        If i > 1 Then sb.Append ","
        sb.Append "(?,?,?,?)"
    Next i
    sb.Append ";REPLACE INTO user_quests (user_id,quest_id,state,kills) VALUES "
    For i = 1 To 25
        If i > 1 Then sb.Append ","
        sb.Append "(?,?,?,?)"
    Next i
    sb.Append ";REPLACE INTO user_pets (user_id,slot,npc_id,level,hp) VALUES "
    For i = 1 To 12
        If i > 1 Then sb.Append ","
        sb.Append "(?,?,?,?,?)"
    Next i
    sb.Append ";COMMIT;"
End Sub

Private Sub AppendCharacterSaveFast(ByRef sb As clsFastString)
    Dim i As Long

    sb.Append "BEGIN TRANSACTION;"
    sb.Append "UPDATE user SET gold=?,bank_gold=?,level=?,exp=?,pos_map=?,pos_x=?,pos_y=? WHERE id=?;"
    sb.Append "DELETE FROM user_inventory WHERE user_id=?;"
    sb.Append "INSERT INTO user_inventory (user_id,slot,obj_id,amount,equipped) VALUES "
    For i = 1 To 30
        If i > 1 Then sb.Append ","
        sb.Append "(?,?,?,?,?)"
    Next i
    sb.Append ";DELETE FROM user_bank WHERE user_id=?;"
    sb.Append "INSERT INTO user_bank (user_id,slot,obj_id,amount) VALUES "
    For i = 1 To 40
        If i > 1 Then sb.Append ","
        sb.Append "(?,?,?,?)"
    Next i
    sb.Append ";REPLACE INTO user_spells (user_id,slot,spell_id) VALUES "
    For i = 1 To 35
        If i > 1 Then sb.Append ","
        sb.Append "(?,?,?)"
    Next i
    sb.Append ";REPLACE INTO user_skills (user_id,skill_id,value,elapsed) VALUES "
    For i = 1 To 21
        If i > 1 Then sb.Append ","
        sb.Append "(?,?,?,?)"
    Next i
    sb.Append ";REPLACE INTO user_quests (user_id,quest_id,state,kills) VALUES "
    For i = 1 To 25
        If i > 1 Then sb.Append ","
        sb.Append "(?,?,?,?)"
    Next i
    sb.Append ";REPLACE INTO user_pets (user_id,slot,npc_id,level,hp) VALUES "
    For i = 1 To 12
        If i > 1 Then sb.Append ","
        sb.Append "(?,?,?,?,?)"
    Next i
    sb.Append ";COMMIT;"
End Sub

Private Function BuildCharacterSavePlainText() As String
    Dim s As String
    Dim i As Long

    s = "BEGIN TRANSACTION;"
    s = s & "UPDATE user SET gold=?,bank_gold=?,level=?,exp=?,pos_map=?,pos_x=?,pos_y=? WHERE id=?;"
    s = s & "DELETE FROM user_inventory WHERE user_id=?;"
    s = s & "INSERT INTO user_inventory (user_id,slot,obj_id,amount,equipped) VALUES "
    For i = 1 To 30
        If i > 1 Then s = s & ","
        s = s & "(?,?,?,?,?)"
    Next i
    s = s & ";DELETE FROM user_bank WHERE user_id=?;"
    s = s & "INSERT INTO user_bank (user_id,slot,obj_id,amount) VALUES "
    For i = 1 To 40
        If i > 1 Then s = s & ","
        s = s & "(?,?,?,?)"
    Next i
    s = s & ";REPLACE INTO user_spells (user_id,slot,spell_id) VALUES "
    For i = 1 To 35
        If i > 1 Then s = s & ","
        s = s & "(?,?,?)"
    Next i
    s = s & ";REPLACE INTO user_skills (user_id,skill_id,value,elapsed) VALUES "
    For i = 1 To 21
        If i > 1 Then s = s & ","
        s = s & "(?,?,?,?)"
    Next i
    s = s & ";REPLACE INTO user_quests (user_id,quest_id,state,kills) VALUES "
    For i = 1 To 25
        If i > 1 Then s = s & ","
        s = s & "(?,?,?,?)"
    Next i
    s = s & ";REPLACE INTO user_pets (user_id,slot,npc_id,level,hp) VALUES "
    For i = 1 To 12
        If i > 1 Then s = s & ","
        s = s & "(?,?,?,?,?)"
    Next i
    s = s & ";COMMIT;"
    BuildCharacterSavePlainText = s
End Function

Private Function BuildReusedClearCString() As String
    Dim sb As New cStringBuilder
    Dim i  As Long

    For i = 1 To 25
        sb.Clear
        AppendSmallSqlCString sb
    Next i

    BuildReusedClearCString = sb.ToString
End Function

Private Function BuildReusedClearFast() As String
    Dim sb As New clsFastString
    Dim i  As Long

    For i = 1 To 25
        sb.Clear
        AppendSmallSqlFast sb
    Next i

    BuildReusedClearFast = sb.ToString
End Function

Private Function BuildReusedClearPlain() As String
    Dim s As String
    Dim i As Long

    For i = 1 To 25
        s = vbNullString
        s = BuildSmallSqlPlainText()
    Next i

    BuildReusedClearPlain = s
End Function

Private Function BuildJsonCString() As String
    Dim sb As New cStringBuilder
    AppendJsonCString sb
    BuildJsonCString = sb.ToString
End Function

Private Function BuildJsonFast() As String
    Dim sb As New clsFastString
    AppendJsonFast sb
    BuildJsonFast = sb.ToString
End Function

Private Function BuildJsonPlain() As String
    BuildJsonPlain = BuildJsonPlainText()
End Function

Private Sub AppendJsonCString(ByRef sb As cStringBuilder)
    Dim i      As Long
    Dim sample As String

    sb.Append "({""Records"":["
    For i = 1 To 20
        If i > 1 Then sb.Append "," & vbNewLine
        sample = JsonSampleText(i)
        sb.Append "{""id"":"""
        sb.Append CStr(i)
        sb.Append """,""name"":"""
        sb.Append EncodeJsonCString(sample)
        sb.Append """,""description"":"""
        sb.Append EncodeJsonCString(sample & sample & " long field " & String$(50, Chr$(65 + (i Mod 26))))
        sb.Append """}"
    Next i
    sb.Append vbNewLine & "],""RecordCount"":""20""})"
End Sub

Private Sub AppendJsonFast(ByRef sb As clsFastString)
    Dim i      As Long
    Dim sample As String

    sb.Append "({""Records"":["
    For i = 1 To 20
        If i > 1 Then sb.Append "," & vbNewLine
        sample = JsonSampleText(i)
        sb.Append "{""id"":"""
        sb.Append CStr(i)
        sb.Append """,""name"":"""
        sb.Append EncodeJsonFast(sample)
        sb.Append """,""description"":"""
        sb.Append EncodeJsonFast(sample & sample & " long field " & String$(50, Chr$(65 + (i Mod 26))))
        sb.Append """}"
    Next i
    sb.Append vbNewLine & "],""RecordCount"":""20""})"
End Sub

Private Function BuildJsonPlainText() As String
    Dim s      As String
    Dim i      As Long
    Dim sample As String

    s = "({""Records"":["
    For i = 1 To 20
        If i > 1 Then s = s & "," & vbNewLine
        sample = JsonSampleText(i)
        s = s & "{""id"":"""
        s = s & CStr(i)
        s = s & """,""name"":"""
        s = s & EncodeJsonPlain(sample)
        s = s & """,""description"":"""
        s = s & EncodeJsonPlain(sample & sample & " long field " & String$(50, Chr$(65 + (i Mod 26))))
        s = s & """}"
    Next i
    s = s & vbNewLine & "],""RecordCount"":""20""})"
    BuildJsonPlainText = s
End Function

Private Function JsonSampleText(ByVal index As Long) As String
    Select Case index Mod 4
        Case 0
            JsonSampleText = "normal AO character field " & CStr(index)
        Case 1
            JsonSampleText = "quote "" slash / backslash \ tab" & vbTab & " end"
        Case 2
            JsonSampleText = "line one" & vbCrLf & "line two" & vbLf & "field " & CStr(index)
        Case Else
            JsonSampleText = "long-" & String$(100, Chr$(65 + (index Mod 26))) & "-tail"
    End Select
End Function

Private Function EncodeJsonCString(ByVal text As String) As String
    Dim sb As New cStringBuilder
    AppendEncodedJsonCString sb, text
    EncodeJsonCString = sb.ToString
End Function

Private Function EncodeJsonFast(ByVal text As String) As String
    Dim sb As New clsFastString
    AppendEncodedJsonFast sb, text
    EncodeJsonFast = sb.ToString
End Function

Private Function EncodeJsonPlain(ByVal text As String) As String
    Dim s As String
    Dim i As Long
    Dim c As String

    For i = 1 To Len(text)
        c = Mid$(text, i, 1)
        Select Case c
            Case """": s = s & "\"""
            Case "\": s = s & "\\"
            Case "/": s = s & "\/"
            Case vbTab: s = s & "\t"
            Case vbCr: s = s & "\r"
            Case vbLf: s = s & "\n"
            Case Else: s = s & c
        End Select
    Next i

    EncodeJsonPlain = s
End Function

Private Sub AppendEncodedJsonCString(ByRef sb As cStringBuilder, ByVal text As String)
    Dim i As Long
    Dim c As String

    For i = 1 To Len(text)
        c = Mid$(text, i, 1)
        Select Case c
            Case """": sb.Append "\"""
            Case "\": sb.Append "\\"
            Case "/": sb.Append "\/"
            Case vbTab: sb.Append "\t"
            Case vbCr: sb.Append "\r"
            Case vbLf: sb.Append "\n"
            Case Else: sb.Append c
        End Select
    Next i
End Sub

Private Sub AppendEncodedJsonFast(ByRef sb As clsFastString, ByVal text As String)
    Dim i As Long
    Dim c As String

    For i = 1 To Len(text)
        c = Mid$(text, i, 1)
        Select Case c
            Case """": sb.Append "\"""
            Case "\": sb.Append "\\"
            Case "/": sb.Append "\/"
            Case vbTab: sb.Append "\t"
            Case vbCr: sb.Append "\r"
            Case vbLf: sb.Append "\n"
            Case Else: sb.Append c
        End Select
    Next i
End Sub

Private Function BuildManySmallCString() As String
    Dim sb As New cStringBuilder
    Dim i  As Long

    For i = 1 To 300
        sb.Append "a"
        sb.Append ","
        sb.Append "x"
    Next i

    BuildManySmallCString = sb.ToString
End Function

Private Function BuildManySmallFast() As String
    Dim sb As New clsFastString
    Dim i  As Long

    For i = 1 To 300
        sb.Append "a"
        sb.Append ","
        sb.Append "x"
    Next i

    BuildManySmallFast = sb.ToString
End Function

Private Function BuildManySmallPlain() As String
    Dim s As String
    Dim i As Long

    For i = 1 To 300
        s = s & "a"
        s = s & ","
        s = s & "x"
    Next i

    BuildManySmallPlain = s
End Function

Private Function BuildLargeAppendsCString() As String
    Dim sb As New cStringBuilder
    AppendLargeAppendsCString sb
    BuildLargeAppendsCString = sb.ToString
End Function

Private Function BuildLargeAppendsFast() As String
    Dim sb As New clsFastString
    AppendLargeAppendsFast sb
    BuildLargeAppendsFast = sb.ToString
End Function

Private Function BuildLargeAppendsPlain() As String
    BuildLargeAppendsPlain = String$(100, Asc("A")) & String$(1000, Asc("B")) & String$(10000, Asc("C"))
End Function

Private Sub AppendLargeAppendsCString(ByRef sb As cStringBuilder)
    sb.Append String$(100, Asc("A"))
    sb.Append String$(1000, Asc("B"))
    sb.Append String$(10000, Asc("C"))
End Sub

Private Sub AppendLargeAppendsFast(ByRef sb As clsFastString)
    sb.Append String$(100, Asc("A"))
    sb.Append String$(1000, Asc("B"))
    sb.Append String$(10000, Asc("C"))
End Sub

Private Sub RunRealSaveShareCase(ByVal iterations As Long)
    Dim expectedStrings As String
    Dim expectedTotal   As String
    Dim fastStrings     As String
    Dim fastTotal       As String
    Dim plainStrings    As String
    Dim plainTotal      As String
    Dim cStringMs       As Long
    Dim cTotalMs        As Long
    Dim fastStringMs    As Long
    Dim fastTotalMs     As Long
    Dim plainStringMs   As Long
    Dim plainTotalMs    As Long

    expectedStrings = BuildRealSaveStringsCString()
    fastStrings = BuildRealSaveStringsFast(False)
    plainStrings = BuildRealSaveStringsPlain()
    expectedTotal = BuildSimulatedSaveTotalCString()
    fastTotal = BuildSimulatedSaveTotalFast(False)
    plainTotal = BuildSimulatedSaveTotalPlain()

    If expectedStrings <> fastStrings Or expectedStrings <> plainStrings Or expectedTotal <> fastTotal Or expectedTotal <> plainTotal Then
        m_HasMismatch = True
        PrintLine "Case: Simulated SaveChangesInUser split, " & CStr(iterations) & " iterations"
        PrintLine "ERROR: output mismatch in real-path split benchmark"
        PrintLine ""
        Exit Sub
    End If

    cStringMs = Measure(iterations, CASE_REAL_SAVE_STRINGS_CSTRING)
    fastStringMs = Measure(iterations, CASE_REAL_SAVE_STRINGS_FAST)
    plainStringMs = Measure(iterations, CASE_REAL_SAVE_STRINGS_PLAIN)
    cTotalMs = Measure(iterations, CASE_SIM_SAVE_TOTAL_CSTRING)
    fastTotalMs = Measure(iterations, CASE_SIM_SAVE_TOTAL_FAST)
    plainTotalMs = Measure(iterations, CASE_SIM_SAVE_TOTAL_PLAIN)

    PrintLine "Case: Simulated SaveChangesInUser split, " & CStr(iterations) & " iterations"
    PrintLine "cStringBuilder string construction: " & CStr(cStringMs) & " ms"
    PrintLine "cStringBuilder full save prep:      " & CStr(cTotalMs) & " ms"
    PrintLine "cStringBuilder string share:        " & PercentText(cStringMs, cTotalMs)
    PrintLine "clsFastString string construction:  " & CStr(fastStringMs) & " ms"
    PrintLine "clsFastString full save prep:       " & CStr(fastTotalMs) & " ms"
    PrintLine "clsFastString string share:         " & PercentText(fastStringMs, fastTotalMs)
    PrintLine "Plain & string construction:        " & CStr(plainStringMs) & " ms"
    PrintLine "Plain & full save prep:             " & CStr(plainTotalMs) & " ms"
    PrintLine "Plain & string share:               " & PercentText(plainStringMs, plainTotalMs)
    PrintLine "Full DB save path: not run against Database.db; no database rows were inserted, updated, or deleted."
    PrintLine ""
End Sub

Private Sub RunGrowthClearCase(ByVal iterations As Long)
    Dim expected As String
    Dim actual   As String
    Dim cMs      As Long
    Dim fastMs   As Long
    Dim plainMs  As Long

    expected = BuildGrowthClearCString()
    actual = BuildGrowthClearFast(False)
    If expected <> actual Then
        m_HasMismatch = True
        PrintLine "Case: Growth beyond 50000 then Clear"
        PrintLine "ERROR: output mismatch"
        PrintLine ""
        Exit Sub
    End If

    cMs = Measure(iterations, CASE_GROWTH_CLEAR_CSTRING)
    fastMs = Measure(iterations, CASE_GROWTH_CLEAR_FAST)
    plainMs = Measure(iterations, CASE_GROWTH_CLEAR_PLAIN)
    actual = BuildGrowthClearFast(False)

    PrintLine "Case: Growth beyond 50000 then Clear, " & CStr(iterations) & " iterations"
    PrintLine "cStringBuilder: " & CStr(cMs) & " ms"
    PrintLine "clsFastString:  " & CStr(fastMs) & " ms"
    PrintLine "Plain &:        " & CStr(plainMs) & " ms"
    PrintLine "clsFastString retained capacity after Clear: " & CStr(m_LastFastCapacity) & " chars"
    PrintLine "clsFastString length after final small append: " & CStr(m_LastFastLength) & " chars"
    PrintLine "Impact: " & ImpactText(cMs, fastMs)
    PrintLine ""
End Sub

Private Sub RunGrowthCopyComparison(ByVal iterations As Long)
    Dim fullCopyResult As String
    Dim usedCopyResult As String
    Dim fullMs         As Long
    Dim usedMs         As Long
    Dim fullCapacity   As Long
    Dim usedCapacity   As Long

    fullCopyResult = BuildGrowthManyFast(False)
    fullCapacity = m_LastFastCapacity
    usedCopyResult = BuildGrowthManyFast(True)
    usedCapacity = m_LastFastCapacity

    If fullCopyResult <> usedCopyResult Then
        m_HasMismatch = True
        PrintLine "Case: clsFastString growth copy strategy"
        PrintLine "ERROR: output mismatch between full-buffer and used-length copy"
        PrintLine ""
        Exit Sub
    End If

    fullMs = Measure(iterations, CASE_GROWTH_MANY_FAST_FULL)
    usedMs = Measure(iterations, CASE_GROWTH_MANY_FAST_USED)

    PrintLine "Case: clsFastString growth copy strategy, " & CStr(iterations) & " iterations"
    PrintLine "Full-buffer copy: " & CStr(fullMs) & " ms, retained capacity " & CStr(fullCapacity) & " chars"
    PrintLine "Used-length copy: " & CStr(usedMs) & " ms, retained capacity " & CStr(usedCapacity) & " chars"
    If fullMs > 0 Then
        PrintLine "Impact: used-length copy is " & Format$(((CDbl(fullMs) - CDbl(usedMs)) / CDbl(fullMs)) * 100#, "0.0") & "% faster than full-buffer copy"
    Else
        PrintLine "Impact: not measurable because full-buffer copy time was 0 ms"
    End If
    PrintLine ""
End Sub

Private Function PercentText(ByVal partMs As Long, ByVal totalMs As Long) As String
    If totalMs <= 0 Then
        PercentText = "not measurable"
    Else
        PercentText = Format$((CDbl(partMs) / CDbl(totalMs)) * 100#, "0.0") & "%"
    End If
End Function

Private Function BuildRealSaveStringsCString() As String
    Dim sb As New cStringBuilder
    AppendRealSaveDynamicSqlCString sb
    BuildRealSaveStringsCString = sb.ToString
End Function

Private Function BuildRealSaveStringsFast(ByVal copyUsedOnly As Boolean) As String
    Dim sb As New clsFastString
    sb.CopyUsedOnly = copyUsedOnly
    AppendRealSaveDynamicSqlFast sb
    BuildRealSaveStringsFast = sb.ToString
    m_LastFastCapacity = sb.Capacity
    m_LastFastLength = sb.Length
End Function

Private Function BuildRealSaveStringsPlain() As String
    Dim s As String
    AppendRealSaveDynamicSqlPlain s
    BuildRealSaveStringsPlain = s
End Function

Private Sub AppendRealSaveDynamicSqlCString(ByRef sb As cStringBuilder)
    AppendDirtySkillsCString sb, REAL_NUM_SKILLS
    sb.AppendNL ""
    AppendDirtyQuestsCString sb, REAL_MAX_QUESTS, 4, 3
    sb.AppendNL ""
    AppendQuestsDoneCString sb, 160
    sb.AppendNL ""
    AppendInventorySkinsCString sb, REAL_MAX_SKINS, 30, 12
End Sub

Private Sub AppendRealSaveDynamicSqlFast(ByRef sb As clsFastString)
    AppendDirtySkillsFast sb, REAL_NUM_SKILLS
    sb.AppendNL ""
    AppendDirtyQuestsFast sb, REAL_MAX_QUESTS, 4, 3
    sb.AppendNL ""
    AppendQuestsDoneFast sb, 160
    sb.AppendNL ""
    AppendInventorySkinsFast sb, REAL_MAX_SKINS, 30, 12
End Sub

Private Sub AppendRealSaveDynamicSqlPlain(ByRef s As String)
    s = s & BuildDirtySkillsPlain(REAL_NUM_SKILLS) & vbCrLf
    s = s & BuildDirtyQuestsPlain(REAL_MAX_QUESTS, 4, 3) & vbCrLf
    s = s & BuildQuestsDonePlain(160) & vbCrLf
    s = s & BuildInventorySkinsPlain(REAL_MAX_SKINS, 30, 12)
End Sub

Private Function BuildSimulatedSaveTotalCString() As String
    Dim sb     As New cStringBuilder
    Dim params As Long

    params = SimulateSaveParameterPrep()
    sb.Append "params="
    sb.Append CStr(params)
    sb.Append "|"
    AppendRealSaveDynamicSqlCString sb
    BuildSimulatedSaveTotalCString = sb.ToString
End Function

Private Function BuildSimulatedSaveTotalFast(ByVal copyUsedOnly As Boolean) As String
    Dim sb     As New clsFastString
    Dim params As Long

    sb.CopyUsedOnly = copyUsedOnly
    params = SimulateSaveParameterPrep()
    sb.Append "params="
    sb.Append CStr(params)
    sb.Append "|"
    AppendRealSaveDynamicSqlFast sb
    BuildSimulatedSaveTotalFast = sb.ToString
    m_LastFastCapacity = sb.Capacity
    m_LastFastLength = sb.Length
End Function

Private Function BuildSimulatedSaveTotalPlain() As String
    Dim s      As String
    Dim params As Long

    params = SimulateSaveParameterPrep()
    s = "params=" & CStr(params) & "|"
    AppendRealSaveDynamicSqlPlain s
    BuildSimulatedSaveTotalPlain = s
End Function

Private Function SimulateSaveParameterPrep() As Long
    Dim params() As Variant
    Dim i        As Long
    Dim p        As Long
    Dim checksum As Long

    ReDim params(0 To 61)
    For i = 0 To 61
        params(i) = i + 1000
        checksum = checksum + CLng(params(i))
    Next i

    ReDim params(0 To REAL_MAX_SPELLS * 3 - 1)
    For i = 1 To REAL_MAX_SPELLS
        params(p) = 1001
        params(p + 1) = i
        params(p + 2) = (i Mod 120) + 1
        checksum = checksum + CLng(params(p + 2))
        p = p + 3
    Next i

    p = 0
    ReDim params(0 To REAL_MAX_INVENTORY * 6 - 1)
    For i = 1 To REAL_MAX_INVENTORY
        params(p) = 1001
        params(p + 1) = i
        params(p + 2) = 3000 + i
        params(p + 3) = (i Mod 99) + 1
        params(p + 4) = (i Mod 2)
        params(p + 5) = i Mod 7
        checksum = checksum + CLng(params(p + 2)) + CLng(params(p + 3))
        p = p + 6
    Next i

    p = 0
    ReDim params(0 To REAL_MAX_BANK * 5 - 1)
    For i = 1 To REAL_MAX_BANK
        params(p) = 1001
        params(p + 1) = i
        params(p + 2) = 5000 + i
        params(p + 3) = (i Mod 50) + 1
        params(p + 4) = i Mod 11
        checksum = checksum + CLng(params(p + 2)) + CLng(params(p + 3))
        p = p + 5
    Next i

    p = 0
    ReDim params(0 To REAL_NUM_SKILLS * 3 - 1)
    For i = 1 To REAL_NUM_SKILLS
        params(p) = 1001
        params(p + 1) = i
        params(p + 2) = i Mod 100
        checksum = checksum + CLng(params(p + 2))
        p = p + 3
    Next i

    p = 0
    ReDim params(0 To REAL_MAX_PETS * 3 - 1)
    For i = 1 To REAL_MAX_PETS
        params(p) = 1001
        params(p + 1) = i
        params(p + 2) = 800 + i
        checksum = checksum + CLng(params(p + 2))
        p = p + 3
    Next i

    SimulateSaveParameterPrep = checksum
End Function

Private Sub AppendDirtySkillsCString(ByRef sb As cStringBuilder, ByVal dirtyCount As Long)
    Dim i As Long
    sb.Append "REPLACE INTO skillpoint (user_id, number, value) VALUES "
    For i = 1 To dirtyCount
        sb.Append "(?, ?, ?)"
        If i < dirtyCount Then sb.Append ", "
    Next i
End Sub

Private Sub AppendDirtySkillsFast(ByRef sb As clsFastString, ByVal dirtyCount As Long)
    Dim i As Long
    sb.Append "REPLACE INTO skillpoint (user_id, number, value) VALUES "
    For i = 1 To dirtyCount
        sb.Append "(?, ?, ?)"
        If i < dirtyCount Then sb.Append ", "
    Next i
End Sub

Private Function BuildDirtySkillsPlain(ByVal dirtyCount As Long) As String
    Dim s As String
    Dim i As Long
    s = "REPLACE INTO skillpoint (user_id, number, value) VALUES "
    For i = 1 To dirtyCount
        s = s & "(?, ?, ?)"
        If i < dirtyCount Then s = s & ", "
    Next i
    BuildDirtySkillsPlain = s
End Function

Private Sub AppendDirtyQuestsCString(ByRef sb As cStringBuilder, ByVal questCount As Long, ByVal requiredNpcs As Long, ByVal requiredTargets As Long)
    Dim slot As Long
    Dim k    As Long
    sb.Append "REPLACE INTO quest (user_id, number, quest_id, npcs, npcstarget) VALUES "
    For slot = 1 To questCount
        sb.Append "("
        sb.Append "1001, "
        sb.Append CStr(slot)
        sb.Append ", "
        sb.Append CStr(200 + slot)
        sb.Append ", '"
        For k = 1 To requiredNpcs
            sb.Append CStr((slot * 10) + k)
            If k < requiredNpcs Then sb.Append "-"
        Next k
        sb.Append "', '"
        For k = 1 To requiredTargets
            sb.Append CStr((slot * 20) + k)
            If k < requiredTargets Then sb.Append "-"
        Next k
        sb.Append "')"
        If slot < questCount Then sb.Append ", "
    Next slot
End Sub

Private Sub AppendDirtyQuestsFast(ByRef sb As clsFastString, ByVal questCount As Long, ByVal requiredNpcs As Long, ByVal requiredTargets As Long)
    Dim slot As Long
    Dim k    As Long
    sb.Append "REPLACE INTO quest (user_id, number, quest_id, npcs, npcstarget) VALUES "
    For slot = 1 To questCount
        sb.Append "("
        sb.Append "1001, "
        sb.Append CStr(slot)
        sb.Append ", "
        sb.Append CStr(200 + slot)
        sb.Append ", '"
        For k = 1 To requiredNpcs
            sb.Append CStr((slot * 10) + k)
            If k < requiredNpcs Then sb.Append "-"
        Next k
        sb.Append "', '"
        For k = 1 To requiredTargets
            sb.Append CStr((slot * 20) + k)
            If k < requiredTargets Then sb.Append "-"
        Next k
        sb.Append "')"
        If slot < questCount Then sb.Append ", "
    Next slot
End Sub

Private Function BuildDirtyQuestsPlain(ByVal questCount As Long, ByVal requiredNpcs As Long, ByVal requiredTargets As Long) As String
    Dim s    As String
    Dim slot As Long
    Dim k    As Long
    s = "REPLACE INTO quest (user_id, number, quest_id, npcs, npcstarget) VALUES "
    For slot = 1 To questCount
        s = s & "(" & "1001, " & CStr(slot) & ", " & CStr(200 + slot) & ", '"
        For k = 1 To requiredNpcs
            s = s & CStr((slot * 10) + k)
            If k < requiredNpcs Then s = s & "-"
        Next k
        s = s & "', '"
        For k = 1 To requiredTargets
            s = s & CStr((slot * 20) + k)
            If k < requiredTargets Then s = s & "-"
        Next k
        s = s & "')"
        If slot < questCount Then s = s & ", "
    Next slot
    BuildDirtyQuestsPlain = s
End Function

Private Sub AppendQuestsDoneCString(ByRef sb As cStringBuilder, ByVal doneCount As Long)
    Dim i As Long
    sb.Append "REPLACE INTO quest_done (user_id, quest_id) VALUES "
    For i = 1 To doneCount
        sb.Append "(?, ?)"
        If i < doneCount Then sb.Append ", "
    Next i
End Sub

Private Sub AppendQuestsDoneFast(ByRef sb As clsFastString, ByVal doneCount As Long)
    Dim i As Long
    sb.Append "REPLACE INTO quest_done (user_id, quest_id) VALUES "
    For i = 1 To doneCount
        sb.Append "(?, ?)"
        If i < doneCount Then sb.Append ", "
    Next i
End Sub

Private Function BuildQuestsDonePlain(ByVal doneCount As Long) As String
    Dim s As String
    Dim i As Long
    s = "REPLACE INTO quest_done (user_id, quest_id) VALUES "
    For i = 1 To doneCount
        s = s & "(?, ?)"
        If i < doneCount Then s = s & ", "
    Next i
    BuildQuestsDonePlain = s
End Function

Private Sub AppendInventorySkinsCString(ByRef sb As cStringBuilder, ByVal slotCount As Long, ByVal upsertCount As Long, ByVal deleteCount As Long)
    Dim i As Long
    sb.Append "INSERT INTO inventory_item_skins (user_id, skin_id, type_skin, skin_equipped) VALUES "
    For i = 1 To upsertCount
        If i > 1 Then sb.Append ","
        sb.Append "(?, ?, ?, ?)"
    Next i
    sb.Append " ON CONFLICT(user_id, skin_id) DO UPDATE SET type_skin = excluded.type_skin, skin_equipped = excluded.skin_equipped"
    sb.AppendNL ""
    sb.Append "DELETE FROM inventory_item_skins WHERE user_id = ? AND skin_id IN ("
    For i = 1 To deleteCount
        If i > 1 Then sb.Append ","
        sb.Append "?"
    Next i
    sb.Append ")"
End Sub

Private Sub AppendInventorySkinsFast(ByRef sb As clsFastString, ByVal slotCount As Long, ByVal upsertCount As Long, ByVal deleteCount As Long)
    Dim i As Long
    sb.Append "INSERT INTO inventory_item_skins (user_id, skin_id, type_skin, skin_equipped) VALUES "
    For i = 1 To upsertCount
        If i > 1 Then sb.Append ","
        sb.Append "(?, ?, ?, ?)"
    Next i
    sb.Append " ON CONFLICT(user_id, skin_id) DO UPDATE SET type_skin = excluded.type_skin, skin_equipped = excluded.skin_equipped"
    sb.AppendNL ""
    sb.Append "DELETE FROM inventory_item_skins WHERE user_id = ? AND skin_id IN ("
    For i = 1 To deleteCount
        If i > 1 Then sb.Append ","
        sb.Append "?"
    Next i
    sb.Append ")"
End Sub

Private Function BuildInventorySkinsPlain(ByVal slotCount As Long, ByVal upsertCount As Long, ByVal deleteCount As Long) As String
    Dim s As String
    Dim i As Long
    s = "INSERT INTO inventory_item_skins (user_id, skin_id, type_skin, skin_equipped) VALUES "
    For i = 1 To upsertCount
        If i > 1 Then s = s & ","
        s = s & "(?, ?, ?, ?)"
    Next i
    s = s & " ON CONFLICT(user_id, skin_id) DO UPDATE SET type_skin = excluded.type_skin, skin_equipped = excluded.skin_equipped"
    s = s & vbCrLf & "DELETE FROM inventory_item_skins WHERE user_id = ? AND skin_id IN ("
    For i = 1 To deleteCount
        If i > 1 Then s = s & ","
        s = s & "?"
    Next i
    s = s & ")"
    BuildInventorySkinsPlain = s
End Function

Private Function BuildGrowthClearCString() As String
    Dim sb As New cStringBuilder
    Dim i  As Long
    For i = 1 To 8
        sb.Clear
        sb.Append String$(60000, Asc("X"))
        sb.Clear
        sb.Append "small-after-clear"
    Next i
    BuildGrowthClearCString = sb.ToString
End Function

Private Function BuildGrowthClearFast(ByVal copyUsedOnly As Boolean) As String
    Dim sb As New clsFastString
    Dim i  As Long
    sb.CopyUsedOnly = copyUsedOnly
    For i = 1 To 8
        sb.Clear
        sb.Append String$(60000, Asc("X"))
        sb.Clear
        sb.Append "small-after-clear"
    Next i
    BuildGrowthClearFast = sb.ToString
    m_LastFastCapacity = sb.Capacity
    m_LastFastLength = sb.Length
End Function

Private Function BuildGrowthClearPlain() As String
    Dim s As String
    Dim i As Long
    For i = 1 To 8
        s = vbNullString
        s = String$(60000, Asc("X"))
        s = vbNullString
        s = "small-after-clear"
    Next i
    BuildGrowthClearPlain = s
End Function

Private Function BuildGrowthManyFast(ByVal copyUsedOnly As Boolean) As String
    Dim sb As New clsFastString
    Dim i  As Long
    sb.CopyUsedOnly = copyUsedOnly
    For i = 1 To 25
        sb.Append String$(49000, Asc("A") + (i Mod 20))
        sb.Append CStr(i)
    Next i
    BuildGrowthManyFast = sb.ToString
    m_LastFastCapacity = sb.Capacity
    m_LastFastLength = sb.Length
End Function
