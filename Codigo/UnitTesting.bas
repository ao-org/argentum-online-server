Attribute VB_Name = "UnitTesting"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Marquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
#If UNIT_TEST = 1 Then
    Public public_key      As String
    Public private_key     As String
    Public encrypted_token As String
    Public decrypted_token As String
    Public character_name  As String

    ' Test runner result tracking state
    Private TotalTests     As Long
    Private PassedTests    As Long
    Private FailedTests    As Long
    Private FailedTestNames() As String
    Private FailedTestCount As Long
    Private TotalElapsed   As Double
    Private LastCheckpoint As String
    Private CaptureTaming As Boolean
    Private TamingCreates As Collection
    Private TamingRecipients As Collection
    Private TamingRoutesValid As Boolean

        Private Const SUITE_COUNT As Integer = 44

' Independent of Init: startup can fail before the test runner is initialized.
Public Sub Progress(ByVal checkpoint As String, Optional ByVal startRun As Boolean = False)
    Dim number As Long, description As String, source As String
    Dim helpFile As String, helpContext As Long
    number = Err.Number: description = Err.Description: source = Err.Source
    helpFile = Err.HelpFile: helpContext = Err.HelpContext
    On Error GoTo Failed
    Dim f As Integer, opened As Boolean
    LastCheckpoint = checkpoint
    f = FreeFile
    If startRun Then
        Open App.Path & "\test_progress.txt" For Output As #f
    Else
        Open App.Path & "\test_progress.txt" For Append As #f
    End If
    opened = True
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & checkpoint
Done:
    On Error Resume Next
    If opened Then Close #f
    ' Logging must not erase the error that an existing handler will report.
    Err.Number = number: Err.Description = description: Err.Source = source
    Err.HelpFile = helpFile: Err.HelpContext = helpContext
    Exit Sub
Failed:
    Debug.Print "Cannot write test_progress.txt: " & Err.Number & " " & Err.Description
    Resume Done
End Sub

Public Sub ProgressError(ByVal operation As String, ByVal number As Long, ByVal description As String, Optional ByVal fatal As Boolean = False)
    Dim detail As String
    detail = operation & ": Err.Number=" & CStr(number) & "; Err.Description=" & description & "; last checkpoint=" & LastCheckpoint
    Call Progress("ERROR " & detail)
    If fatal Then
        Call RunTestError("Fatal startup", detail)
        Call WriteResultsToFile(App.Path & "\test_results.txt")
    End If
End Sub

Public Sub Init()
    On Error GoTo Init_Err
    Dim failure As String
    TotalTests = 0
    PassedTests = 0
    FailedTests = 0
    FailedTestCount = 0
    TotalElapsed = 0
    ReDim FailedTestNames(0)
    ' PYMMO=0 never loads PrivateKey. Use a local fixture, not an empty AES key.
    #If PYMMO = 0 Then
    private_key = "000102030405060708090A0B0C0D0E0F"
    #Else
    private_key = PrivateKey
    #End If
    character_name = "seneca"
    character_name = RandomName(16)
    'Hardcoded token for unit testing...

    decrypted_token = "G7H5wKOKZvebZxHtnkRtJNvL/AHWEw3dHCyBTzXVvdTe3bQAJHePsFfV/Ecgm9Wk"
    Call Progress("BEFORE test token encryption")
    encrypted_token = AO20CryptoSysWrapper.ENCRYPT(private_key, decrypted_token)
    Call Progress("AFTER test token encryption")
    public_key = mid$(decrypted_token, 1, 16)
    
    'Add a fake token to be using when exercising the protocol for LoginNewChar
    Call AddTokenDatabase(encrypted_token, decrypted_token, "MORGOLOCK2002@YAHOO.COM.AR")

    Exit Sub
Init_Failed:
    On Error GoTo 0
    Call RunTestError("UnitTesting.Init", failure)
    Exit Sub
Init_Err:
    failure = "Err.Number=" & CStr(Err.Number) & "; " & Err.Description
    Call ProgressError("UnitTesting.Init", Err.Number, Err.Description)
    Resume Init_Failed
End Sub

Public Sub shutdown()
    Call UnitClient.Disconnect
End Sub

Sub test_make_user(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer)
    UserList(UserIndex).pos.Map = Map
    UserList(UserIndex).pos.x = x
    UserList(UserIndex).pos.y = y
    Call MakeUserChar(True, 17, UserIndex, Map, x, y, 1)
End Sub

Function test_percentage() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    Debug.Assert (Porcentaje(100#, 1#) = 1)
    Debug.Assert (Porcentaje(100#, 2#) = 2)
    Debug.Assert (Porcentaje(100#, 5#) = 5)
    Debug.Assert (Porcentaje(100#, 10#) = 10)
    Debug.Assert (Porcentaje(100#, 100#) = 100)
    Dim i As Integer
    For i = 1 To 100
            Debug.Assert Porcentaje(100#, i) = i
    Next i
    For i = 1 To 1000
            Debug.Assert Porcentaje(1000#, i) = i * 10
    Next i
    Debug.Print "Porcentaje took " & sw.ElapsedMilliseconds; " ms"
    test_percentage = True
End Function

Function test_distance() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    Debug.Assert Distance(0, 0, 0, 0) = 0
    Dim i As Integer
    For i = 1 To 100
            Debug.Assert Distance(i, 0, 0, 0) = i
    Next i
    For i = 1 To 1000
           Debug.Assert Distance(i, 0, -i, 0) = i + i
    Next i
    Debug.Print "distace took " & sw.ElapsedMilliseconds; " ms"
    test_distance = True
End Function


Function test_random_number() As Boolean
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Debug.Assert RandomNumber(0, 0) = 0
    Debug.Assert RandomNumber(-1, -1) = -1
    Debug.Assert RandomNumber(1, 1) = 1
    Dim i As Integer
    Dim n As Integer
    For i = 1 To 1000
          n = RandomNumber(0, i)
          Debug.Assert n >= 0 And n <= i
    Next i
    For i = 1 To 1000
          n = RandomNumber(-i, 0)
          Debug.Assert n >= -i And n <= 0
    Next i
    
    Debug.Print "random_bumber took " & sw.ElapsedMilliseconds; " ms"
    test_random_number = True
End Function


Function test_maths() As Boolean
    test_maths = test_percentage() And test_random_number() And test_distance()
End Function

Function test_make_user_char() As Boolean
    'Create first User
    Call test_make_user(1, 1, 54, 51)
    Debug.Assert (MapData(54, 51, 1).UserIndex = 1)
    Debug.Assert (UserList(1).Char.charindex <> 0)
    'Delete first user
    Call EraseUserChar(1, False, False)
    Debug.Assert (MapData(54, 55, 1).UserIndex = 0)
    Debug.Assert (UserList(1).Char.charindex = 0)
    'Delete all NPCs5
    Dim i
    For i = 1 To UBound(NpcList)
            If NpcList(i).Char.charindex <> 0 Then
                Call EraseNPCChar(1)
            End If
    Next i
    
    'Create two users on the same map pos
    Call test_make_user(2, 1, 54, 56)
    Debug.Assert (MapData(54, 56, 1).UserIndex = 2)
    Debug.Assert (UserList(2).Char.charindex <> 0)
    
    Call test_make_user(1, 1, 50, 46)
    Debug.Assert (MapData(50, 46, 1).UserIndex = 1)
    Debug.Assert (UserList(1).Char.charindex <> 0)
    Debug.Assert (UserList(2).Char.charindex <> UserList(1).Char.charindex)
    
    'Delete user 2
    Call EraseUserChar(2, False, False)
    Debug.Assert (MapData(54, 56, 1).UserIndex = 0)
    Debug.Assert (UserList(2).Char.charindex = 0)
    'Create user 2 again
    Call test_make_user(2, 1, 54, 56)
    Debug.Assert (MapData(54, 56, 1).UserIndex = 2)
    Debug.Assert (UserList(2).Char.charindex <> 0)
    
    For i = 1 To UBound(UserList)
        If UserList(i).Char.charindex <> 0 Then
            Call EraseUserChar(i, False, True)
        End If
    Next i
    
    Call test_make_user(1, 1, 64, 66)
    Debug.Assert (MapData(64, 66, 1).UserIndex = 1)
    Debug.Assert (UserList(1).Char.charindex <> 0)
    Debug.Assert (UserList(1).Char.charindex = 1)
    
    
    Call test_make_user(1, 1, 68, 66)
    Debug.Assert (MapData(68, 66, 1).UserIndex = 1)
    Debug.Assert (UserList(1).Char.charindex <> 0)
    test_make_user_char = True
End Function



Function test_suite() As Boolean
    Call RunAllSuites
    test_suite = (FailedTests = 0)
End Function

' ============================================================
' Test Runner Core
' ============================================================

Public Sub RunTest(ByVal testName As String, ByVal testResult As Boolean)
    TotalTests = TotalTests + 1
    
    If testResult Then
        PassedTests = PassedTests + 1
    Else
        FailedTests = FailedTests + 1
        FailedTestCount = FailedTestCount + 1
        ReDim Preserve FailedTestNames(FailedTestCount)
        FailedTestNames(FailedTestCount) = testName
    End If
End Sub

Public Sub RunTestError(ByVal testName As String, ByVal errorDesc As String)
    TotalTests = TotalTests + 1
    FailedTests = FailedTests + 1
    FailedTestCount = FailedTestCount + 1
    ReDim Preserve FailedTestNames(FailedTestCount)
    FailedTestNames(FailedTestCount) = testName & " - Error: " & errorDesc
End Sub

Private Function RunSuite(ByVal suiteIndex As Integer) As Boolean
    On Error GoTo RunSuite_Err
    
    Select Case suiteIndex
        Case 1: RunSuite = Unit_Math.test_suite_math()
        Case 2: RunSuite = Unit_Bitmask.test_suite_bitmask()
        Case 3: RunSuite = Unit_StringValidation.test_suite_strings()
        Case 4: RunSuite = Unit_Pathfinding.test_suite_pathfinding()
        Case 5: RunSuite = Unit_Characters.test_suite_characters()
        Case 6: RunSuite = Unit_ElapsedTime.test_suite_elapsed_time()
        Case 7: RunSuite = Unit_Timer.test_suite_timer()
        Case 8: RunSuite = Unit_Queue.test_suite_queue()
        Case 9: RunSuite = Unit_UserNames.test_suite_usernames()
        Case 10: RunSuite = Unit_General.test_suite_general()
        Case 11: RunSuite = Unit_Factions.test_suite_factions()
        Case 12: RunSuite = Unit_Commerce.test_suite_commerce()
        Case 13: RunSuite = Unit_MapBounds.test_suite_mapbounds()
        Case 14: RunSuite = Unit_StringBuilder.test_suite_stringbuilder()
        Case 15: RunSuite = Unit_CombatMath.test_suite_combatmath()
        Case 16: RunSuite = Unit_Crafting.test_suite_crafting()
        Case 17: RunSuite = Unit_GameStatus.test_suite_gamestatus()
        Case 18: RunSuite = Unit_IniManager.test_suite_ini_manager()
        Case 19: RunSuite = Unit_WorldTime.test_suite_world_time()
        Case 20: RunSuite = Unit_TimeFormat.test_suite_time_format()
        Case 21: RunSuite = Unit_StringUtils.test_suite_string_utils()
        Case 22: RunSuite = Unit_Spawn.test_suite_spawn()
        Case 23: RunSuite = Unit_CommerceExt.test_suite_commerce_ext()
        Case 24: RunSuite = Unit_ScoreBoard.test_suite_scoreboard()
        Case 25: RunSuite = Unit_Uptime.test_suite_uptime()
        Case 26: RunSuite = Unit_CryptoConvert.test_suite_crypto_convert()
        Case 27: RunSuite = Unit_IntBitmask.test_suite_int_bitmask()
        Case 28: RunSuite = Unit_BitShift.test_suite_bit_shift()
        Case 29: RunSuite = Unit_TimerStruct.test_suite_timer_struct()
        Case 30: RunSuite = Unit_CircularLog.test_suite_circular_log()
        Case 31: RunSuite = Unit_StringValid.test_suite_string_valid()
        Case 32: RunSuite = Unit_SumarTiempo.test_suite_sumar_tiempo()
        Case 33: RunSuite = Unit_Network_Aurora.test_suite_network_aurora()
        Case 34: RunSuite = Unit_Weather.test_suite_weather()
        Case 35: RunSuite = test_suite_remort_persistence()
        Case 36: RunSuite = test_suite_remort_capability_state()
        Case 37: RunSuite = Unit_NpcCrossMapPursuit.test_suite_npc_cross_map_pursuit()
        Case 38: RunSuite = Unit_AdminVisibility.test_suite_admin_visibility()
        Case 39: RunSuite = modMagicLightOrbs.test_suite_light_orbs()
        Case 40: RunSuite = test_suite_taming()
        Case 41: RunSuite = Unit_Traps.test_suite_traps()
        Case 42: RunSuite = Unit_InventoryDestruction.test_suite_inventory_destruction()
        Case 43: RunSuite = Unit_NpcBow.test_suite_npc_bow()
        Case 44: RunSuite = Unit_AdjacentPlayers.TestSuite()
        Case Else
            Call RunTestError("suite registration " & CStr(SuiteIndex), "No registered suite")
            RunSuite = False
    End Select
    Exit Function
    
RunSuite_Err:
    Call ProgressError("suite " & CStr(suiteIndex), Err.Number, Err.Description)
    Call RunTestError("suite execution", Err.Description)
    RunSuite = False
End Function

Public Sub RunAllSuites()
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Dim i As Integer, failuresBefore As Long
    For i = 1 To SUITE_COUNT
        Call Progress("BEFORE suite " & CStr(i))
        failuresBefore = FailedTests
        If Not RunSuite(i) Then
            If FailedTests = failuresBefore Then Call RunTestError("suite " & CStr(i), "Suite returned failure")
        End If
        Call Progress("AFTER suite " & CStr(i) & "; new failures=" & CStr(FailedTests - failuresBefore))
    Next i
    
    TotalElapsed = sw.ElapsedMilliseconds
End Sub

Public Sub WriteResultsToFile(ByVal filePath As String)
    On Error GoTo WriteResultsToFile_Err
    Dim f As Integer, opened As Boolean, temporaryPath As String
    temporaryPath = filePath & ".tmp"
    f = FreeFile
    Open temporaryPath For Output As #f
    opened = True
    Print #f, "=== AO20 TEST REPORT ==="
    Print #f, "Total: " & TotalTests & " | Passed: " & PassedTests & " | Failed: " & FailedTests
    
    If FailedTestCount > 0 Then
        Print #f, "Failed tests:"
        Dim i As Long
        For i = 1 To FailedTestCount
            Print #f, "  - " & FailedTestNames(i)
        Next i
    End If
    
    Print #f, "Total time: " & Format$(TotalElapsed, "0.00") & " ms"
    
    If FailedTests = 0 Then
        Print #f, "RESULT: PASS"
    Else
        Print #f, "RESULT: FAIL"
    End If
    Close #f
    opened = False
    ' Publish only the closed final report, never a partial result or checkpoint.
    If LenB(Dir$(filePath)) <> 0 Then Kill filePath
    Name temporaryPath As filePath
    Exit Sub
WriteResultsToFile_Err:
    Call ProgressError("WriteResultsToFile", Err.Number, Err.Description)
    If opened Then Close #f
End Sub

Public Function GetFailedTests() As Long
    GetFailedTests = FailedTests
End Function

Public Function TestFailureReport() As Boolean
    Dim savedTotal As Long, savedPassed As Long, savedFailed As Long, savedCount As Long
    Dim savedNames() As String, f As Integer, opened As Boolean
    Dim path As String, contents As String, line As String, caught As String
    savedTotal = TotalTests: savedPassed = PassedTests
    savedFailed = FailedTests: savedCount = FailedTestCount
    savedNames = FailedTestNames
    On Error GoTo CaughtError
    TotalTests = 32767: PassedTests = 32767: FailedTests = 32767
    FailedTestCount = 0: ReDim FailedTestNames(0)
    Err.Raise 6, , "injected runner failure"
CaughtError:
    caught = CStr(Err.Number) & ": " & Err.Description
    Resume RecordFailure
RecordFailure:
    On Error GoTo Failed
    Call RunTestError("runner regression", caught)
    If TotalTests <> 32768 Or FailedTests <> 32768 Then GoTo Restore
    Call RunTest("runner success counter", True)
    If PassedTests <> 32768 Or TotalTests <> 32769 Then GoTo Restore
    path = App.Path & "\test_runner_regression.txt"
    Call WriteResultsToFile(path)
    f = FreeFile
    Open path For Input As #f
    opened = True
    Do While Not EOF(f)
        Line Input #f, line
        contents = contents & line & vbLf
    Loop
    TestFailureReport = InStr(contents, "6: injected runner failure") > 0 And InStr(contents, "RESULT: FAIL") > 0
Restore:
    On Error Resume Next
    If opened Then Close #f
    If LenB(path) <> 0 Then
        If LenB(Dir$(path)) <> 0 Then Kill path
    End If
    TotalTests = savedTotal: PassedTests = savedPassed
    FailedTests = savedFailed: FailedTestCount = savedCount
    FailedTestNames = savedNames
    Exit Function
Failed:
    TestFailureReport = False
    Resume Restore
End Function

Private Function test_suite_remort_persistence() As Boolean
    Call RunTest("remort_count migration default", test_remort_count_migration_default())
    Call RunTest("remort_count load validation", test_remort_count_load_validation())
    Call RunTest("remort_count normal save SQL", test_remort_count_normal_save_sql())
    Call RunTest("last_logout save SQL", test_last_logout_save_sql())
    Call RunTest("last_logout load SQL", InStr(1, QUERY_LOAD_MAINPJ, "last_logout", vbTextCompare) > 0)
    Call RunTest("last_logout UTC timestamp", FormatPreviousLogoutUtc(1792620900#) = "2026-10-21 22:15:00 UTC")
    Call RunTest("last_logout SQLite text timestamp", FormatPreviousLogoutUtc("1792620900") = "2026-10-21 22:15:00 UTC")
    Call RunTest("last_logout beyond 2038", FormatPreviousLogoutUtc(2147483648#) = "2038-01-19 03:14:08 UTC")
    Call RunTest("last_logout missing omitted", FormatPreviousLogoutUtc(Null) = vbNullString And FormatPreviousLogoutUtc(Empty) = vbNullString)
    Call RunTest("last_logout nonpositive omitted", FormatPreviousLogoutUtc(0) = vbNullString And FormatPreviousLogoutUtc(-1) = vbNullString)
    Call RunTest("last_logout invalid omitted", FormatPreviousLogoutUtc("invalid") = vbNullString And FormatPreviousLogoutUtc(1.5) = vbNullString)
    Call RunTest("last_logout overflow omitted", FormatPreviousLogoutUtc(253402300800#) = vbNullString)
    test_suite_remort_persistence = True
End Function

Private Function test_remort_count_migration_default() As Boolean
    On Error GoTo TestError

    Dim RS As ADODB.Recordset
    Set RS = Query("PRAGMA table_info(user)")
    If RS Is Nothing Then Exit Function

    Do While Not RS.EOF
        If LCase$(CStr(RS.Fields("name").value)) = "remort_count" Then
            Dim defaultValue As String
            defaultValue = Replace$(CStr(RS.Fields("dflt_value").value), "(", vbNullString)
            defaultValue = Replace$(defaultValue, ")", vbNullString)
            test_remort_count_migration_default = _
                    CLng(RS.Fields("notnull").value) = 1 And defaultValue = "0"
            Exit Function
        End If
        RS.MoveNext
    Loop
    Exit Function

TestError:
    test_remort_count_migration_default = False
End Function

Private Function test_remort_count_load_validation() As Boolean
    On Error GoTo TestError

    If NormalizeRemortCount(0) <> 0 Then Exit Function
    If NormalizeRemortCount(3) <> 3 Then Exit Function
    If NormalizeRemortCount(-1) <> 0 Then Exit Function
    If NormalizeRemortCount(2147483648#) <> 0 Then Exit Function

    test_remort_count_load_validation = True
    Exit Function

TestError:
    test_remort_count_load_validation = False
End Function

Private Function test_remort_count_normal_save_sql() As Boolean
    test_remort_count_normal_save_sql = _
            InStr(1, QUERY_LOAD_MAINPJ, "remort_count", vbTextCompare) > 0 And _
            InStr(1, QUERY_UPDATE_MAINPJ, "remort_count = ?", vbTextCompare) > 0
End Function

Private Function test_last_logout_save_sql() As Boolean
    test_last_logout_save_sql = _
            InStr(1, QUERY_UPDATE_MAINPJ_LOGOUT, "last_logout = strftime('%s','now')", vbTextCompare) > 0 And _
            InStr(1, QUERY_UPDATE_MAINPJ_LOGOUT, "remort_count = ?, last_logout", vbTextCompare) > 0 And _
            InStr(1, QUERY_UPDATE_MAINPJ, "last_logout", vbTextCompare) = 0
End Function

Private Function test_suite_remort_capability_state() As Boolean
    Call RunTest("remort eligibility and priority", test_remort_eligibility_and_priority())
    Call RunTest("remort eligibility is read-only", test_remort_eligibility_is_read_only())
    Call RunTest("remort capability bit and reset", test_remort_capability_bit_and_reset())
    Call RunTest("remort packet is appended", CInt(ServerPacketID.eRemortState) = CInt(ServerPacketID.eShowPickUpObj) + 1)
    Call RunTest("remort operation packet IDs are appended", CInt(ServerPacketID.eRemortResult) = CInt(ServerPacketID.eRemortState) + 1 And CInt(ClientPacketID.eRequestRemort) = CInt(ClientPacketID.eHooClientCapabilities) + 1)
    Call RunTest("targeted spell packet IDs are appended", CInt(ServerPacketID.eHooTargetedSpellCastResult) = CInt(ServerPacketID.eRemortResult) + 1 And CInt(ClientPacketID.eHooTargetedSpellCast) = CInt(ClientPacketID.eRequestRemort) + 1)
    Call RunTest("house door packet IDs and legacy key IDs", test_house_door_packet_ids_and_capability())
    Call RunTest("house door actions validate before mutation", test_house_door_actions())
    Call RunTest("targeted spell capability and feature gate", test_targeted_spell_capability())
    Call RunTest("targeted spell retry interval is wrap safe", test_targeted_spell_retry_interval())
    Call RunTest("targeted spell NPC resolution rejects stale mappings", test_targeted_spell_target_resolution())
    Call RunTest("targeted spell eligibility rejects unsupported spells", test_targeted_spell_eligibility())
    Call RunTest("targeted spell range matches legacy vision boundaries", test_targeted_spell_range_boundaries())
    Call RunTest("targeted spell early rejections are read-only", test_targeted_spell_early_rejections())
    Call RunTest("remort equipment and count eligibility", test_remort_equipment_and_count_eligibility())
    Call RunTest("remort live reset and preservation", test_remort_live_reset_and_preservation())
    test_suite_remort_capability_state = True
End Function

Private Function test_house_door_packet_ids_and_capability() As Boolean
    On Error GoTo TestError
    Dim OriginalFeatureEnabled As Boolean
    OriginalFeatureEnabled = IsFeatureEnabled(HOO_FEATURE_HOUSE_DOOR_ACTIONS_V1)
    Call SetFeatureToggle(HOO_FEATURE_HOUSE_DOOR_ACTIONS_V1, True)
    If AcceptedHooCapabilityMask(HOO_CAP_PROTOCOL_VERSION, HOO_CAP_HOUSE_DOOR_ACTIONS_V1) <> HOO_CAP_HOUSE_DOOR_ACTIONS_V1 Then GoTo TestDone
    Call SetFeatureToggle(HOO_FEATURE_HOUSE_DOOR_ACTIONS_V1, False)
    If AcceptedHooCapabilityMask(HOO_CAP_PROTOCOL_VERSION, HOO_CAP_HOUSE_DOOR_ACTIONS_V1) <> 0 Then GoTo TestDone
    If HOO_CAP_HOUSE_DOOR_ACTIONS_V1 <> &H8& Then GoTo TestDone
    If CInt(ClientPacketID.eUseKey) <> 209 Or CInt(ServerPacketID.eUpdateUserKey) <> 159 Then GoTo TestDone
    If CInt(ClientPacketID.eHooHouseDoorAction) <> CInt(ClientPacketID.eHooTargetedSpellCast) + 1 Then GoTo TestDone
    If CInt(ServerPacketID.eHooHouseDoorActionResult) <> CInt(ServerPacketID.eHooTargetedSpellCastResult) + 1 Then GoTo TestDone
    test_house_door_packet_ids_and_capability = True
TestDone:
    Call SetFeatureToggle(HOO_FEATURE_HOUSE_DOOR_ACTIONS_V1, OriginalFeatureEnabled)
    Exit Function
TestError:
    test_house_door_packet_ids_and_capability = False
    Resume TestDone
End Function

Private Function test_house_door_actions() As Boolean
    On Error GoTo TestError
    Dim OriginalObj(1 To 4) As t_ObjData
    Dim i As Integer
    For i = 1 To 4
        OriginalObj(i) = ObjData(i)
    Next i
    Dim OriginalTileObj As t_Obj
    Dim OriginalPairObj As t_Obj
    OriginalTileObj = MapData(50, 50, 1).ObjInfo
    OriginalPairObj = MapData(47, 51, 1).ObjInfo
    Dim OriginalBlocked As Byte
    Dim OriginalBlockedSouth As Byte
    OriginalBlocked = MapData(50, 50, 1).Blocked
    OriginalBlockedSouth = MapData(50, 51, 1).Blocked
    Dim OriginalMap As Integer
    Dim OriginalX As Byte
    Dim OriginalY As Byte
    Dim OriginalTimer As Long
    OriginalMap = UserList(1).pos.Map
    OriginalX = UserList(1).pos.x
    OriginalY = UserList(1).pos.y
    OriginalTimer = UserList(1).Counters.TimerUsar
    Dim OriginalKeys(1 To MAXKEYS) As Integer
    For i = 1 To MAXKEYS
        OriginalKeys(i) = UserList(1).Keys(i)
        UserList(1).Keys(i) = 0
    Next i

    For i = 1 To 3
        ObjData(i).OBJType = e_OBJType.otDoors
        ObjData(i).clave = 410
        ObjData(i).IndexAbierta = 1
        ObjData(i).IndexCerrada = 2
        ObjData(i).IndexCerradaLlave = 3
        ObjData(i).Subtipo = 4
        ObjData(i).GrhIndex = 1
    Next i
    ObjData(1).Cerrada = 0: ObjData(1).Llave = 0
    ObjData(2).Cerrada = 1: ObjData(2).Llave = 0
    ObjData(3).Cerrada = 1: ObjData(3).Llave = 1
    ObjData(4).OBJType = e_OBJType.otKeys: ObjData(4).clave = 410
    UserList(1).pos.Map = 1: UserList(1).pos.x = 50: UserList(1).pos.y = 50
    MapData(50, 50, 1).ObjInfo.ObjIndex = 2
    MapData(50, 50, 1).ObjInfo.amount = 1

    If ExecuteHooHouseDoorAction(1, eHooHouseDoorAction_Open, 50, 50) <> eHooHouseDoorActionResult_Success Then GoTo TestDone
    If MapData(50, 50, 1).ObjInfo.ObjIndex <> 1 Then GoTo TestDone
    If ExecuteHooHouseDoorAction(1, eHooHouseDoorAction_Close, 50, 50) <> eHooHouseDoorActionResult_Success Then GoTo TestDone
    If MapData(50, 50, 1).ObjInfo.ObjIndex <> 2 Then GoTo TestDone
    If ExecuteHooHouseDoorAction(1, 250, 50, 50) <> eHooHouseDoorActionResult_InvalidAction Then GoTo TestDone
    If MapData(50, 50, 1).ObjInfo.ObjIndex <> 2 Then GoTo TestDone
    UserList(1).pos.x = 1: UserList(1).pos.y = 1
    If ExecuteHooHouseDoorAction(1, eHooHouseDoorAction_Open, 50, 50) <> eHooHouseDoorActionResult_TooFarAway Then GoTo TestDone
    UserList(1).pos.x = 50: UserList(1).pos.y = 50
    MapData(50, 50, 1).ObjInfo.ObjIndex = 4
    If ExecuteHooHouseDoorAction(1, eHooHouseDoorAction_Open, 50, 50) <> eHooHouseDoorActionResult_NotDoor Then GoTo TestDone
    MapData(50, 50, 1).ObjInfo.ObjIndex = 2
    If ExecuteHooHouseDoorAction(1, eHooHouseDoorAction_Lock, 50, 50) <> eHooHouseDoorActionResult_NoAccess Then GoTo TestDone
    UserList(1).Keys(1) = 4
    UserList(1).Counters.TimerUsar = 0
    If ExecuteHooHouseDoorAction(1, eHooHouseDoorAction_Lock, 50, 50) <> eHooHouseDoorActionResult_Success Then GoTo TestDone
    If MapData(50, 50, 1).ObjInfo.ObjIndex <> 3 Then GoTo TestDone
    UserList(1).Counters.TimerUsar = GetTickCountRaw()
    If ExecuteHooHouseDoorAction(1, eHooHouseDoorAction_Unlock, 50, 50) <> eHooHouseDoorActionResult_Cooldown Then GoTo TestDone
    If MapData(50, 50, 1).ObjInfo.ObjIndex <> 3 Then GoTo TestDone
    UserList(1).Counters.TimerUsar = 0
    If ExecuteHooHouseDoorAction(1, eHooHouseDoorAction_Unlock, 50, 50) <> eHooHouseDoorActionResult_Success Then GoTo TestDone
    If MapData(50, 50, 1).ObjInfo.ObjIndex <> 2 Then GoTo TestDone
    MapData(50, 50, 1).ObjInfo.ObjIndex = 3
    UserList(1).Counters.TimerUsar = 0
    If ExecuteHooHouseDoorAction(1, eHooHouseDoorAction_UnlockAndOpen, 50, 50) <> eHooHouseDoorActionResult_Success Then GoTo TestDone
    If MapData(50, 50, 1).ObjInfo.ObjIndex <> 1 Then GoTo TestDone
    MapData(50, 50, 1).ObjInfo.ObjIndex = 2
    ObjData(1).Subtipo = 1
    MapData(47, 51, 1).ObjInfo.ObjIndex = 2
    MapData(47, 51, 1).ObjInfo.amount = 1
    If ExecuteHooHouseDoorAction(1, eHooHouseDoorAction_Open, 50, 50) <> eHooHouseDoorActionResult_Success Then GoTo TestDone
    If MapData(50, 50, 1).ObjInfo.ObjIndex <> 1 Or MapData(47, 51, 1).ObjInfo.ObjIndex <> 1 Then GoTo TestDone
    MapData(50, 50, 1).ObjInfo.ObjIndex = 2
    MapData(47, 51, 1).ObjInfo.ObjIndex = 0
    If ExecuteHooHouseDoorAction(1, eHooHouseDoorAction_Open, 50, 50) <> eHooHouseDoorActionResult_Unavailable Then GoTo TestDone
    If MapData(50, 50, 1).ObjInfo.ObjIndex <> 2 Then GoTo TestDone
    test_house_door_actions = True
TestDone:
    For i = 1 To 4
        ObjData(i) = OriginalObj(i)
        UserList(1).Keys(i) = OriginalKeys(i)
    Next i
    For i = 5 To MAXKEYS
        UserList(1).Keys(i) = OriginalKeys(i)
    Next i
    MapData(50, 50, 1).ObjInfo = OriginalTileObj
    MapData(47, 51, 1).ObjInfo = OriginalPairObj
    MapData(50, 50, 1).Blocked = OriginalBlocked
    MapData(50, 51, 1).Blocked = OriginalBlockedSouth
    UserList(1).pos.Map = OriginalMap
    UserList(1).pos.x = OriginalX
    UserList(1).pos.y = OriginalY
    UserList(1).Counters.TimerUsar = OriginalTimer
    Exit Function
TestError:
    test_house_door_actions = False
    Resume TestDone
End Function

Private Function test_targeted_spell_range_boundaries() As Boolean
    On Error GoTo TestError
    Dim OriginalUserX As Byte
    Dim OriginalUserY As Byte
    Dim OriginalNpcX As Byte
    Dim OriginalNpcY As Byte
    OriginalUserX = UserList(1).pos.x
    OriginalUserY = UserList(1).pos.y
    OriginalNpcX = NpcList(1).pos.x
    OriginalNpcY = NpcList(1).pos.y

    UserList(1).pos.x = 50
    UserList(1).pos.y = 50

    NpcList(1).pos.x = 61: NpcList(1).pos.y = 50
    If Not IsHooTargetedSpellTargetInRange(1, 1) Then GoTo TestDone
    NpcList(1).pos.x = 39
    If Not IsHooTargetedSpellTargetInRange(1, 1) Then GoTo TestDone
    NpcList(1).pos.x = 62
    If IsHooTargetedSpellTargetInRange(1, 1) Then GoTo TestDone
    NpcList(1).pos.x = 38
    If IsHooTargetedSpellTargetInRange(1, 1) Then GoTo TestDone

    NpcList(1).pos.x = 50: NpcList(1).pos.y = 59
    If Not IsHooTargetedSpellTargetInRange(1, 1) Then GoTo TestDone
    NpcList(1).pos.y = 41
    If Not IsHooTargetedSpellTargetInRange(1, 1) Then GoTo TestDone
    NpcList(1).pos.y = 60
    If IsHooTargetedSpellTargetInRange(1, 1) Then GoTo TestDone
    NpcList(1).pos.y = 40
    If IsHooTargetedSpellTargetInRange(1, 1) Then GoTo TestDone

    NpcList(1).pos.x = 61: NpcList(1).pos.y = 59
    If Not IsHooTargetedSpellTargetInRange(1, 1) Then GoTo TestDone
    NpcList(1).pos.x = 62
    If IsHooTargetedSpellTargetInRange(1, 1) Then GoTo TestDone
    NpcList(1).pos.x = 61: NpcList(1).pos.y = 60
    If IsHooTargetedSpellTargetInRange(1, 1) Then GoTo TestDone

    test_targeted_spell_range_boundaries = True
TestDone:
    UserList(1).pos.x = OriginalUserX
    UserList(1).pos.y = OriginalUserY
    NpcList(1).pos.x = OriginalNpcX
    NpcList(1).pos.y = OriginalNpcY
    Exit Function
TestError:
    test_targeted_spell_range_boundaries = False
    Resume TestDone
End Function

Private Function test_targeted_spell_early_rejections() As Boolean
    On Error GoTo TestError
    Dim OriginalMapping As Integer
    Dim OriginalActive As Boolean
    Dim OriginalHp As Integer
    Dim OriginalCharIndex As Integer
    Dim OriginalNpcMap As Integer
    Dim OriginalNpcX As Byte
    Dim OriginalNpcY As Byte
    Dim OriginalUserMap As Integer
    Dim OriginalUserX As Byte
    Dim OriginalUserY As Byte
    Dim OriginalSpellIndex As Integer
    Dim OriginalAuto As Byte
    Dim OriginalAreaRadio As Long
    Dim OriginalAreaAfecta As Integer
    Dim OriginalTarget As e_TargetType
    Dim OriginalEffect As e_TargetEffectType
    Dim OriginalType As e_TipoHechizo
    Dim OriginalBowTimer As Long
    Dim OriginalHitMagicTimer As Long
    Dim OriginalMagicTimer As Long
    OriginalMapping = CharList(1)
    OriginalActive = NpcList(1).flags.NPCActive
    OriginalHp = NpcList(1).Stats.MinHp
    OriginalCharIndex = NpcList(1).Char.charindex
    OriginalNpcMap = NpcList(1).pos.Map
    OriginalNpcX = NpcList(1).pos.x
    OriginalNpcY = NpcList(1).pos.y
    OriginalUserMap = UserList(1).pos.Map
    OriginalUserX = UserList(1).pos.x
    OriginalUserY = UserList(1).pos.y
    OriginalSpellIndex = UserList(1).Stats.UserHechizos(1)
    OriginalAuto = Hechizos(1).AutoLanzar
    OriginalAreaRadio = Hechizos(1).AreaRadio
    OriginalAreaAfecta = Hechizos(1).AreaAfecta
    OriginalTarget = Hechizos(1).Target
    OriginalEffect = Hechizos(1).TargetEffectType
    OriginalType = Hechizos(1).Tipo
    OriginalBowTimer = UserList(1).Counters.TimerPuedeUsarArco
    OriginalHitMagicTimer = UserList(1).Counters.TimerGolpeMagia
    OriginalMagicTimer = UserList(1).Counters.TimerLanzarSpell

    CharList(1) = 1
    NpcList(1).flags.NPCActive = True
    NpcList(1).Stats.MinHp = 10
    NpcList(1).Char.charindex = 1
    NpcList(1).pos.Map = 1
    NpcList(1).pos.x = 20
    NpcList(1).pos.y = 20
    UserList(1).pos.Map = 1
    UserList(1).pos.x = 20
    UserList(1).pos.y = 20
    UserList(1).Stats.UserHechizos(1) = 1
    Hechizos(1).AutoLanzar = 0
    Hechizos(1).AreaRadio = 0
    Hechizos(1).AreaAfecta = 0
    Hechizos(1).Target = e_TargetType.uNPC
    Hechizos(1).TargetEffectType = e_TargetEffectType.eNegative
    Hechizos(1).Tipo = e_TipoHechizo.uPropiedades

    Dim RetryAfterMs As Long
    If ExecuteHooTargetedSpellCast(1, 1, 0, RetryAfterMs) <> eHooTargetedSpellCastResult_InvalidTarget Then GoTo TestDone
    If ExecuteHooTargetedSpellCast(1, 0, 1, RetryAfterMs) <> eHooTargetedSpellCastResult_InvalidSpell Then GoTo TestDone
    NpcList(1).pos.Map = 2
    If ExecuteHooTargetedSpellCast(1, 1, 1, RetryAfterMs) <> eHooTargetedSpellCastResult_InvalidTarget Then GoTo TestDone
    NpcList(1).pos.Map = 1
    NpcList(1).pos.x = 100
    If ExecuteHooTargetedSpellCast(1, 1, 1, RetryAfterMs) <> eHooTargetedSpellCastResult_OutOfRange Then GoTo TestDone
    NpcList(1).pos.x = 20
    Dim NowRaw As Long
    NowRaw = GetTickCountRaw()
    UserList(1).Counters.TimerPuedeUsarArco = NowRaw
    UserList(1).Counters.TimerGolpeMagia = NowRaw
    UserList(1).Counters.TimerLanzarSpell = NowRaw
    If ExecuteHooTargetedSpellCast(1, 1, 1, RetryAfterMs) <> eHooTargetedSpellCastResult_RateLimited Then GoTo TestDone
    If RetryAfterMs <= 0 Then GoTo TestDone
    If UserList(1).Counters.TimerPuedeUsarArco <> NowRaw Then GoTo TestDone
    If UserList(1).Counters.TimerGolpeMagia <> NowRaw Then GoTo TestDone
    If UserList(1).Counters.TimerLanzarSpell <> NowRaw Then GoTo TestDone
    test_targeted_spell_early_rejections = True
TestDone:
    CharList(1) = OriginalMapping
    NpcList(1).flags.NPCActive = OriginalActive
    NpcList(1).Stats.MinHp = OriginalHp
    NpcList(1).Char.charindex = OriginalCharIndex
    NpcList(1).pos.Map = OriginalNpcMap
    NpcList(1).pos.x = OriginalNpcX
    NpcList(1).pos.y = OriginalNpcY
    UserList(1).pos.Map = OriginalUserMap
    UserList(1).pos.x = OriginalUserX
    UserList(1).pos.y = OriginalUserY
    UserList(1).Stats.UserHechizos(1) = OriginalSpellIndex
    Hechizos(1).AutoLanzar = OriginalAuto
    Hechizos(1).AreaRadio = OriginalAreaRadio
    Hechizos(1).AreaAfecta = OriginalAreaAfecta
    Hechizos(1).Target = OriginalTarget
    Hechizos(1).TargetEffectType = OriginalEffect
    Hechizos(1).Tipo = OriginalType
    UserList(1).Counters.TimerPuedeUsarArco = OriginalBowTimer
    UserList(1).Counters.TimerGolpeMagia = OriginalHitMagicTimer
    UserList(1).Counters.TimerLanzarSpell = OriginalMagicTimer
    Exit Function
TestError:
    test_targeted_spell_early_rejections = False
    Resume TestDone
End Function

Private Function test_targeted_spell_target_resolution() As Boolean
    On Error GoTo TestError
    Dim OriginalMapping As Integer
    Dim OriginalActive As Boolean
    Dim OriginalHp As Integer
    Dim OriginalCharIndex As Integer
    OriginalMapping = CharList(1)
    OriginalActive = NpcList(1).flags.NPCActive
    OriginalHp = NpcList(1).Stats.MinHp
    OriginalCharIndex = NpcList(1).Char.charindex

    CharList(1) = 1
    NpcList(1).flags.NPCActive = True
    NpcList(1).Stats.MinHp = 10
    NpcList(1).Char.charindex = 1
    If ResolveHooTargetedSpellNpc(1) <> 1 Then GoTo TestDone
    NpcList(1).Char.charindex = 2
    If ResolveHooTargetedSpellNpc(1) <> 0 Then GoTo TestDone
    NpcList(1).Char.charindex = 1
    NpcList(1).Stats.MinHp = 0
    If ResolveHooTargetedSpellNpc(1) <> 0 Then GoTo TestDone
    test_targeted_spell_target_resolution = True
TestDone:
    CharList(1) = OriginalMapping
    NpcList(1).flags.NPCActive = OriginalActive
    NpcList(1).Stats.MinHp = OriginalHp
    NpcList(1).Char.charindex = OriginalCharIndex
    Exit Function
TestError:
    test_targeted_spell_target_resolution = False
    Resume TestDone
End Function

Private Function test_targeted_spell_eligibility() As Boolean
    On Error GoTo TestError
    Dim OriginalSpellIndex As Integer
    Dim OriginalAuto As Byte
    Dim OriginalAreaRadio As Long
    Dim OriginalAreaAfecta As Integer
    Dim OriginalTarget As e_TargetType
    Dim OriginalEffect As e_TargetEffectType
    Dim OriginalType As e_TipoHechizo
    OriginalSpellIndex = UserList(1).Stats.UserHechizos(1)
    OriginalAuto = Hechizos(1).AutoLanzar
    OriginalAreaRadio = Hechizos(1).AreaRadio
    OriginalAreaAfecta = Hechizos(1).AreaAfecta
    OriginalTarget = Hechizos(1).Target
    OriginalEffect = Hechizos(1).TargetEffectType
    OriginalType = Hechizos(1).Tipo

    UserList(1).Stats.UserHechizos(1) = 1
    Hechizos(1).AutoLanzar = 0
    Hechizos(1).AreaRadio = 0
    Hechizos(1).AreaAfecta = 0
    Hechizos(1).Target = e_TargetType.uNPC
    Hechizos(1).TargetEffectType = e_TargetEffectType.eNegative
    Hechizos(1).Tipo = e_TipoHechizo.uPropiedades
    If Not IsHooTargetedSpellEligible(1, 1) Then GoTo TestDone
    Hechizos(1).AutoLanzar = 1
    If IsHooTargetedSpellEligible(1, 1) Then GoTo TestDone
    Hechizos(1).AutoLanzar = 0
    Hechizos(1).AreaRadio = 1
    If IsHooTargetedSpellEligible(1, 1) Then GoTo TestDone
    Hechizos(1).AreaRadio = 0
    Hechizos(1).Target = e_TargetType.uUsuarios
    If IsHooTargetedSpellEligible(1, 1) Then GoTo TestDone
    Hechizos(1).Target = e_TargetType.uNPC
    Hechizos(1).TargetEffectType = e_TargetEffectType.ePositive
    If IsHooTargetedSpellEligible(1, 1) Then GoTo TestDone
    test_targeted_spell_eligibility = True
TestDone:
    UserList(1).Stats.UserHechizos(1) = OriginalSpellIndex
    Hechizos(1).AutoLanzar = OriginalAuto
    Hechizos(1).AreaRadio = OriginalAreaRadio
    Hechizos(1).AreaAfecta = OriginalAreaAfecta
    Hechizos(1).Target = OriginalTarget
    Hechizos(1).TargetEffectType = OriginalEffect
    Hechizos(1).Tipo = OriginalType
    Exit Function
TestError:
    test_targeted_spell_eligibility = False
    Resume TestDone
End Function

Private Function test_targeted_spell_capability() As Boolean
    On Error GoTo TestError
    Dim OriginalFeatureEnabled As Boolean
    OriginalFeatureEnabled = IsFeatureEnabled(HOO_FEATURE_TARGETED_SPELL_CAST_V1)
    Call SetFeatureToggle(HOO_FEATURE_TARGETED_SPELL_CAST_V1, True)
    If AcceptedHooCapabilityMask(HOO_CAP_PROTOCOL_VERSION, HOO_CAP_TARGETED_SPELL_CAST_V1) <> HOO_CAP_TARGETED_SPELL_CAST_V1 Then GoTo TestDone
    Call SetFeatureToggle(HOO_FEATURE_TARGETED_SPELL_CAST_V1, False)
    If AcceptedHooCapabilityMask(HOO_CAP_PROTOCOL_VERSION, HOO_CAP_TARGETED_SPELL_CAST_V1) <> 0 Then GoTo TestDone
    test_targeted_spell_capability = (HOO_CAP_TARGETED_SPELL_CAST_V1 = &H4&)
TestDone:
    Call SetFeatureToggle(HOO_FEATURE_TARGETED_SPELL_CAST_V1, OriginalFeatureEnabled)
    Exit Function
TestError:
    test_targeted_spell_capability = False
    Resume TestDone
End Function

Private Function test_targeted_spell_retry_interval() As Boolean
    test_targeted_spell_retry_interval = _
        IntervalRemainingMs(1000, 500, 1250) = 250 And _
        IntervalRemainingMs(1000, 500, 1500) = 0 And _
        IntervalRemainingMs(2147483600, 200, -2147483596) = 100
End Function

Private Function test_remort_eligibility_and_priority() As Boolean
    On Error GoTo TestError
    Dim OriginalLevel As Byte
    Dim OriginalDead As Byte
    Dim OriginalInParty As Boolean
    Dim OriginalQuests(1 To MAXUSERQUESTS) As Integer
    Dim QuestSlot As Integer
    Dim OriginalRemortCount As Long
    Dim OriginalEquipped(1 To MAX_INVENTORY_SLOTS) As Byte
    Dim InventorySlot As Integer
    OriginalLevel = UserList(1).Stats.ELV
    OriginalDead = UserList(1).flags.Muerto
    OriginalInParty = UserList(1).Grupo.EnGrupo
    OriginalRemortCount = UserList(1).Stats.RemortCount
    UserList(1).Stats.RemortCount = 0
    For InventorySlot = 1 To MAX_INVENTORY_SLOTS
        OriginalEquipped(InventorySlot) = UserList(1).invent.Object(InventorySlot).Equipped
        UserList(1).invent.Object(InventorySlot).Equipped = 0
    Next InventorySlot
    For QuestSlot = 1 To MAXUSERQUESTS
        OriginalQuests(QuestSlot) = UserList(1).QuestStats.Quests(QuestSlot).QuestIndex
        UserList(1).QuestStats.Quests(QuestSlot).QuestIndex = 0
    Next QuestSlot

    UserList(1).Stats.ELV = STAT_MAXELV - 1
    UserList(1).flags.Muerto = 1
    If GetRemortEligibility(1) <> eRemortEligibility_BelowRequiredLevel Then GoTo TestDone

    UserList(1).Stats.ELV = STAT_MAXELV
    If GetRemortEligibility(1) <> eRemortEligibility_Dead Then GoTo TestDone

    UserList(1).flags.Muerto = 0
    UserList(1).QuestStats.Quests(1).QuestIndex = 1
    UserList(1).Grupo.EnGrupo = True
    If GetRemortEligibility(1) <> eRemortEligibility_ActiveQuest Then GoTo TestDone

    UserList(1).QuestStats.Quests(1).QuestIndex = 0
    If GetRemortEligibility(1) <> eRemortEligibility_InParty Then GoTo TestDone

    UserList(1).Grupo.EnGrupo = False
    If GetRemortEligibility(1) <> eRemortEligibility_Eligible Then GoTo TestDone
    test_remort_eligibility_and_priority = True

TestDone:
    UserList(1).Stats.ELV = OriginalLevel
    UserList(1).flags.Muerto = OriginalDead
    UserList(1).Grupo.EnGrupo = OriginalInParty
    UserList(1).Stats.RemortCount = OriginalRemortCount
    For InventorySlot = 1 To MAX_INVENTORY_SLOTS
        UserList(1).invent.Object(InventorySlot).Equipped = OriginalEquipped(InventorySlot)
    Next InventorySlot
    For QuestSlot = 1 To MAXUSERQUESTS
        UserList(1).QuestStats.Quests(QuestSlot).QuestIndex = OriginalQuests(QuestSlot)
    Next QuestSlot
    Exit Function
TestError:
    test_remort_eligibility_and_priority = False
    Resume TestDone
End Function

Private Function test_remort_eligibility_is_read_only() As Boolean
    On Error GoTo TestError
    Dim OriginalRemortCount As Long
    Dim OriginalLevel As Byte
    Dim OriginalExp As Long
    Dim OriginalSkillPts As Integer
    Dim OriginalHp As Integer
    Dim OriginalSkill As Byte
    OriginalRemortCount = UserList(1).Stats.RemortCount
    OriginalLevel = UserList(1).Stats.ELV
    OriginalExp = UserList(1).Stats.Exp
    OriginalSkillPts = UserList(1).Stats.SkillPts
    OriginalHp = UserList(1).Stats.MinHp
    OriginalSkill = UserList(1).Stats.UserSkills(1)

    Call GetRemortEligibility(1)
    test_remort_eligibility_is_read_only = _
        UserList(1).Stats.RemortCount = OriginalRemortCount And _
        UserList(1).Stats.ELV = OriginalLevel And _
        UserList(1).Stats.Exp = OriginalExp And _
        UserList(1).Stats.SkillPts = OriginalSkillPts And _
        UserList(1).Stats.MinHp = OriginalHp And _
        UserList(1).Stats.UserSkills(1) = OriginalSkill
    Exit Function
TestError:
    test_remort_eligibility_is_read_only = False
End Function

Private Function test_remort_capability_bit_and_reset() As Boolean
    On Error GoTo TestError
    Dim OriginalNegotiated As Boolean
    Dim OriginalVersion As Byte
    Dim OriginalMask As Long
    Dim OriginalFeatureEnabled As Boolean
    OriginalNegotiated = UserList(1).HooCapabilities.Negotiated
    OriginalVersion = UserList(1).HooCapabilities.ProtocolVersion
    OriginalMask = UserList(1).HooCapabilities.CapabilityMask
    OriginalFeatureEnabled = IsFeatureEnabled(HOO_FEATURE_REMORT_V1)

    Call SetFeatureToggle(HOO_FEATURE_REMORT_V1, True)
    If AcceptedHooCapabilityMask(HOO_CAP_PROTOCOL_VERSION, HOO_CAP_REMORT_V1) <> HOO_CAP_REMORT_V1 Then GoTo TestDone
    Call SetFeatureToggle(HOO_FEATURE_REMORT_V1, False)
    If AcceptedHooCapabilityMask(HOO_CAP_PROTOCOL_VERSION, HOO_CAP_REMORT_V1) <> 0 Then GoTo TestDone
    Call SetFeatureToggle(HOO_FEATURE_REMORT_V1, True)
    If AcceptedHooCapabilityMask(HOO_CAP_PROTOCOL_VERSION + 1, HOO_CAP_REMORT_V1) <> 0 Then GoTo TestDone
    Call ResetHooClientCapabilities(1)
    If UserSupportsHooCapability(1, HOO_CAP_REMORT_V1) Then GoTo TestDone
    UserList(1).HooCapabilities.Negotiated = True
    UserList(1).HooCapabilities.ProtocolVersion = HOO_CAP_PROTOCOL_VERSION
    UserList(1).HooCapabilities.CapabilityMask = HOO_CAP_ADJACENT_CHARACTERS_V1 Or HOO_CAP_REMORT_V1
    If Not UserSupportsHooCapability(1, HOO_CAP_REMORT_V1) Then GoTo TestDone
    If Not UserSupportsHooCapability(1, HOO_CAP_ADJACENT_CHARACTERS_V1) Then GoTo TestDone
    Call ResetHooClientCapabilities(1)
    If UserList(1).HooCapabilities.Negotiated Then GoTo TestDone
    If UserList(1).HooCapabilities.CapabilityMask <> 0 Then GoTo TestDone
    test_remort_capability_bit_and_reset = (HOO_CAP_REMORT_V1 = &H2&)

TestDone:
    Call SetFeatureToggle(HOO_FEATURE_REMORT_V1, OriginalFeatureEnabled)
    UserList(1).HooCapabilities.Negotiated = OriginalNegotiated
    UserList(1).HooCapabilities.ProtocolVersion = OriginalVersion
    UserList(1).HooCapabilities.CapabilityMask = OriginalMask
    Exit Function
TestError:
    test_remort_capability_bit_and_reset = False
    Resume TestDone
End Function

Private Function test_remort_equipment_and_count_eligibility() As Boolean
    On Error GoTo TestError

    Dim OriginalLevel As Byte
    Dim OriginalDead As Byte
    Dim OriginalInParty As Boolean
    Dim OriginalCount As Long
    Dim OriginalEquipped As Byte
    Dim OriginalObjIndex As Integer
    Dim OriginalMinLevel As Byte
    Dim OriginalMaxLevel As Byte
    Dim OriginalSkillIndex As Byte
    Dim OriginalSkillRequired As Byte
    Dim OriginalQuests(1 To MAXUSERQUESTS) As Integer
    Dim QuestSlot As Integer

    OriginalLevel = UserList(1).Stats.ELV
    OriginalDead = UserList(1).flags.Muerto
    OriginalInParty = UserList(1).Grupo.EnGrupo
    OriginalCount = UserList(1).Stats.RemortCount
    OriginalEquipped = UserList(1).invent.Object(1).Equipped
    OriginalObjIndex = UserList(1).invent.Object(1).ObjIndex
    For QuestSlot = 1 To MAXUSERQUESTS
        OriginalQuests(QuestSlot) = UserList(1).QuestStats.Quests(QuestSlot).QuestIndex
        UserList(1).QuestStats.Quests(QuestSlot).QuestIndex = 0
    Next QuestSlot
    OriginalMinLevel = ObjData(1).MinELV
    OriginalMaxLevel = ObjData(1).MaxLEV
    OriginalSkillIndex = ObjData(1).SkillIndex
    OriginalSkillRequired = ObjData(1).SkillRequerido

    UserList(1).Stats.ELV = STAT_MAXELV
    UserList(1).flags.Muerto = 0
    UserList(1).Grupo.EnGrupo = False
    UserList(1).QuestStats.Quests(1).QuestIndex = 0
    UserList(1).Stats.RemortCount = 0
    UserList(1).invent.Object(1).Equipped = 1
    UserList(1).invent.Object(1).ObjIndex = 1
    ObjData(1).MinELV = 2
    ObjData(1).MaxLEV = 0
    ObjData(1).SkillIndex = 0
    ObjData(1).SkillRequerido = 0
    If GetRemortEligibility(1) <> eRemortEligibility_InvalidEquipment Then GoTo TestDone

    ObjData(1).MinELV = 0
    ObjData(1).SkillIndex = 1
    ObjData(1).SkillRequerido = 1
    If GetRemortEligibility(1) <> eRemortEligibility_InvalidEquipment Then GoTo TestDone

    UserList(1).invent.Object(1).Equipped = 0
    UserList(1).Stats.RemortCount = 2147483647
    If GetRemortEligibility(1) <> eRemortEligibility_RemortLimitReached Then GoTo TestDone
    test_remort_equipment_and_count_eligibility = True

TestDone:
    UserList(1).Stats.ELV = OriginalLevel
    UserList(1).flags.Muerto = OriginalDead
    UserList(1).Grupo.EnGrupo = OriginalInParty
    UserList(1).Stats.RemortCount = OriginalCount
    UserList(1).invent.Object(1).Equipped = OriginalEquipped
    UserList(1).invent.Object(1).ObjIndex = OriginalObjIndex
    For QuestSlot = 1 To MAXUSERQUESTS
        UserList(1).QuestStats.Quests(QuestSlot).QuestIndex = OriginalQuests(QuestSlot)
    Next QuestSlot
    ObjData(1).MinELV = OriginalMinLevel
    ObjData(1).MaxLEV = OriginalMaxLevel
    ObjData(1).SkillIndex = OriginalSkillIndex
    ObjData(1).SkillRequerido = OriginalSkillRequired
    Exit Function
TestError:
    test_remort_equipment_and_count_eligibility = False
    Resume TestDone
End Function

Private Function test_remort_live_reset_and_preservation() As Boolean
    On Error GoTo TestError

    Dim OriginalLevel As Byte, OriginalDead As Byte, OriginalExp As Long
    Dim OriginalCount As Long, OriginalSkillPts As Integer, OriginalModified As Boolean
    Dim OriginalMaxHp As Integer, OriginalMinHp As Integer, OriginalShield As Long
    Dim OriginalMaxMana As Integer, OriginalMinMana As Integer
    Dim OriginalMaxSta As Integer, OriginalMinSta As Integer
    Dim OriginalMaxHit As Integer, OriginalMinHit As Integer
    Dim OriginalMaxWater As Integer, OriginalMinWater As Integer
    Dim OriginalMaxHunger As Integer, OriginalMinHunger As Integer
    Dim OriginalGold As Long, OriginalBankGold As Long, OriginalSpell As Integer
    Dim OriginalGuild As Integer, OriginalClass As e_Class
    Dim OriginalConstitution As Byte, OriginalIntelligence As Byte
    Dim OriginalParty As Boolean
    Dim OriginalQuests(1 To MAXUSERQUESTS) As Integer
    Dim QuestSlot As Integer
    Dim OriginalSkills(1 To NUMSKILLS) As Byte
    Dim OriginalDirty(1 To NUMSKILLS) As Boolean
    Dim OriginalEquipped(1 To MAX_INVENTORY_SLOTS) As Byte
    Dim SkillIndex As Integer, Slot As Integer

    With UserList(1)
        OriginalLevel = .Stats.ELV
        OriginalDead = .flags.Muerto
        OriginalExp = .Stats.Exp
        OriginalCount = .Stats.RemortCount
        OriginalSkillPts = .Stats.SkillPts
        OriginalModified = .flags.ModificoSkills
        OriginalMaxHp = .Stats.MaxHp: OriginalMinHp = .Stats.MinHp: OriginalShield = .Stats.shield
        OriginalMaxMana = .Stats.MaxMAN: OriginalMinMana = .Stats.MinMAN
        OriginalMaxSta = .Stats.MaxSta: OriginalMinSta = .Stats.MinSta
        OriginalMaxHit = .Stats.MaxHit: OriginalMinHit = .Stats.MinHIT
        OriginalMaxWater = .Stats.MaxAGU: OriginalMinWater = .Stats.MinAGU
        OriginalMaxHunger = .Stats.MaxHam: OriginalMinHunger = .Stats.MinHam
        OriginalGold = .Stats.GLD: OriginalBankGold = .Stats.Banco
        OriginalSpell = .Stats.UserHechizos(1): OriginalGuild = .GuildIndex
        OriginalClass = .clase
        OriginalConstitution = .Stats.UserAtributos(e_Atributos.Constitucion)
        OriginalIntelligence = .Stats.UserAtributos(e_Atributos.Inteligencia)
        OriginalParty = .Grupo.EnGrupo
        For QuestSlot = 1 To MAXUSERQUESTS
            OriginalQuests(QuestSlot) = .QuestStats.Quests(QuestSlot).QuestIndex
            .QuestStats.Quests(QuestSlot).QuestIndex = 0
        Next QuestSlot
        For SkillIndex = 1 To NUMSKILLS
            OriginalSkills(SkillIndex) = .Stats.UserSkills(SkillIndex)
            OriginalDirty(SkillIndex) = .Stats.SkillDirty(SkillIndex)
            .Stats.UserSkills(SkillIndex) = 75
            .Stats.SkillDirty(SkillIndex) = False
        Next SkillIndex
        For Slot = 1 To MAX_INVENTORY_SLOTS
            OriginalEquipped(Slot) = .invent.Object(Slot).Equipped
            .invent.Object(Slot).Equipped = 0
        Next Slot
        .clase = e_Class.Mage
        .Stats.UserAtributos(e_Atributos.Constitucion) = 18
        .Stats.UserAtributos(e_Atributos.Inteligencia) = 18
        .Stats.ELV = STAT_MAXELV
        .flags.Muerto = 0
        .Grupo.EnGrupo = False
        .Stats.Exp = 123456
        .Stats.SkillPts = 99
        .Stats.MaxHp = 500: .Stats.MinHp = 250: .Stats.shield = 50
        .Stats.MaxMAN = 500: .Stats.MinMAN = 250
        .Stats.MaxSta = 500: .Stats.MinSta = 250
        .Stats.MaxHit = 100: .Stats.MinHIT = 90
        .Stats.MaxAGU = 25: .Stats.MinAGU = 10
        .Stats.MaxHam = 25: .Stats.MinHam = 10
        .Stats.RemortCount = 3
        .Stats.GLD = 12345: .Stats.Banco = 23456
        .Stats.UserHechizos(1) = 42
        .GuildIndex = 7
    End With

    If Not ApplyRemortProgression(1) Then GoTo TestDone
    With UserList(1)
        If .Stats.ELV <> 1 Or .Stats.Exp <> 0 Or .Stats.RemortCount <> 4 Then GoTo TestDone
        If .Stats.SkillPts <> 10 Then GoTo TestDone
        If .Stats.MaxHp <> 18 Or .Stats.MinHp <> 18 Or .Stats.shield <> 0 Then GoTo TestDone
        If .Stats.MaxMAN <> CInt(18 * ModClase(.clase).ManaInicial) Or .Stats.MinMAN <> .Stats.MaxMAN Then GoTo TestDone
        If .Stats.MaxSta <> 60 Or .Stats.MinSta <> 60 Then GoTo TestDone
        If .Stats.MinHIT <> 1 Or .Stats.MaxHit <> 2 Then GoTo TestDone
        If .Stats.MinAGU <> 100 Or .Stats.MinHam <> 100 Then GoTo TestDone
        For SkillIndex = 1 To NUMSKILLS
            If .Stats.UserSkills(SkillIndex) <> 0 Or Not .Stats.SkillDirty(SkillIndex) Then GoTo TestDone
        Next SkillIndex
        If .Stats.GLD <> 12345 Or .Stats.Banco <> 23456 Then GoTo TestDone
        If .Stats.UserHechizos(1) <> 42 Or .GuildIndex <> 7 Then GoTo TestDone
    End With
    If ApplyRemortProgression(1) Then GoTo TestDone
    If UserList(1).Stats.RemortCount <> 4 Then GoTo TestDone
    test_remort_live_reset_and_preservation = True

TestDone:
    With UserList(1)
        .Stats.ELV = OriginalLevel: .flags.Muerto = OriginalDead: .Stats.Exp = OriginalExp
        .Stats.RemortCount = OriginalCount: .Stats.SkillPts = OriginalSkillPts
        .flags.ModificoSkills = OriginalModified
        .Stats.MaxHp = OriginalMaxHp: .Stats.MinHp = OriginalMinHp: .Stats.shield = OriginalShield
        .Stats.MaxMAN = OriginalMaxMana: .Stats.MinMAN = OriginalMinMana
        .Stats.MaxSta = OriginalMaxSta: .Stats.MinSta = OriginalMinSta
        .Stats.MaxHit = OriginalMaxHit: .Stats.MinHIT = OriginalMinHit
        .Stats.MaxAGU = OriginalMaxWater: .Stats.MinAGU = OriginalMinWater
        .Stats.MaxHam = OriginalMaxHunger: .Stats.MinHam = OriginalMinHunger
        .Stats.GLD = OriginalGold: .Stats.Banco = OriginalBankGold
        .Stats.UserHechizos(1) = OriginalSpell: .GuildIndex = OriginalGuild
        .clase = OriginalClass
        .Stats.UserAtributos(e_Atributos.Constitucion) = OriginalConstitution
        .Stats.UserAtributos(e_Atributos.Inteligencia) = OriginalIntelligence
        .Grupo.EnGrupo = OriginalParty
        For QuestSlot = 1 To MAXUSERQUESTS
            .QuestStats.Quests(QuestSlot).QuestIndex = OriginalQuests(QuestSlot)
        Next QuestSlot
        For SkillIndex = 1 To NUMSKILLS
            .Stats.UserSkills(SkillIndex) = OriginalSkills(SkillIndex)
            .Stats.SkillDirty(SkillIndex) = OriginalDirty(SkillIndex)
        Next SkillIndex
        For Slot = 1 To MAX_INVENTORY_SLOTS
            .invent.Object(Slot).Equipped = OriginalEquipped(Slot)
        Next Slot
    End With
    Exit Function
TestError:
    test_remort_live_reset_and_preservation = False
    Resume TestDone
End Function

Private Function test_suite_taming() As Boolean
    Call test_suite_taming_pure
    Call RunTest("taming rejects an NPC-owned creature", test_taming_npc_owner())
        Call RunTest("taming refreshes same NPC only for owner", test_taming_owner_refresh())
        Call RunTest("taming NoMascotas success stores without ghost", test_taming_special_map(True))
        Call RunTest("taming replacement wild NPC remains ordinary", test_taming_special_map(False))
    test_suite_taming = True
End Function

Private Function test_suite_taming_pure() As Boolean
    Call RunTest("taming power preserves class formulas and rounding", test_taming_power())
    Call RunTest("taming power preserves classic supported range", test_taming_power_range())
    Call RunTest("taming impossible attempts differ from failed rolls", test_taming_results())
    Call RunTest("taming invalid skill overflow is not silently rebalanced", test_taming_overflow())
    test_suite_taming_pure = True
End Function

Private Function test_taming_power() As Boolean
    On Error GoTo TestError
    If GetTamingPower(21, 100, e_Class.Druid) <> 350 Then Exit Function
    If GetTamingPower(18, 100, e_Class.Druid) <> 300 Then Exit Function
    ' Integer assignment uses round-to-even, not truncation or always-round-up.
    If GetTamingPower(21, 1, e_Class.Druid) <> 4 Then Exit Function
    If GetTamingPower(15, 1, e_Class.Druid) <> 2 Then Exit Function
    If GetTamingPower(59, 1, e_Class.Hunter) <> 0 Then Exit Function
    If GetTamingPower(59, 3, e_Class.Cleric) <> 2 Then Exit Function
    If GetTamingPower(21, 100, e_Class.Hunter) <> 18 Then Exit Function
    If GetTamingPower(21, 100, e_Class.Cleric) <> 18 Then Exit Function
    If GetTamingPower(21, 100, e_Class.Warrior) <> 18 Then Exit Function
    test_taming_power = True
    Exit Function
TestError:
    test_taming_power = False
End Function

Private Function test_taming_power_range() As Boolean
    On Error GoTo TestError
    Dim Charisma As Integer, Skill As Integer, Product As Integer
    ' All Byte attributes and legitimate skills fit the existing Integer product.
    For Charisma = 0 To 255
        For Skill = 0 To 100
            Product = CInt(Charisma) * CInt(Skill)
            If GetTamingPower(CByte(Charisma), CByte(Skill), e_Class.Druid) <> CInt(Product / 6) Then Exit Function
            If GetTamingPower(CByte(Charisma), CByte(Skill), e_Class.Hunter) <> CInt(Product / 118) Then Exit Function
        Next Skill
    Next Charisma
    test_taming_power_range = True
    Exit Function
TestError:
    test_taming_power_range = False
End Function

Private Function test_taming_results() As Boolean
    On Error GoTo TestError
    If TamingPresentationResultForMessage(MSG_CRIATURA_HA_ACEPTADO_COMO_SU_AMO) <> eHooTamingResult_Success Then Exit Function
    If TamingPresentationResultForMessage(MSG_NO_TAME_FAILED) <> eHooTamingResult_FailedRoll Then Exit Function
    If TamingPresentationResultForMessage(MSG_NO_PUEDES_DOMAR_ESA_CRIATURA) <> eHooTamingResult_Rejected Then Exit Function
    If TamingPresentationResultForMessage(0) <> eHooTamingResult_Rejected Then Exit Function
    #If PYMMO = 1 Then
        If CInt(ServerPacketID.eReservedHooMagicLightOrb) <> 206 Then Exit Function
        If CInt(ServerPacketID.eHooTamingResult) <> 207 Then Exit Function
    #Else
        If CInt(ServerPacketID.eHooTamingResult) <> CInt(ServerPacketID.eHooHouseDoorActionResult) + 2 Then Exit Function
    #End If
    Dim Roll As Integer, Successes As Integer
    For Roll = 1 To 5
        If GetTamingAttemptMessage(301, 300, Roll) <> MSG_NO_PUEDES_DOMAR_ESA_CRIATURA Then Exit Function
        If GetTamingAttemptMessage(1, 0, Roll) <> MSG_NO_PUEDES_DOMAR_ESA_CRIATURA Then Exit Function
        If Roll = 1 Then
            If GetTamingAttemptMessage(300, 300, Roll) <> MSG_CRIATURA_HA_ACEPTADO_COMO_SU_AMO Then Exit Function
            If GetTamingAttemptMessage(299, 300, Roll) <> MSG_CRIATURA_HA_ACEPTADO_COMO_SU_AMO Then Exit Function
            ' Direct callers retain the original comparison; the network rejects <= 0.
            If GetTamingAttemptMessage(0, 300, Roll) <> MSG_CRIATURA_HA_ACEPTADO_COMO_SU_AMO Then Exit Function
            If GetTamingAttemptMessage(-1, 300, Roll) <> MSG_CRIATURA_HA_ACEPTADO_COMO_SU_AMO Then Exit Function
            Successes = Successes + 1
        Else
            If GetTamingAttemptMessage(300, 300, Roll) <> MSG_NO_TAME_FAILED Then Exit Function
            If GetTamingAttemptMessage(299, 300, Roll) <> MSG_NO_TAME_FAILED Then Exit Function
            If GetTamingAttemptMessage(0, 300, Roll) <> MSG_NO_TAME_FAILED Then Exit Function
            If GetTamingAttemptMessage(-1, 300, Roll) <> MSG_NO_TAME_FAILED Then Exit Function
        End If
    Next Roll
    test_taming_results = (Successes = 1)
    Exit Function
TestError:
    test_taming_results = False
End Function

Private Function test_taming_overflow() As Boolean
    On Error GoTo ExpectedOverflow
    Dim Power As Integer
    ' Invalid/corrupt skills above the 100 cap keep the legacy overflow behavior.
    Power = GetTamingPower(255, 255, e_Class.Druid)
    Exit Function
ExpectedOverflow:
    test_taming_overflow = (Err.Number = 6)
End Function

Public Function CaptureTamingPacket(ByVal Route As Long, ByVal Recipient As Integer) As Boolean
    If Not CaptureTaming Then Exit Function
    Dim writer As Network.Writer, bytes() As Byte, value As Variant
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(bytes)
    Call writer.Clear()
    If CLng(bytes(0)) + CLng(bytes(1)) * 256& = ServerPacketID.eCharacterCreate Then
        value = bytes
        Call TamingCreates.Add(value)
        Call TamingRecipients.Add(Recipient)
        If Route <> SendTarget.ToIndex Then TamingRoutesValid = False
    End If
    CaptureTamingPacket = True
End Function

Private Function test_taming_owner_refresh() As Boolean
    On Error GoTo TestError
    Dim u As Integer, n As Integer, i As Integer, seed As Long, result As e_HooTamingResult
    Dim oldUser As t_User, emptyUser As t_User, oldNpc As t_Npc, emptyNpc As t_Npc
    Dim tile As t_MapBlock, saved As Boolean, bytes As Variant, oldNoPets As Boolean
    For i = UBound(UserList) To 2 Step -1
        If Not UserList(i).flags.UserLogged And Not UserList(i).ConnectionDetails.ConnIDValida Then
            u = i: Exit For
        End If
    Next i
    For i = UBound(NpcList) To 1 Step -1
        If Not NpcList(i).flags.NPCActive Then n = i: Exit For
    Next i
    If u = 0 Or n = 0 Then Exit Function
    oldUser = UserList(u): oldNpc = NpcList(n): tile = MapData(50, 50, 1)
    oldNoPets = MapInfo(1).NoMascotas: saved = True
    UserList(u) = emptyUser: NpcList(n) = emptyNpc
    With UserList(u)
        .VersionId = 1: .Id = 123
        .flags.Privilegios = e_PlayerType.User
        .clase = e_Class.Druid
        .pos.Map = 1: .pos.x = 49: .pos.y = 50
        .Stats.ELV = 40: .Stats.UserSkills(e_Skill.Domar) = 100
        .Stats.UserAtributos(e_Atributos.Carisma) = 21
    End With
    With NpcList(n)
        .VersionId = 1: .Numero = 1: .Char.charindex = 77
        .pos.Map = 1: .pos.x = 50: .pos.y = 50
        .flags.Domable = 1: .MinTameLevel = 1
        .flags.Respawn = 1
    End With
    MapInfo(1).NoMascotas = 0
    Set TamingCreates = New Collection: Set TamingRecipients = New Collection
    CaptureTaming = True: TamingRoutesValid = True
    Call MakeNPCChar(False, u, n, 1, 50, 50)
    bytes = TamingCreates(1)
    If bytes(UBound(bytes) - 29) <> 1 Then GoTo TestDone
    Set TamingCreates = New Collection: Set TamingRecipients = New Collection
    ' Exercise real rolls with a bounded deterministic sequence, not a gameplay override.
    Dim ignored As Single
    ignored = Rnd(-1): Randomize 42
    For seed = 1 To 100
        Call DoDomar(u, n, result)
        If result = eHooTamingResult_Success Then Exit For
        If result <> eHooTamingResult_FailedRoll Or TamingCreates.Count <> 0 Then GoTo TestDone
    Next seed
    If result <> eHooTamingResult_Success Or TamingCreates.Count <> 1 Then GoTo TestDone
    bytes = TamingCreates(1)
    If TamingRecipients(1) <> u Or bytes(2) <> 77 Or bytes(3) <> 0 Then GoTo TestDone
    If bytes(9) <> 50 Or bytes(10) <> 50 Or bytes(UBound(bytes) - 29) <> 2 Then GoTo TestDone
    If UserList(u).MascotasIndex(1).ArrayIndex <> n Or NpcList(n).Char.charindex <> 77 Then GoTo TestDone
    If NpcList(n).MaestroUser.ArrayIndex <> u Or UserList(u).NroMascotas <> 1 Then GoTo TestDone
    Call MakeNPCChar(False, 1, n, 1, 50, 50)
    bytes = TamingCreates(2)
    If TamingRecipients(2) <> 1 Or bytes(UBound(bytes) - 29) <> 1 Then GoTo TestDone
    Call DoDomar(u, n, result)
    If result <> eHooTamingResult_Rejected Or TamingCreates.Count <> 2 Then GoTo TestDone
    test_taming_owner_refresh = TamingRoutesValid
TestDone:
    CaptureTaming = False
    Set TamingCreates = Nothing: Set TamingRecipients = Nothing
    If saved Then
        UserList(u) = oldUser: NpcList(n) = oldNpc: MapData(50, 50, 1) = tile
        MapInfo(1).NoMascotas = oldNoPets
    End If
    Exit Function
TestError:
    Call Progress("taming ownership fixture: " & CStr(Err.Number) & " " & Err.Description)
    Resume TestDone
End Function

Private Function test_taming_special_map(ByVal noPets As Boolean) As Boolean
    On Error GoTo TestError
    Dim u As Integer, n As Integer, replacement As Integer, i As Long, charindex As Integer
    Dim oldUser As t_User, emptyUser As t_User, oldNoPets As Boolean, saved As Boolean
    Dim origin As t_WorldPos, result As e_HooTamingResult, bytes As Variant, ignored As Single
    Dim activeBefore(1 To MaxNPCs) As Boolean
    For i = UBound(UserList) To 2 Step -1
        If Not UserList(i).flags.UserLogged And Not UserList(i).ConnectionDetails.ConnIDValida Then u = i: Exit For
    Next i
    If u = 0 Then Exit Function
    oldUser = UserList(u): oldNoPets = MapInfo(1).NoMascotas: saved = True
    UserList(u) = emptyUser
    With UserList(u)
        .VersionId = 1: .Id = 123: .clase = e_Class.Druid
        .flags.Privilegios = e_PlayerType.User
        .pos.Map = 1: .Stats.ELV = 40
        .Stats.UserSkills(e_Skill.Domar) = 100: .Stats.UserAtributos(e_Atributos.Carisma) = 21
    End With
    Set TamingCreates = New Collection: Set TamingRecipients = New Collection
    CaptureTaming = True: TamingRoutesValid = True
    origin.Map = 1: origin.x = 50: origin.y = 50
    ' Exercise the real allocator, spawn and removal with an ordinary loaded creature.
    n = CrearNPC(501, 1, origin)
    If n = 0 Then GoTo TestDone
    charindex = NpcList(n).Char.charindex
    UserList(u).pos = NpcList(n).pos
    NpcList(n).flags.Domable = 1: NpcList(n).MinTameLevel = 1
    If noPets Then NpcList(n).flags.Respawn = 1 Else NpcList(n).flags.Respawn = 0
    MapInfo(1).NoMascotas = noPets
    Set TamingCreates = New Collection: Set TamingRecipients = New Collection
    For i = 1 To MaxNPCs
        activeBefore(i) = NpcList(i).flags.NPCActive
    Next i
    ignored = Rnd(-1): Randomize 42
    For i = 1 To 100
        Call DoDomar(u, n, result)
        If result = eHooTamingResult_Success Then Exit For
        If result <> eHooTamingResult_FailedRoll Or TamingCreates.Count <> 0 Then GoTo TestDone
    Next i
    If result <> eHooTamingResult_Success Then GoTo TestDone
    If noPets Then
        If TamingCreates.Count <> 0 Or NpcList(n).flags.NPCActive Then GoTo TestDone
        If UserList(u).NroMascotas <> 1 Or UserList(u).MascotasType(1) <> 501 Then GoTo TestDone
        If IsValidNpcRef(UserList(u).MascotasIndex(1)) Then GoTo TestDone
    Else
        For i = 1 To MaxNPCs
            If NpcList(i).flags.NPCActive And Not activeBefore(i) Then
                If replacement <> 0 Then GoTo TestDone
                replacement = i
            End If
        Next i
        If replacement = 0 Or replacement = n Or TamingCreates.Count <> 1 Then GoTo TestDone
        bytes = TamingCreates(1)
        If CLng(bytes(2)) + CLng(bytes(3)) * 256& <> charindex Then GoTo TestDone
        If TamingRecipients(1) <> u Or bytes(UBound(bytes) - 29) <> 2 Then GoTo TestDone
        If UserList(u).MascotasIndex(1).ArrayIndex <> n Then GoTo TestDone
        If IsValidUserRef(NpcList(replacement).MaestroUser) Or NpcList(replacement).Numero <> 501 Then GoTo TestDone
        Call MakeNPCChar(False, u, replacement, NpcList(replacement).pos.Map, NpcList(replacement).pos.x, NpcList(replacement).pos.y)
        bytes = TamingCreates(2)
        If CLng(bytes(2)) + CLng(bytes(3)) * 256& = charindex Then GoTo TestDone
        If bytes(UBound(bytes) - 29) <> 1 Then GoTo TestDone
    End If
    test_taming_special_map = True
TestDone:
    If replacement > 0 Then
        If NpcList(replacement).flags.NPCActive Then Call QuitarNPC(replacement, eNewPet)
    End If
    If n > 0 Then
        If NpcList(n).flags.NPCActive Then Call QuitarNPC(n, eNewPet)
    End If
    CaptureTaming = False
    Set TamingCreates = Nothing: Set TamingRecipients = Nothing
    If saved Then
        UserList(u) = oldUser: MapInfo(1).NoMascotas = oldNoPets
    End If
    Exit Function
TestError:
    Call Progress("taming map fixture: " & CStr(Err.Number) & " " & Err.Description)
    Resume TestDone
End Function

Private Function test_taming_npc_owner() As Boolean
    On Error GoTo TestError
    Dim UserIndex As Integer, NpcIndex As Integer, OwnerIndex As Integer, i As Integer
    Dim OldUser As t_User, EmptyUser As t_User, OldNpc As t_Npc, EmptyNpc As t_Npc
    Dim Saved As Boolean
    ' Use idle slots without touching spawned NPCs or depending on map contents.
    For i = UBound(UserList) To LBound(UserList) Step -1
        If Not UserList(i).flags.UserLogged And Not UserList(i).ConnectionDetails.ConnIDValida Then
            UserIndex = i
            Exit For
        End If
    Next i
    For i = UBound(NpcList) To LBound(NpcList) Step -1
        If Not NpcList(i).flags.NPCActive Then
            If NpcIndex = 0 Then
                NpcIndex = i
            Else
                OwnerIndex = i
                Exit For
            End If
        End If
    Next i
    If UserIndex = 0 Or NpcIndex = 0 Or OwnerIndex = 0 Then Exit Function
    OldUser = UserList(UserIndex)
    OldNpc = NpcList(NpcIndex)
    Saved = True
    UserList(UserIndex) = EmptyUser
    NpcList(NpcIndex) = EmptyNpc
    With UserList(UserIndex)
        .flags.UltimoMensaje = MSG_NO_TAME_FAILED
        .flags.Privilegios = e_PlayerType.User
        .clase = e_Class.Druid
        .Stats.ELV = 40
        .Stats.UserSkills(e_Skill.Domar) = 100
        .Stats.UserAtributos(e_Atributos.Carisma) = 21
    End With
    ' No roll can tame this fixture; capped skill prevents progression side effects.
    NpcList(NpcIndex).flags.Domable = 32767
    NpcList(NpcIndex).MinTameLevel = 1
    Call SetNpcRef(NpcList(NpcIndex).MaestroNPC, OwnerIndex)
    If Not IsValidNpcRef(NpcList(NpcIndex).MaestroNPC) Then GoTo TestDone
    Call DoDomar(UserIndex, NpcIndex)
    If UserList(UserIndex).flags.UltimoMensaje <> MSG_NO_TAME_FAILED Then GoTo TestDone
    If UserList(UserIndex).NroMascotas <> 0 Then GoTo TestDone
    ' Negative control: removing only the owner must reach the power rejection.
    ' This also prevents an unrelated early exit from making the test pass.
    Call ClearNpcRef(NpcList(NpcIndex).MaestroNPC)
    Call DoDomar(UserIndex, NpcIndex)
    test_taming_npc_owner = (UserList(UserIndex).flags.UltimoMensaje = MSG_NO_PUEDES_DOMAR_ESA_CRIATURA And UserList(UserIndex).NroMascotas = 0)
TestDone:
    If Saved Then
        UserList(UserIndex) = OldUser
        NpcList(NpcIndex) = OldNpc
    End If
    Exit Function
TestError:
    test_taming_npc_owner = False
    Resume TestDone
End Function

#End If
