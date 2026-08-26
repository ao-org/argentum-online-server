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
    Private TotalTests     As Integer
    Private PassedTests    As Integer
    Private FailedTests    As Integer
    Private FailedTestNames() As String
    Private FailedTestCount As Integer
    Private TotalElapsed   As Double

    Private Const SUITE_COUNT As Integer = 39

Public Sub Init()
    On Error GoTo Init_Err
    'We can mock the key value to test errors...
    private_key = PrivateKey
    character_name = "seneca"
    character_name = RandomName(16)
    'Hardcoded token for unit testing...

    decrypted_token = "G7H5wKOKZvebZxHtnkRtJNvL/AHWEw3dHCyBTzXVvdTe3bQAJHePsFfV/Ecgm9Wk"
    encrypted_token = AO20CryptoSysWrapper.ENCRYPT(private_key, decrypted_token)
    public_key = mid$(decrypted_token, 1, 16)
    
    'Add a fake token to be using when exercising the protocol for LoginNewChar
    Call AddTokenDatabase(encrypted_token, decrypted_token, "MORGOLOCK2002@YAHOO.COM.AR")

Init_Resume:
    ' Reset test runner state
    TotalTests = 0
    PassedTests = 0
    FailedTests = 0
    FailedTestCount = 0
    TotalElapsed = 0
    ReDim FailedTestNames(0)
    Exit Sub
    
Init_Err:
    Resume Init_Resume
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
    Debug.Assert (MapData(1, 54, 51).UserIndex = 1)
    Debug.Assert (UserList(1).Char.charindex <> 0)
    'Delete first user
    Call EraseUserChar(1, False, False)
    Debug.Assert (MapData(1, 54, 55).UserIndex = 0)
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
    Debug.Assert (MapData(1, 54, 56).UserIndex = 2)
    Debug.Assert (UserList(2).Char.charindex <> 0)
    
    Call test_make_user(1, 1, 50, 46)
    Debug.Assert (MapData(1, 50, 46).UserIndex = 1)
    Debug.Assert (UserList(1).Char.charindex <> 0)
    Debug.Assert (UserList(2).Char.charindex <> UserList(1).Char.charindex)
    
    'Delete user 2
    Call EraseUserChar(2, False, False)
    Debug.Assert (MapData(1, 54, 56).UserIndex = 0)
    Debug.Assert (UserList(2).Char.charindex = 0)
    'Create user 2 again
    Call test_make_user(2, 1, 54, 56)
    Debug.Assert (MapData(1, 54, 56).UserIndex = 2)
    Debug.Assert (UserList(2).Char.charindex <> 0)
    
    For i = 1 To UBound(UserList)
        If UserList(i).Char.charindex <> 0 Then
            Call EraseUserChar(i, False, True)
        End If
    Next i
    
    Call test_make_user(1, 1, 64, 66)
    Debug.Assert (MapData(1, 64, 66).UserIndex = 1)
    Debug.Assert (UserList(1).Char.charindex <> 0)
    Debug.Assert (UserList(1).Char.charindex = 1)
    
    
    Call test_make_user(1, 1, 68, 66)
    Debug.Assert (MapData(1, 68, 66).UserIndex = 1)
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
        Case 33: RunSuite = Unit_Partition.test_suite_partition()
#If DIRECT_PLAY = 1 Then
        Case 34: RunSuite = Unit_Network.test_suite_network()
#End If
#If DIRECT_PLAY = 0 Then
        Case 35: RunSuite = Unit_Network_Aurora.test_suite_network_aurora()
#End If
        Case 36: RunSuite = Unit_Weather.test_suite_weather()
        Case 37: RunSuite = test_suite_remort_persistence()
        Case 38: RunSuite = test_suite_remort_capability_state()
        Case 39: RunSuite = Unit_NpcCrossMapPursuit.test_suite_npc_cross_map_pursuit()
        Case Else
            RunSuite = False
    End Select
    Exit Function
    
RunSuite_Err:
    RunSuite = False
End Function

Public Sub RunAllSuites()
    Dim sw As Instruments
    Set sw = New Instruments
    sw.start
    
    Dim i As Integer
    For i = 1 To SUITE_COUNT
        Call RunSuite(i)
    Next i
    
    TotalElapsed = sw.ElapsedMilliseconds
End Sub

Public Sub WriteResultsToFile(ByVal filePath As String)
    On Error GoTo WriteResultsToFile_Err
    Dim f As Integer
    f = FreeFile
    Open filePath For Output As #f
    Print #f, "=== AO20 TEST REPORT ==="
    Print #f, "Total: " & TotalTests & " | Passed: " & PassedTests & " | Failed: " & FailedTests
    
    If FailedTestCount > 0 Then
        Print #f, "Failed tests:"
        Dim i As Integer
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
    Exit Sub
WriteResultsToFile_Err:
    Close #f
End Sub

Public Function GetFailedTests() As Integer
    GetFailedTests = FailedTests
End Function

Private Function test_suite_remort_persistence() As Boolean
    Call RunTest("remort_count migration default", test_remort_count_migration_default())
    Call RunTest("remort_count load validation", test_remort_count_load_validation())
    Call RunTest("remort_count normal save SQL", test_remort_count_normal_save_sql())
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

Private Function test_suite_remort_capability_state() As Boolean
    Call RunTest("remort eligibility and priority", test_remort_eligibility_and_priority())
    Call RunTest("remort eligibility is read-only", test_remort_eligibility_is_read_only())
    Call RunTest("remort capability bit and reset", test_remort_capability_bit_and_reset())
    Call RunTest("remort packet is appended", CInt(ServerPacketID.eRemortState) = CInt(ServerPacketID.eShowPickUpObj) + 1)
    Call RunTest("remort operation packet IDs are appended", CInt(ServerPacketID.eRemortResult) = CInt(ServerPacketID.eRemortState) + 1 And CInt(ClientPacketID.eRequestRemort) = CInt(ClientPacketID.eHooClientCapabilities) + 1)
    Call RunTest("targeted spell packet IDs are appended", CInt(ServerPacketID.eHooTargetedSpellCastResult) = CInt(ServerPacketID.eRemortResult) + 1 And CInt(ClientPacketID.eHooTargetedSpellCast) = CInt(ClientPacketID.eRequestRemort) + 1)
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

#End If
