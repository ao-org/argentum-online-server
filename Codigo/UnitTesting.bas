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

    Private Const SUITE_COUNT As Integer = 4

Public Sub Init()
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
    
    ' Reset test runner state
    TotalTests = 0
    PassedTests = 0
    FailedTests = 0
    FailedTestCount = 0
    TotalElapsed = 0
    ReDim FailedTestNames(0)
    
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
        Debug.Print "[PASS] " & testName
    Else
        FailedTests = FailedTests + 1
        FailedTestCount = FailedTestCount + 1
        ReDim Preserve FailedTestNames(FailedTestCount)
        FailedTestNames(FailedTestCount) = testName
        Debug.Print "[FAIL] " & testName
        If RunningInVB() Then
            Debug.Assert False
        End If
    End If
End Sub

Public Sub RunTestError(ByVal testName As String, ByVal errorDesc As String)
    TotalTests = TotalTests + 1
    FailedTests = FailedTests + 1
    FailedTestCount = FailedTestCount + 1
    ReDim Preserve FailedTestNames(FailedTestCount)
    FailedTestNames(FailedTestCount) = testName & " - Error: " & errorDesc
    Debug.Print "[FAIL] " & testName & " - Error: " & errorDesc
    If RunningInVB() Then
        Debug.Assert False
    End If
End Sub

Private Function RunSuite(ByVal suiteIndex As Integer) As Boolean
    On Error GoTo RunSuite_Err
    
    Select Case suiteIndex
        Case 1: RunSuite = Unit_Math.test_suite_math()
        Case 2: RunSuite = Unit_Bitmask.test_suite_bitmask()
        Case 3: RunSuite = Unit_StringValidation.test_suite_strings()
        Case 4: RunSuite = Unit_Pathfinding.test_suite_pathfinding()
        Case Else
            RunSuite = False
    End Selectt "[ERROR] Suite " & suiteIndex & " raised error: " & Err.Description
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
    
    Call PrintTestReport
    
    ' Run protocol integration tests after synchronous suites
    Debug.Print "Running proto suite, trying to connect to 127.0.0.1:7667"
    Call UnitClient.Init
    Call UnitClient.Connect("127.0.0.1", "7667")
End Sub

Public Sub PrintTestReport()
    Debug.Print ""
    Debug.Print "=== AO20 TEST REPORT ==="
    Debug.Print "Total: " & TotalTests & " | Passed: " & PassedTests & " | Failed: " & FailedTests
    
    If FailedTestCount > 0 Then
        Debug.Print "Failed tests:"
        Dim i As Integer
        For i = 1 To FailedTestCount
            Debug.Print "  - " & FailedTestNames(i)
        Next i
    End If
    
    Debug.Print "Total time: " & Format$(TotalElapsed, "0.00") & " ms"
    
    If FailedTests = 0 Then
        Debug.Print "ALL TESTS PASSED"
    Else
        Debug.Print "TESTS FAILED"
    End If
    Debug.Print "========================="
    Debug.Print ""
End Sub

#End If
    Debug.Print "========================="
    Debug.Print ""
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

#End If