Attribute VB_Name = "Unit_IniManager"
Option Explicit
#If UNIT_TEST = 1 Then

Private Const TEMP_INI_FILE As String = "test_temp_ini.ini"
Private Const TEMP_INI_DUMP_FILE As String = "test_temp_ini_dump.ini"

Public Function test_suite_ini_manager() As Boolean
    Call UnitTesting.RunTest("test_ini_parse_and_getvalue", test_ini_parse_and_getvalue())
    Call UnitTesting.RunTest("test_ini_keyexists_true", test_ini_keyexists_true())
    Call UnitTesting.RunTest("test_ini_keyexists_false", test_ini_keyexists_false())
    Call UnitTesting.RunTest("test_ini_changevalue_existing", test_ini_changevalue_existing())
    Call UnitTesting.RunTest("test_ini_changevalue_new_section", test_ini_changevalue_new_section())
    Call UnitTesting.RunTest("test_ini_nodescount", test_ini_nodescount())
    Call UnitTesting.RunTest("test_ini_entriescount", test_ini_entriescount())
    Call UnitTesting.RunTest("test_ini_pbt_dump_round_trip", test_ini_pbt_dump_round_trip())
    Call UnitTesting.RunTest("test_ini_pbt_changevalue_roundtrip", test_ini_pbt_changevalue_roundtrip())
    
    Call delete_temp_files
    test_suite_ini_manager = True
End Function

' Helper: writes a temp INI file with known content and returns the full path.
Private Function write_temp_ini(ByVal content As String) As String
    Dim filePath As String
    filePath = App.Path & "\" & TEMP_INI_FILE
    
    Dim f As Integer
    f = FreeFile
    Open filePath For Output As #f
    Print #f, content
    Close #f
    
    write_temp_ini = filePath
End Function

' Cleanup helper: removes temp files created during tests.
Private Sub delete_temp_files()
    On Error Resume Next
    Kill App.Path & "\" & TEMP_INI_FILE
    Kill App.Path & "\" & TEMP_INI_DUMP_FILE
    On Error GoTo 0
End Sub

' Req 6.1: Parse valid INI and retrieve correct values via GetValue.
Private Function test_ini_parse_and_getvalue() As Boolean
    On Error GoTo Err_Handler
    Dim filePath As String
    filePath = write_temp_ini("[SECTION1]" & vbCrLf & "Key1=Value1" & vbCrLf & "Key2=42" & vbCrLf & "[SECTION2]" & vbCrLf & "Name=Test")
    
    Dim ini As New clsIniManager
    Call ini.Initialize(filePath)
    
    test_ini_parse_and_getvalue = True
    
    If ini.GetValue("SECTION1", "Key1") <> "Value1" Then test_ini_parse_and_getvalue = False: Exit Function
    If ini.GetValue("SECTION1", "Key2") <> "42" Then test_ini_parse_and_getvalue = False: Exit Function
    If ini.GetValue("SECTION2", "Name") <> "Test" Then test_ini_parse_and_getvalue = False: Exit Function
    
    ' Non-existent key should return empty string (default)
    If ini.GetValue("SECTION1", "Missing") <> "" Then test_ini_parse_and_getvalue = False: Exit Function
    
    Exit Function
Err_Handler:
    test_ini_parse_and_getvalue = False
End Function

' Req 6.2: KeyExists returns True for an existing section.
Private Function test_ini_keyexists_true() As Boolean
    On Error GoTo Err_Handler
    Dim filePath As String
    filePath = write_temp_ini("[MYSECTION]" & vbCrLf & "A=1")
    
    Dim ini As New clsIniManager
    Call ini.Initialize(filePath)
    
    test_ini_keyexists_true = ini.KeyExists("MYSECTION")
    Exit Function
Err_Handler:
    test_ini_keyexists_true = False
End Function

' Req 6.3: KeyExists returns False for a non-existent section.
Private Function test_ini_keyexists_false() As Boolean
    On Error GoTo Err_Handler
    Dim filePath As String
    filePath = write_temp_ini("[EXISTING]" & vbCrLf & "A=1")
    
    Dim ini As New clsIniManager
    Call ini.Initialize(filePath)
    
    test_ini_keyexists_false = Not ini.KeyExists("NONEXISTENT")
    Exit Function
Err_Handler:
    test_ini_keyexists_false = False
End Function

' Req 6.4: ChangeValue on an existing key updates the value retrievable via GetValue.
Private Function test_ini_changevalue_existing() As Boolean
    On Error GoTo Err_Handler
    Dim filePath As String
    filePath = write_temp_ini("[CONFIG]" & vbCrLf & "Port=7666")
    
    Dim ini As New clsIniManager
    Call ini.Initialize(filePath)
    
    Call ini.ChangeValue("CONFIG", "Port", "8080")
    
    test_ini_changevalue_existing = (ini.GetValue("CONFIG", "Port") = "8080")
    Exit Function
Err_Handler:
    test_ini_changevalue_existing = False
End Function

' Req 6.5: ChangeValue for a new section and key creates the section and stores the value.
Private Function test_ini_changevalue_new_section() As Boolean
    On Error GoTo Err_Handler
    Dim filePath As String
    filePath = write_temp_ini("[INIT]" & vbCrLf & "X=1")
    
    Dim ini As New clsIniManager
    Call ini.Initialize(filePath)
    
    ' Add a completely new section and key
    Call ini.ChangeValue("NEWSECTION", "NewKey", "NewValue")
    
    test_ini_changevalue_new_section = True
    
    If ini.GetValue("NEWSECTION", "NewKey") <> "NewValue" Then test_ini_changevalue_new_section = False: Exit Function
    If Not ini.KeyExists("NEWSECTION") Then test_ini_changevalue_new_section = False: Exit Function
    
    Exit Function
Err_Handler:
    test_ini_changevalue_new_section = False
End Function

' Req 6.6: NodesCount matches the number of sections in the file.
Private Function test_ini_nodescount() As Boolean
    On Error GoTo Err_Handler
    Dim filePath As String
    filePath = write_temp_ini("[SEC1]" & vbCrLf & "A=1" & vbCrLf & "[SEC2]" & vbCrLf & "B=2" & vbCrLf & "[SEC3]" & vbCrLf & "C=3")
    
    Dim ini As New clsIniManager
    Call ini.Initialize(filePath)
    
    test_ini_nodescount = (ini.NodesCount = 3)
    Exit Function
Err_Handler:
    test_ini_nodescount = False
End Function

' Req 6.7: EntriesCount returns the correct number of entries for each section.
Private Function test_ini_entriescount() As Boolean
    On Error GoTo Err_Handler
    Dim filePath As String
    filePath = write_temp_ini("[SMALL]" & vbCrLf & "A=1" & vbCrLf & "[BIG]" & vbCrLf & "X=10" & vbCrLf & "Y=20" & vbCrLf & "Z=30")
    
    Dim ini As New clsIniManager
    Call ini.Initialize(filePath)
    
    test_ini_entriescount = True
    
    If ini.EntriesCount("SMALL") <> 1 Then test_ini_entriescount = False: Exit Function
    If ini.EntriesCount("BIG") <> 3 Then test_ini_entriescount = False: Exit Function
    
    Exit Function
Err_Handler:
    test_ini_entriescount = False
End Function

' Feature: unit-test-coverage, Property 5: IniManager dump round-trip
' Generates 100+ INI contents with varying sections/keys, loads each into
' a clsIniManager, dumps via DumpFile, reloads the dump, and verifies all
' values match (case-insensitive).
Private Function test_ini_pbt_dump_round_trip() As Boolean
    On Error GoTo Fail
    
    Dim iteration As Long
    Dim numSections As Long
    Dim numKeys As Long
    Dim s As Long
    Dim k As Long
    Dim content As String
    Dim secName As String
    Dim keyName As String
    Dim valStr As String
    Dim srcPath As String
    Dim dumpPath As String
    Dim iniSrc As clsIniManager
    Dim iniReloaded As clsIniManager
    Dim origKey As String
    Dim origVal As String
    Dim reloadedVal As String
    
    dumpPath = App.Path & "\" & TEMP_INI_DUMP_FILE
    
    For iteration = 1 To 110
        ' Vary the number of sections (1 to 4) and keys per section (1 to 4)
        numSections = ((iteration - 1) Mod 4) + 1
        numKeys = (((iteration - 1) \ 4) Mod 4) + 1
        
        ' Build INI content string
        content = ""
        For s = 1 To numSections
            secName = "SEC" & s & "I" & iteration
            content = content & "[" & secName & "]" & vbCrLf
            For k = 1 To numKeys
                keyName = "Key" & k
                valStr = "Val" & iteration & "_" & s & "_" & k
                content = content & keyName & "=" & valStr & vbCrLf
            Next k
        Next s
        
        ' Write source INI, load it
        srcPath = write_temp_ini(content)
        Set iniSrc = New clsIniManager
        Call iniSrc.Initialize(srcPath)
        
        ' Dump to second file
        Call iniSrc.DumpFile(dumpPath)
        
        ' Reload the dump
        Set iniReloaded = New clsIniManager
        Call iniReloaded.Initialize(dumpPath)
        
        ' Verify node counts match
        If iniReloaded.NodesCount <> iniSrc.NodesCount Then
            test_ini_pbt_dump_round_trip = False
            Exit Function
        End If
        
        ' Verify all sections and key-value pairs match
        For s = 0 To iniSrc.NodesCount - 1
            secName = iniSrc.GetNode(s)
            
            If iniReloaded.EntriesCount(secName) <> iniSrc.EntriesCount(secName) Then
                test_ini_pbt_dump_round_trip = False
                Exit Function
            End If
            
            For k = 0 To iniSrc.EntriesCount(secName) - 1
                Call iniSrc.GetPair(secName, k, origKey, origVal)
                reloadedVal = iniReloaded.GetValue(secName, origKey)
                
                ' Case-insensitive comparison
                If StrComp(origVal, reloadedVal, vbTextCompare) <> 0 Then
                    test_ini_pbt_dump_round_trip = False
                    Exit Function
                End If
            Next k
        Next s
        
        Set iniSrc = Nothing
        Set iniReloaded = Nothing
    Next iteration
    
    test_ini_pbt_dump_round_trip = True
    Exit Function
Fail:
    test_ini_pbt_dump_round_trip = False
End Function

' Feature: unit-test-coverage, Property 6: IniManager ChangeValue/GetValue round-trip
' Creates a single clsIniManager instance, then loops 110+ times calling
' ChangeValue with unique (section, key, value) triples and immediately
' verifying via GetValue that the returned value matches.
Private Function test_ini_pbt_changevalue_roundtrip() As Boolean
    On Error GoTo Fail
    
    Dim filePath As String
    Dim ini As New clsIniManager
    Dim i As Long
    Dim secName As String
    Dim keyName As String
    Dim valStr As String
    Dim retrieved As String
    
    ' Start with a minimal INI file so the manager is initialized
    filePath = write_temp_ini("[INIT]" & vbCrLf & "Seed=0")
    Call ini.Initialize(filePath)
    
    ' Loop 110 times with deterministic unique triples
    For i = 1 To 110
        secName = "S" & i
        keyName = "K" & i
        valStr = "V" & (i * 7)
        
        Call ini.ChangeValue(secName, keyName, valStr)
        
        retrieved = ini.GetValue(secName, keyName)
        
        If retrieved <> valStr Then
            test_ini_pbt_changevalue_roundtrip = False
            Exit Function
        End If
    Next i
    
    test_ini_pbt_changevalue_roundtrip = True
    Exit Function
Fail:
    test_ini_pbt_changevalue_roundtrip = False
End Function

#End If
