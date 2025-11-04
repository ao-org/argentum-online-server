Attribute VB_Name = "modPartition"
' ===== modPartitions.bas =====
Option Explicit
' --- INI & IO helpers ---
Private Declare Function GetPrivateProfileStringA Lib "kernel32" ( _
    ByVal lpAppName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    
' Public API:
'   Call Partitions_Load(700)                             ' parse App.Path & "\Partitions.ini"
'   s = OwnerForMap(101)                                  ' -> "S2" (example)
'   ok = GetEndpoint("S2", host$, port&)                  ' -> True, "10.0.0.4", 6503
'   shard$ = GetShardIdFromCmdline()                      ' -> "S1" or "" if not present
'
' Ownership syntax (in Partitions.ini, section [ownership]):
'   1-18 = S1
'   19-26,101 = S2
'   default = S1   ' optional fallback for unassigned maps
'
' Endpoints (in Partitions.ini, sections [servers.<ID>]):
'   [servers.S1]
'   host = 10.0.0.4
'   port = 6501


Public ShardID As String
' -------- Internal state --------
Private gMaxMap As Long
Private gOwner() As String           ' 1..gMaxMap ? serverId
Private gHosts As Object             ' Scripting.Dictionary: serverId ? host
Private gPorts As Object             ' Scripting.Dictionary: serverId ? port (Long)
Private gErr As String

Public Function IsShardingEnabled() As Boolean
    IsShardingEnabled = ShardID <> vbNullString
End Function

Public Sub TryInitShard()
    ShardID = GetShardIdFromCmdline()                      ' -> "S1" or "" if not present
End Sub

Public Function OwnerForMap(ByVal mapId As Long) As String
    If mapId < 1 Or mapId > gMaxMap Then
        OwnerForMap = ""
    Else
        OwnerForMap = gOwner(mapId)
    End If
End Function

Public Function GetEndpoint(ByVal serverId As String, ByRef host As String, ByRef port As Long) As Boolean
    Dim key As String
    key = LCase$(Trim$(serverId))
    If gHosts Is Nothing Or gPorts Is Nothing Then
        host = "": port = 0: GetEndpoint = False: Exit Function
    End If
    If gHosts.Exists(key) And gPorts.Exists(key) Then
        host = CStr(gHosts.Item(key))
        port = CLng(gPorts.Item(key))
        GetEndpoint = (LenB(host) > 0 And port > 0)
    Else
        host = "": port = 0
        GetEndpoint = False
    End If
End Function

Public Sub Partitions_Load(ByVal maxMap As Long)
    Dim iniPath As String
    Dim text As String, lines() As String
    Dim section As String, line As String
    Dim i As Long, eqPos As Long
    Dim leftSpec As String, serverId As String
    Dim defId As String
    Dim serversNeeded As Object ' Scripting.Dictionary used as a set
    Dim sid As Variant

    gMaxMap = maxMap
    iniPath = App.path & "\Partitions.ini"

    ReDim gOwner(1 To gMaxMap)
    Set gHosts = CreateObject("Scripting.Dictionary")
    Set gPorts = CreateObject("Scripting.Dictionary")
    Set serversNeeded = CreateObject("Scripting.Dictionary")
    gErr = vbNullString

    text = ReadAllText(iniPath)
    If LenB(text) = 0 Then RaiseErr "Partitions_Load", "Partitions.ini missing or empty: " & iniPath

    lines = Split(text, vbCrLf)

    ' Pass 1: [ownership]
    For i = 0 To UBound(lines)
        line = Trim$(lines(i))
        If LenB(line) = 0 Or Left$(line, 1) = ";" Then GoTo cont

        If Left$(line, 1) = "[" Then
            section = mid$(line, 2, Len(line) - 2)
            GoTo cont
        End If

        If LCase$(section) = "ownership" Then
            eqPos = InStr(1, line, "=", vbTextCompare)
            If eqPos = 0 Then AddErr "ownership: missing '=' on line " & CStr(i + 1): GoTo cont

            leftSpec = Trim$(Left$(line, eqPos - 1))
            serverId = Trim$(mid$(line, eqPos + 1))

            If LCase$(leftSpec) = "default" Then
                defId = serverId
                GoTo cont
            End If
            If LenB(serverId) = 0 Then
                AddErr "ownership: empty server id on line " & CStr(i + 1)
                GoTo cont
            End If

            ApplyOwnershipSpec leftSpec, serverId, i + 1
            If Not serversNeeded.Exists(LCase$(serverId)) Then serversNeeded.Add LCase$(serverId), True
        End If
cont:
    Next i

    ' Fill defaults if provided
    If LenB(defId) > 0 Then
        For i = 1 To gMaxMap
            If LenB(gOwner(i)) = 0 Then gOwner(i) = defId
        Next i
    End If

    ' Validate that every map is assigned
    For i = 1 To gMaxMap
        If LenB(gOwner(i)) = 0 Then
            AddErr "Unassigned map " & CStr(i) & " (and possibly others)"
            Exit For
        End If
    Next i

    ' Pass 2: [servers.*] endpoints for referenced owners
    For Each sid In serversNeeded.Keys
        Dim sidStr As String, h As String, p As Long
        sidStr = CStr(sid)
        h = IniRead(iniPath, "servers." & sidStr, "host", "")
        p = SafeCLng(IniRead(iniPath, "servers." & sidStr, "port", "0"))
        If LenB(h) = 0 Or p <= 0 Then
            AddErr "servers." & sidStr & ": missing/invalid host or port"
        Else
            AddEndpoint sidStr, h, p
        End If
    Next sid

    If LenB(gErr) > 0 Then RaiseErr "Partitions_Load", gErr
End Sub

' Returns "S1" from "/shard=S1" ("" if absent)
Public Function GetShardIdFromCmdline() As String
    Dim cmd As String, pref As String, p As Long, rest As String, sp As Long, v As String
    pref = "/shard="
    cmd = Command$
    p = InStr(1, cmd, pref, vbTextCompare)
    If p = 0 Then GetShardIdFromCmdline = "": Exit Function
    rest = mid$(cmd, p + Len(pref))
    sp = InStr(rest, " ")
    If sp > 0 Then v = Left$(rest, sp - 1) Else v = rest
    GetShardIdFromCmdline = Trim$(v)
End Function

Private Function SafeCLng(ByVal s As String) As Long
    On Error GoTo fail
    SafeCLng = CLng(Trim$(s))
    Exit Function
fail:
    SafeCLng = 0
End Function


' -------- Helpers --------
Private Sub ApplyOwnershipSpec(ByVal spec As String, ByVal serverId As String, ByVal lineNo As Long)
    Dim parts() As String, tok As String
    Dim dash As Long, a As Long, b As Long
    Dim i As Long, m As Long

    parts = Split(spec, ",")
    For i = 0 To UBound(parts)
        tok = Trim$(parts(i))
        If LenB(tok) = 0 Then GoTo nextTok

        dash = InStr(1, tok, "-", vbBinaryCompare)
        If dash = 0 Then
            If Not TryParseInt(tok, a) Then AddErr "ownership: bad token '" & tok & "' on line " & CStr(lineNo): GoTo nextTok
            b = a
        Else
            If Not TryParseInt(Left$(tok, dash - 1), a) Then AddErr "ownership: bad start in '" & tok & "' on line " & CStr(lineNo): GoTo nextTok
            If Not TryParseInt(mid$(tok, dash + 1), b) Then AddErr "ownership: bad end in '" & tok & "' on line " & CStr(lineNo): GoTo nextTok
            If b < a Then AddErr "ownership: reversed range '" & tok & "' on line " & CStr(lineNo): GoTo nextTok
        End If

        If a < 1 Or b > gMaxMap Then
            AddErr "ownership: out-of-bounds '" & tok & "' (1.." & CStr(gMaxMap) & ") line " & CStr(lineNo)
            GoTo nextTok
        End If

        For m = a To b
            If LenB(gOwner(m)) > 0 And StrComp(LCase$(gOwner(m)), LCase$(serverId), vbBinaryCompare) <> 0 Then
                AddErr "ownership overlap: map " & CStr(m) & " already owned by " & gOwner(m) & " (line " & CStr(lineNo) & ")"
            Else
                gOwner(m) = serverId
            End If
        Next m
nextTok:
    Next i
End Sub

Private Sub AddEndpoint(ByVal serverId As String, ByVal host As String, ByVal port As Long)
    Dim key As String
    key = LCase$(serverId)
    If gHosts.Exists(key) Then gHosts.Remove key
    If gPorts.Exists(key) Then gPorts.Remove key
    gHosts.Add key, host
    gPorts.Add key, port
End Sub

Private Sub AddErr(ByVal msg As String)
    If LenB(gErr) = 0 Then gErr = msg Else gErr = gErr & vbCrLf & msg
End Sub

Private Function TryParseInt(ByVal s As String, ByRef outVal As Long) As Boolean
    On Error GoTo fail
    outVal = CLng(Trim$(s))
    TryParseInt = True
    Exit Function
fail:
    TryParseInt = False
End Function


Private Function IniRead(ByVal filePath As String, ByVal section As String, ByVal key As String, ByVal defaultValue As String) As String
    Dim buf As String, n As Long
    buf = String$(1024, vbNullChar)
    n = GetPrivateProfileStringA(section, key, defaultValue, buf, Len(buf), filePath)
    If n > 0 Then IniRead = Left$(buf, n) Else IniRead = defaultValue
End Function

Private Function ReadAllText(ByVal path As String) As String
    Dim f As Integer, size As Long, content As String
    f = FreeFile
    Open path For Binary As #f
    size = LOF(f)
    If size > 0 Then
        content = String$(size, vbNullChar)
        Get #f, , content
    Else
        content = ""
    End If
    Close #f
    ReadAllText = content
End Function

Private Sub RaiseErr(ByVal where As String, ByVal msg As String)
    Err.raise vbObjectError + 410, where, msg
End Sub


