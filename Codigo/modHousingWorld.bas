Attribute VB_Name = "modHousingWorld"
Option Explicit

Private Type t_HousingRegion
    Key As Integer
    Door As Boolean
    Map As Integer
    Resource As Integer
    X1 As Integer
    Y1 As Integer
    X2 As Integer
    Y2 As Integer
    LockedObject As Integer
End Type
Private Type t_HousingKeyPlan
    UserIndex As Integer
    Version As Long
    Keys(1 To MAXKEYS) As Integer
End Type
Private ordinaryOwners As Object
Private blockedAccounts As Object
Private doorKeys() As Integer
Private regions() As t_HousingRegion
Private regionCount As Long
Private mappedKeys As Object
Private plans() As t_HousingKeyPlan
Private planCount As Long
Private ready As Boolean

Public Function HousingInitialize(ByVal connectionString As String, ByVal regionFile As String) As Boolean
    On Error GoTo Failed
    ready = False
    Set ordinaryOwners = CreateObject("Scripting.Dictionary")
    Set blockedAccounts = CreateObject("Scripting.Dictionary")
    Set mappedKeys = CreateObject("Scripting.Dictionary")
    If Not HousingDbOpen(connectionString) Then Exit Function
    Dim rs As ADODB.Recordset, key As Long, owner As Long, obj As Long, code As Long
    Dim lockKeys As Object
    Set lockKeys = CreateObject("Scripting.Dictionary")
    Set rs = HousingQuery("SELECT h.key_obj,h.account_id FROM house_key h WHERE NOT EXISTS (SELECT 1 FROM castle c WHERE c.inside_key_obj_id=h.key_obj) LIMIT 32768")
    Do While Not rs.EOF
        key = CLng(rs.Fields(0).Value): owner = 0
        If key < 1 Or key > UBound(ObjData) Or key > 32767 Then Err.Raise 5, "Housing", "Invalid catalogue key"
        If ObjData(key).OBJType <> e_OBJType.otKeys Or ObjData(key).clave = 0 Then Err.Raise 5, "Housing", "Invalid key object"
        code = ObjData(key).clave
        If lockKeys.Exists(CStr(code)) Then Err.Raise 5, "Housing", "Ambiguous lock code"
        Call lockKeys.Add(CStr(code), CInt(key))
        If Not IsNull(rs.Fields(1).Value) Then
            owner = CLng(rs.Fields(1).Value)
            If owner <= 0 Then Err.Raise 5, "Housing", "Invalid catalogue owner"
        End If
        Call ordinaryOwners.Add(CStr(key), owner)
        Call rs.MoveNext()
    Loop
    Call rs.Close()
    ReDim doorKeys(LBound(ObjData) To UBound(ObjData))
    For obj = LBound(ObjData) To UBound(ObjData)
        If ObjData(obj).OBJType = e_OBJType.otDoors Then
            If lockKeys.Exists(CStr(ObjData(obj).clave)) Then doorKeys(obj) = lockKeys(CStr(ObjData(obj).clave))
        End If
    Next
    Call LoadRegions(regionFile)
    ready = True
    HousingInitialize = True
    Exit Function
Failed:
    Dim failure As String
    failure = Err.Description
    Call LogError("Housing initialization unavailable: " & failure)
    Call HousingDbClose()
#If UNIT_TEST = 1 Then
    Err.Raise vbObjectError + 731, "Housing initialization", failure
#End If
End Function

Private Sub LoadRegions(ByVal path As String)
    On Error GoTo Failed
    Dim f As Integer, opened As Boolean, line As String, fields() As String
    Dim r As t_HousingRegion, i As Long, j As Long, hasDoor As Boolean, hasInterior As Boolean
    regionCount = 0
    ReDim regions(1 To 4096)
    f = FreeFile
    Open path For Input As #f
    opened = True
    Do While Not EOF(f)
        Line Input #f, line
        line = Trim$(line)
        If Len(line) > 0 Then
            If Left$(line, 1) <> "#" Then
                fields = Split(line, ",")
                If UBound(fields) <> 8 Or regionCount = 4096 Then Err.Raise 5, "Housing", "Invalid region row"
                r.Key = CInt(fields(0)): r.Door = (fields(1) = "D")
                If fields(1) <> "D" And fields(1) <> "R" Then Err.Raise 5
                r.Map = CInt(fields(2)): r.Resource = CInt(fields(3))
                r.X1 = CInt(fields(4)): r.Y1 = CInt(fields(5))
                r.X2 = CInt(fields(6)): r.Y2 = CInt(fields(7)): r.LockedObject = CInt(fields(8))
                If Not ordinaryOwners.Exists(CStr(r.Key)) Then Err.Raise 5, "Housing", "Region key missing from ordinary catalogue"
                If Not InMapBounds(r.Map, r.X1, r.Y1) Or Not InMapBounds(r.Map, r.X2, r.Y2) Then Err.Raise 5, "Housing", "Region outside map"
                If MapInfo(r.Map).MapResource <> r.Resource Or r.X1 > r.X2 Or r.Y1 > r.Y2 Then Err.Raise 5, "Housing", "Invalid map mapping"
                If r.Door Then
                    If r.X1 <> r.X2 Or r.Y1 <> r.Y2 Then Err.Raise 5
                    If Not InMapBounds(r.Map, r.X1 - 1, r.Y1 + 1) Then Err.Raise 5
                    If Not ValidDoor(r) Then Err.Raise 5, "Housing", "Invalid door variants"
                ElseIf r.LockedObject <> 0 Then
                    Err.Raise 5
                End If
                For j = 1 To regionCount
                    If Not r.Door And Not regions(j).Door And r.Map = regions(j).Map And r.Key <> regions(j).Key Then
                        If r.X1 <= regions(j).X2 And r.X2 >= regions(j).X1 And r.Y1 <= regions(j).Y2 And r.Y2 >= regions(j).Y1 Then Err.Raise 5, "Housing", "Overlapping property regions"
                    End If
                Next
                regionCount = regionCount + 1: regions(regionCount) = r
            End If
        End If
    Loop
    Close #f
    opened = False
    For i = 1 To regionCount
        hasDoor = False: hasInterior = False
        For j = 1 To regionCount
            If regions(j).Key = regions(i).Key Then
                If regions(j).Door Then hasDoor = True Else hasInterior = True
            End If
        Next
        If Not hasDoor Or Not hasInterior Then Err.Raise 5, "Housing", "Incomplete property mapping"
        mappedKeys(CStr(regions(i).Key)) = True
    Next
    ' Every installed door using these keys must be represented, not merely one entrance.
    Dim m As Long, x As Long, y As Long, obj As Long, found As Boolean
    For m = 1 To NumMaps
        If MapInfo(m).MapResource <> 0 Then
            For y = 1 To 100
                For x = 1 To 100
                    If InMapBounds(CInt(m), CInt(x), CInt(y)) Then
                        obj = MapData(x, y, m).ObjInfo.ObjIndex
                        If obj > 0 And obj <= UBound(doorKeys) Then
                            If mappedKeys.Exists(CStr(doorKeys(obj))) Then
                                found = False
                                For i = 1 To regionCount
                                    If regions(i).Door And regions(i).Map = m And regions(i).X1 = x And regions(i).Y1 = y And regions(i).Key = doorKeys(obj) Then found = True
                                Next
                                If Not found Then Err.Raise 5, "Housing", "Unmapped additional door"
                            End If
                        End If
                    End If
                Next
            Next
        End If
    Next
    Exit Sub
Failed:
    Dim number As Long, detail As String
    number = Err.Number: detail = Err.Description
    If opened Then Close #f
    Err.Raise number, "Housing regions", detail
End Sub

Private Function ValidDoor(ByRef r As t_HousingRegion) As Boolean
    Dim ids(1 To 3) As Integer, i As Long, obj As Integer
    obj = r.LockedObject
    If obj < 1 Or obj > UBound(ObjData) Then Exit Function
    ids(1) = ObjData(obj).IndexAbierta: ids(2) = ObjData(obj).IndexCerrada: ids(3) = obj
    For i = 1 To 3
        obj = ids(i)
        If obj < 1 Or obj > UBound(ObjData) Then Exit Function
        If ObjData(obj).OBJType <> e_OBJType.otDoors Or ObjData(obj).Subtipo <> 0 Then Exit Function
        If ObjData(obj).clave <> ObjData(r.Key).clave Then Exit Function
        If ObjData(obj).IndexAbierta <> ids(1) Or ObjData(obj).IndexCerrada <> ids(2) Or ObjData(obj).IndexCerradaLlave <> ids(3) Then Exit Function
    Next
    If ObjData(ids(1)).Cerrada <> 0 Or ObjData(ids(2)).Cerrada <> 1 Or ObjData(ids(2)).Llave <> 0 Or ObjData(ids(3)).Cerrada <> 1 Or ObjData(ids(3)).Llave <> 1 Then Exit Function
    obj = MapData(r.X1, r.Y1, r.Map).ObjInfo.ObjIndex
    ValidDoor = (obj = ids(1) Or obj = ids(2) Or obj = ids(3))
End Function

Public Function HousingMapped(ByVal key As Integer) As Boolean
    If mappedKeys Is Nothing Then Exit Function
    HousingMapped = mappedKeys.Exists(CStr(key))
End Function

Public Function HousingOrdinaryKey(ByVal key As Integer) As Boolean
    If ordinaryOwners Is Nothing Then Exit Function
    HousingOrdinaryKey = ordinaryOwners.Exists(CStr(key))
End Function

Public Function HousingDoorAccess(ByVal userIndex As Integer, ByVal doorObject As Integer) As Boolean
    On Error GoTo Failed
    If Not ready Then Exit Function
    HousingDoorAccess = True
    If doorObject < LBound(doorKeys) Or doorObject > UBound(doorKeys) Then Exit Function
    Dim key As Integer
    key = doorKeys(doorObject)
    If key = 0 Then Exit Function
    HousingDoorAccess = False
    If blockedAccounts.Exists(CStr(UserList(userIndex).AccountID)) Then Exit Function
    HousingDoorAccess = UserList(userIndex).AccountID > 0 And ordinaryOwners(CStr(key)) = UserList(userIndex).AccountID
    Exit Function
Failed:
    HousingDoorAccess = False
End Function

Public Function HousingPrepareKeys(ByVal accountId As Long, ByVal newKey As Integer) As Boolean
    Dim u As Long, slot As Long, freeSlot As Long, key As Integer
    planCount = 0
    ReDim plans(1 To UBound(UserList))
    For u = LBound(UserList) To UBound(UserList)
        ' Includes retained logout/countdown users and partially loaded characters.
        If UserList(u).AccountID = accountId And UserList(u).Id > 0 Then
            planCount = planCount + 1
            plans(planCount).UserIndex = CInt(u): plans(planCount).Version = UserList(u).VersionId
            freeSlot = 0
            For slot = 1 To MAXKEYS
                key = UserList(u).Keys(slot)
                If HousingOrdinaryKey(key) Then key = 0
                plans(planCount).Keys(slot) = key
                If key = 0 And freeSlot = 0 Then freeSlot = slot
            Next
            If newKey <> 0 Then
                If freeSlot = 0 Then Exit Function
                plans(planCount).Keys(freeSlot) = newKey
            End If
        End If
    Next
    HousingPrepareKeys = True
End Function

Public Sub HousingApplyKeys(ByVal accountId As Long, ByVal newKey As Integer)
    Dim i As Long, slot As Long, u As Integer, key As Variant
    For Each key In ordinaryOwners.Keys
        If ordinaryOwners(key) = accountId Then ordinaryOwners(key) = 0
    Next
    If newKey <> 0 Then ordinaryOwners(CStr(newKey)) = accountId
    If blockedAccounts.Exists(CStr(accountId)) Then Call blockedAccounts.Remove(CStr(accountId))
    ' Update all authorization first; network delivery is not permission authority.
    For i = 1 To planCount
        u = plans(i).UserIndex
        If UserList(u).VersionId <> plans(i).Version Or UserList(u).AccountID <> accountId Then Err.Raise 5, "Housing", "Loaded character changed"
        For slot = 1 To MAXKEYS
            UserList(u).Keys(slot) = plans(i).Keys(slot)
        Next
    Next
    For i = 1 To planCount
        u = plans(i).UserIndex
        If UserList(u).flags.UserLogged Then
            For slot = 1 To MAXKEYS
                Call WriteUpdateUserKey(u, CByte(slot), UserList(u).Keys(slot))
            Next
        End If
    Next
End Sub

Public Sub HousingBlockAccount(ByVal accountId As Long)
    Dim u As Long, slot As Long
    blockedAccounts(CStr(accountId)) = True
    For u = LBound(UserList) To UBound(UserList)
        If UserList(u).AccountID = accountId Then
            For slot = 1 To MAXKEYS
                If HousingOrdinaryKey(UserList(u).Keys(slot)) Then
                    UserList(u).Keys(slot) = 0
                    If UserList(u).flags.UserLogged Then Call WriteUpdateUserKey(CInt(u), CByte(slot), 0)
                End If
            Next
        End If
    Next
    Call LogError("Housing reconciliation unresolved; ordinary key access blocked until authoritative status/login refresh")
End Sub

Public Function HousingInteriorKey(ByRef pos As t_WorldPos) As Integer
    Dim i As Long
    For i = 1 To regionCount
        With regions(i)
            If Not .Door And .Map = pos.Map Then
                If pos.x >= .X1 And pos.x <= .X2 And pos.y >= .Y1 And pos.y <= .Y2 Then
                    HousingInteriorKey = .Key
                    Exit Function
                End If
            End If
        End With
    Next
End Function

Public Function HousingCanVacate(ByVal key As Integer) As Boolean
    Dim u As Long, i As Long, x As Long, y As Long
    For u = LBound(UserList) To UBound(UserList)
        If UserList(u).Id > 0 Then
            If HousingInteriorKey(UserList(u).pos) = key Then Exit Function
        End If
    Next
    For i = 1 To regionCount
        With regions(i)
            If .Key = key And .Door Then
                If Not ValidDoor(regions(i)) Then Exit Function
                For y = .Y1 To .Y1 + 1
                    For x = .X1 - 1 To .X1
                        If MapData(x, y, .Map).UserIndex <> 0 Or MapData(x, y, .Map).NpcIndex <> 0 Then Exit Function
                        For u = LBound(UserList) To UBound(UserList)
                            If UserList(u).Id > 0 And UserList(u).pos.Map = .Map And UserList(u).pos.x = x And UserList(u).pos.y = y Then Exit Function
                        Next
                    Next
                Next
            End If
        End With
    Next
    HousingCanVacate = True
End Function

Public Sub HousingCloseDoors(ByVal key As Integer)
    Dim i As Long
    For i = 1 To regionCount
        With regions(i)
            If .Key = key And .Door Then
                MapData(.X1, .Y1, .Map).ObjInfo.ObjIndex = .LockedObject
                Call BloquearPuerta(.Map, .X1, .Y1, True)
                Call modSendData.SendToAreaByPos(.Map, .X1, .Y1, PrepareMessageObjectCreate(.LockedObject, MapData(.X1, .Y1, .Map).ObjInfo.amount, .X1, .Y1))
            End If
        End With
    Next
End Sub

Public Function HousingRefreshLogin(ByVal userIndex As Integer) As Boolean
    Dim request As t_HousingRequest, reply As t_HousingReply
    request.Operation = HousingCatalogue
    HousingRefreshLogin = (HousingDbOperation(UserList(userIndex).AccountID, request, reply) = HousingOkay)
End Function

Public Function HousingLoginNeedsHome(ByVal userIndex As Integer) As Boolean
    Dim key As Integer
    key = HousingInteriorKey(UserList(userIndex).pos)
    If key = 0 Then Exit Function
    HousingLoginNeedsHome = (ordinaryOwners(CStr(key)) <> UserList(userIndex).AccountID)
End Function
