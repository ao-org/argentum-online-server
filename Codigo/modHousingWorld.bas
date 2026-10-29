Attribute VB_Name = "modHousingWorld"
Option Explicit

Private Type t_HousingDoor
    Key As Integer
    Map As Integer
    X1 As Integer
    Y1 As Integer
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
Private doors() As t_HousingDoor
Private doorCount As Long
Private mappedKeys As Object
Private plans() As t_HousingKeyPlan
Private planCount As Long
Private ready As Boolean

Public Function HousingInitialize(ByVal connectionString As String) As Boolean
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
    Call LoadDoors()
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

Private Sub LoadDoors()
    ' Discover actual entrances from existing map objects and catalogue lock codes.
    ' No interior boundaries are inferred from doors, exits or map numbers.
    Dim m As Long, x As Long, y As Long, obj As Long
    Dim door As t_HousingDoor
    doorCount = 0
    ReDim doors(1 To 4096)
    For m = 1 To NumMaps
        If MapInfo(m).MapResource <> 0 Then
            For y = 1 To 100
                For x = 1 To 100
                    If InMapBounds(CInt(m), CInt(x), CInt(y)) Then
                        obj = MapData(x, y, m).ObjInfo.ObjIndex
                        If obj > 0 And obj <= UBound(doorKeys) Then
                            If doorKeys(obj) <> 0 Then
                                If doorCount = 4096 Then Err.Raise 5, "Housing", "Too many house doors"
                                door.Key = doorKeys(obj): door.Map = m
                                door.X1 = x: door.Y1 = y
                                door.LockedObject = ObjData(obj).IndexCerradaLlave
                                If Not InMapBounds(CInt(m), CInt(x - 1), CInt(y + 1)) Then Err.Raise 5, "Housing", "Invalid door threshold"
                                If Not ValidDoor(door) Then Err.Raise 5, "Housing", "Invalid house door variants"
                                doorCount = doorCount + 1: doors(doorCount) = door
                                mappedKeys(CStr(door.Key)) = True
                            End If
                        End If
                    End If
                Next
            Next
        End If
    Next
End Sub

Private Function ValidDoor(ByRef r As t_HousingDoor) As Boolean
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

Public Function HousingCanVacate(ByVal key As Integer) As Boolean
    Dim u As Long, i As Long, x As Long, y As Long
    For i = 1 To doorCount
        With doors(i)
            If .Key = key Then
                If Not ValidDoor(doors(i)) Then Exit Function
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
    For i = 1 To doorCount
        With doors(i)
            If .Key = key Then
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
