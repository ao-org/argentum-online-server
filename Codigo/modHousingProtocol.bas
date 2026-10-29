Attribute VB_Name = "modHousingProtocol"
Option Explicit

Private Type t_HousingSession
    Version As Long
    Account As Long
    Highest As Long
    Payload As String
    Reply As t_HousingReply
    LastTick As Long
    HasTick As Boolean
    NeedsStatus As Boolean
End Type
Private sessions() As t_HousingSession
Private capacity As Long

Public Sub HousingResetSession(ByVal userIndex As Integer)
    If userIndex < 1 Or userIndex > capacity Then Exit Sub
    Dim blankReply As t_HousingSession
    sessions(userIndex) = blankReply
End Sub

Public Function HousingProcessRequest(ByVal userIndex As Integer, ByRef request As t_HousingRequest, ByVal tick As Long, ByRef reply As t_HousingReply) As e_HousingCode
    On Error GoTo Failed
    Dim blankReply As t_HousingReply, emptySession As t_HousingSession, payload As String
    reply = blankReply
    reply.RequestId = request.RequestId: reply.Operation = request.Operation
    reply.Code = HousingInvalid
    If userIndex < 1 Or userIndex > UBound(UserList) Then GoTo Done
    If Not UserList(userIndex).flags.UserLogged Or UserList(userIndex).AccountID <= 0 Then GoTo Done
    If request.RequestId <= 0 Or request.Operation > HousingExchange Or request.ExpectedKey < 0 Or request.TargetKey < 0 Or request.Cursor < 0 Or Len(request.Search) > 96 Then GoTo Done
    If InStr(request.Search, vbNullChar) > 0 Then GoTo Done
    If capacity <> UBound(UserList) Then
        capacity = UBound(UserList)
        ReDim Preserve sessions(1 To capacity)
    End If
    payload = CStr(request.Operation) & ":" & CStr(request.ExpectedKey) & ":" & CStr(request.TargetKey) & ":" & CStr(request.Cursor) & ":" & request.Search
    With sessions(userIndex)
        If .Version <> UserList(userIndex).VersionId Or .Account <> UserList(userIndex).AccountID Then
            sessions(userIndex) = emptySession
            .Version = UserList(userIndex).VersionId: .Account = UserList(userIndex).AccountID
        End If
        If request.RequestId <= .Highest Then
            If request.RequestId = .Highest And payload = .Payload Then reply = .Reply
            GoTo Done
        End If
        ' Flood control is separate from the database's 72-hour ownership rule.
        .Highest = request.RequestId: .Payload = payload
        If .HasTick Then
            If CDbl(tick) - CDbl(.LastTick) >= 0 And CDbl(tick) - CDbl(.LastTick) < 500 Then
                reply.Code = HousingBusy
                .Reply = reply
                GoTo Done
            End If
        End If
        .HasTick = True: .LastTick = tick
        If .NeedsStatus And request.Operation <> HousingCatalogue Then
            reply.Code = HousingUnresolved
        Else
            reply.Code = HousingDbOperation(.Account, request, reply)
            If reply.Code = HousingUnresolved Then
                .NeedsStatus = True
            ElseIf request.Operation = HousingCatalogue And reply.Code = HousingOkay Then
                .NeedsStatus = False
            End If
        End If
        .Reply = reply
    End With
Done:
    HousingProcessRequest = reply.Code
    Exit Function
Failed:
    reply.Code = HousingUnresolved
    If userIndex >= 1 And userIndex <= capacity Then
        sessions(userIndex).Reply = reply
        sessions(userIndex).NeedsStatus = True
    End If
    Call LogError("Housing request failed; refresh status before another mutation")
    HousingProcessRequest = reply.Code
End Function

Public Sub HandleHousingRequest(ByVal userIndex As Integer, ByRef packetReader As Network.reader)
    On Error GoTo Failed
    Dim request As t_HousingRequest, reply As t_HousingReply
    request.RequestId = packetReader.ReadInt32()
    request.Operation = packetReader.ReadInt8()
    request.ExpectedKey = packetReader.ReadInt16()
    request.TargetKey = packetReader.ReadInt16()
    request.Cursor = packetReader.ReadInt16()
    request.Search = packetReader.ReadString8()
    Call HousingProcessRequest(userIndex, request, GetTickCountRaw(), reply)
    Call WriteHousingResult(userIndex, reply)
    Exit Sub
Failed:
    Call CloseSocket(userIndex)
End Sub

Public Sub WriteHousingResult(ByVal userIndex As Integer, ByRef reply As t_HousingReply)
    On Error GoTo Failed
    Dim packetWriter As Network.writer, i As Long
    Set packetWriter = Protocol_Writes.GetWriterBuffer()
    Call packetWriter.WriteInt16(ServerPacketID.eHousingResult)
    Call packetWriter.WriteInt32(reply.RequestId)
    Call packetWriter.WriteInt8(reply.Operation)
    Call packetWriter.WriteInt8(CByte(reply.Code))
    Call packetWriter.WriteInt16(reply.Current.Key)
    Call packetWriter.WriteString8(reply.Current.Description)
    Call packetWriter.WriteString8(reply.AssignedUTC)
    Call packetWriter.WriteString8(reply.EligibleUTC)
    Call packetWriter.WriteString8(reply.ServerUTC)
    Call packetWriter.WriteBool(reply.CanAcquire)
    Call packetWriter.WriteBool(reply.CanRelease)
    Call packetWriter.WriteInt16(reply.NextCursor)
    Call packetWriter.WriteInt8(reply.Count)
    For i = 1 To reply.Count
        Call packetWriter.WriteInt16(reply.Houses(i).Key)
        Call packetWriter.WriteString8(reply.Houses(i).Description)
    Next
    Call modSendData.SendData(ToIndex, userIndex)
    Exit Sub
Failed:
    If Not packetWriter Is Nothing Then Call packetWriter.Clear()
    ' Permission changes never depend on successful UI notification.
    Call CloseSocket(userIndex)
End Sub
