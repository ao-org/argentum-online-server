Attribute VB_Name = "DisconnectDiagnostics"
' Argentum 20 Game Server
'
' Lightweight disconnect diagnostics gated by PYMMO and debug_disconnects.

Option Explicit

Public Sub LogDisconnectDiag( _
    ByVal Source As String, _
    ByVal Action As String, _
    ByVal UserIndex As Integer, _
    ByVal ConnID As Long, _
    ByVal Reason As String, _
    Optional ByVal PacketId As Long = -1, _
    Optional ByVal PacketName As String = vbNullString, _
    Optional ByVal PacketCount As Long = -1, _
    Optional ByVal Extra As String = vbNullString)

    On Error Resume Next

    #If PYMMO = 0 Then
        Exit Sub
    #End If

    If Not IsFeatureEnabled("debug_disconnects") Then Exit Sub

    If PacketName = vbNullString And PacketId >= 0 Then
        PacketName = GetClientPacketName(PacketId)
    End If

    Dim message As String
    message = "disconnect_diag" & _
            " source=" & CleanDisconnectDiagValue(Source) & _
            " action=" & CleanDisconnectDiagValue(Action) & _
            " reason=" & CleanDisconnectDiagValue(Reason) & _
            " userIndex=" & CStr(UserIndex) & _
            " connID=" & CStr(ConnID)

    If PacketId >= 0 Then message = message & " packetId=" & CStr(PacketId)
    If PacketName <> vbNullString Then message = message & " packetName=" & CleanDisconnectDiagValue(PacketName)
    If PacketCount >= 0 Then message = message & " packetCount=" & CStr(PacketCount)

    If UserIndex > 0 Then
        With UserList(UserIndex)
            message = message & _
                    " name=" & CleanDisconnectDiagValue(.name) & _
                    " userId=" & CStr(.Id) & _
                    " accountId=" & CStr(.AccountID) & _
                    " ip=" & CleanDisconnectDiagValue(.ConnectionDetails.IP) & _
                    " logged=" & CStr(.flags.UserLogged) & _
                    " connValid=" & CStr(.ConnectionDetails.ConnIDValida) & _
                    " map=" & CStr(.pos.Map) & _
                    " x=" & CStr(.pos.x) & _
                    " y=" & CStr(.pos.y) & _
                    " idle=" & CStr(.Counters.IdleCount) & _
                    " saliendo=" & CStr(.Counters.Saliendo)
        End With
    End If

    If Extra <> vbNullString Then message = message & " extra=" & CleanDisconnectDiagValue(Extra)

    Call LogInfoServidor(message)
    Call AddLogToCircularBuffer(message)
End Sub

Public Sub LogDisconnectEvent( _
    ByVal Source As String, _
    ByVal Action As String, _
    ByVal UserIndex As Integer, _
    ByVal ConnID As Long, _
    Optional ByVal Reason As String = vbNullString, _
    Optional ByVal PacketId As Long = -1)

    Call LogDisconnectDiag(Source, Action, UserIndex, ConnID, Reason, PacketId)
End Sub

Public Function GetClientPacketName(ByVal PacketId As Long) As String
    On Error Resume Next

    If PacketId < ClientPacketID.eMinPacket Or PacketId >= ClientPacketID.PacketCount Then
        GetClientPacketName = "Invalid"
    Else
        GetClientPacketName = PacketID_to_string(PacketId)
        If GetClientPacketName = vbNullString Then GetClientPacketName = "Unknown"
    End If
End Function

Private Function CleanDisconnectDiagValue(ByVal Value As String) As String
    On Error Resume Next

    CleanDisconnectDiagValue = Replace(Value, vbCr, " ")
    CleanDisconnectDiagValue = Replace(CleanDisconnectDiagValue, vbLf, " ")
    CleanDisconnectDiagValue = Replace(CleanDisconnectDiagValue, vbTab, " ")
End Function
