Attribute VB_Name = "DisconnectDiagnostics"
' Argentum 20 Game Server
'
' Lightweight disconnect diagnostics gated by the debug_disconnects feature flag.

Option Explicit

Public Sub LogDisconnectEvent( _
    ByVal Source As String, _
    ByVal Action As String, _
    ByVal UserIndex As Integer, _
    ByVal ConnID As Long, _
    Optional ByVal Reason As String = vbNullString, _
    Optional ByVal PacketId As Long = -1)

    On Error Resume Next

    If Not IsFeatureEnabled("debug_disconnects") Then Exit Sub

    Dim message As String
    message = "disconnect_diag" & _
            " source=" & Source & _
            " action=" & Action & _
            " userIndex=" & CStr(UserIndex) & _
            " connID=" & CStr(ConnID)

    If Reason <> vbNullString Then
        message = message & " reason=" & Replace(Reason, vbCrLf, " ")
    End If

    If PacketId >= 0 Then
        message = message & " packetId=" & CStr(PacketId)
    End If

    If UserIndex > 0 Then
        With UserList(UserIndex)
            message = message & _
                    " username=" & .name & _
                    " userId=" & CStr(.Id) & _
                    " accountId=" & CStr(.AccountID) & _
                    " ip=" & .ConnectionDetails.IP & _
                    " userLogged=" & CStr(.flags.UserLogged) & _
                    " connIDValida=" & CStr(.ConnectionDetails.ConnIDValida) & _
                    " map=" & CStr(.pos.Map) & _
                    " x=" & CStr(.pos.x) & _
                    " y=" & CStr(.pos.y) & _
                    " idleCount=" & CStr(.Counters.IdleCount) & _
                    " saliendo=" & CStr(.Counters.Saliendo) & _
                    " salir=" & CStr(.Counters.Salir)
        End With
    End If

    Call LogInfoServidor(message)
    Call AddLogToCircularBuffer(message)
End Sub
