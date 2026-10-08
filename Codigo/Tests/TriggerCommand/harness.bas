Attribute VB_Name = "TriggerHarness"
' Argentum 20 Game Server
' Copyright (C) 2026 Noland Studios LTD
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit

Private Declare Sub ExitProcess Lib "kernel32" (ByVal exitCode As Long)

Public Type TestPosition
    Map As Integer
    x As Integer
    y As Integer
End Type
Public Type TestFlags
    Privilegios As e_PlayerType
    UserLogged As Boolean
End Type
Public Type TestUser
    pos As TestPosition
    flags As TestFlags
End Type
Public Type TestTile
    trigger As Long
End Type
Public Enum e_TextChannel
    TEXTCHANNEL_SYSTEM = 1
End Enum
Public Enum e_FontTypeNames
    FONTTYPE_INFO = 1
End Enum
Public Type TestMap
    TileOverrides As Dictionary
End Type
Public Const MSG_TRIGGER As Long = 1498
Public LastUser As Integer
Public MapInfo(1 To 1) As TestMap
Public UserList(1 To 1) As TestUser
Public MapData(1 To 1, 1 To 1, 1 To 1) As TestTile
Public reader As Network.reader
Private PacketBuffer() As Byte
Private PacketWriter As Network.Writer
Private Feedback As String
Private AuditCount As Long
Private PublicationCount As Long
Private ErrorCount As Long
Private Checks As Long
Private Failures As Long
Private Results As String

Public Sub Main()
    On Error GoTo Main_Err
    Dim values As Variant
    Dim roles As Variant
    Dim value As Variant
    Dim role As Variant
    UserList(1).pos.Map = 1
    UserList(1).pos.x = 1
    UserList(1).pos.y = 1
    values = Array(0&, 255&, 256&, 511&, 512&, 1024&, 2048&, 4096&, 8192&, 16383&)
    LastUser = 1
    UserList(1).flags.UserLogged = True
    roles = Array(e_PlayerType.Dios, e_PlayerType.Admin)
    For Each role In roles
        UserList(1).flags.Privilegios = role
        For Each value In values
            Call ResetObservations()
            Call LoadSetPacket(CLng(value))
            Call HandleSetTrigger(1)
            Call Check(MapData(1, 1, 1).trigger = CLng(value), "set preserves Long " & value & "; actual=" & MapData(1, 1, 1).trigger)
            Call Check(Feedback = "Trigger " & value & " on the map 1 1,1", "set confirmation " & value)
            Call Check(AuditCount = 1 And ErrorCount = 0 And PublicationCount = 1, "accepted set audited and published once")
            Call Check(reader.GetAvailable() = 2, "set consumes exactly four payload bytes")
            Call Check(reader.ReadInt16() = 165, "next query packet remains intact")
            Call ResetObservations()
            Call HandleAskTrigger(1)
            Call Check(Feedback = "MAP 1,1,1. = " & value, "query preserves Long " & value)
            Call Check(AuditCount = 1 And ErrorCount = 0, "query audited without errors")
        Next value
    Next role

    roles = Array(e_PlayerType.User, e_PlayerType.Consejero, e_PlayerType.SemiDios, e_PlayerType.RoleMaster, 0, 64, e_PlayerType.Admin Or e_PlayerType.User)
    For Each role In roles
        UserList(1).flags.Privilegios = role
        MapData(1, 1, 1).trigger = 256
        Call ResetObservations()
        Call LoadSetPacket(16383)
        Call HandleSetTrigger(1)
        Call HandleAskTrigger(1)
        Call Check(MapData(1, 1, 1).trigger = 256, "unauthorized role cannot set: " & role)
        Call Check(Feedback = "" And AuditCount = 0 And ErrorCount = 0 And PublicationCount = 0, "unauthorized role cannot query: " & role)
        Call Check(reader.GetAvailable() = 2, "unauthorized set still consumes payload")
    Next role

    UserList(1).flags.Privilegios = e_PlayerType.Admin
    values = Array(-1&, 16384&, 32768&, 65536&, 16909060&, 2147483647)
    For Each value In values
        MapData(1, 1, 1).trigger = 256
        Call ResetObservations()
        Call LoadSetPacket(CLng(value))
        Call HandleSetTrigger(1)
        Call Check(MapData(1, 1, 1).trigger = 256, "negative or unknown mask rejected: " & value)
        Call Check(Feedback = "" And AuditCount = 0 And ErrorCount = 0 And PublicationCount = 0, "invalid mask not published")
        Call Check(reader.GetAvailable() = 2 And reader.ReadInt16() = 165, "rejected mask consumes exactly four bytes")
    Next value

    ' Query transport preserves the complete Long even for an injected noncanonical value.
    MapData(1, 1, 1).trigger = 65536
    Call ResetObservations()
    Call HandleAskTrigger(1)
    Call Check(Feedback = "MAP 1,1,1. = 65536", "query does not truncate a Long")
    Call Check(AuditCount = 1 And ErrorCount = 0, "query audited without errors")

    Dim truncated() As Byte, size As Integer
    For size = 1 To 3
        Call ResetObservations()
        ReDim truncated(0 To size - 1)
        truncated(0) = 9
        Set reader = New Network.reader
        Call reader.SetData(truncated)
        Call HandleSetTrigger(1)
        Call Check(MapData(1, 1, 1).trigger = 65536, "truncated payload cannot change tile")
        Call Check(ErrorCount = 1 And AuditCount = 0 And Feedback = "" And PublicationCount = 0, "truncated payload reports error without publication")
    Next size
    Call Finish()
    Exit Sub
Main_Err:
    Failures = Failures + 1
    Results = Results & "FATAL: " & Err.Description & vbCrLf
    Call Finish()
End Sub

Private Sub LoadSetPacket(ByVal value As Long)
    Set PacketWriter = New Network.Writer
    Call PacketWriter.WriteInt16(164)
    Call PacketWriter.WriteInt32(value)
    Call PacketWriter.WriteInt16(165)
    Call PacketWriter.GetData(PacketBuffer)
    Set reader = New Network.reader
    Call reader.SetData(PacketBuffer)
    Call Check(reader.ReadInt16() = 164, "set packet ID")
End Sub

Private Sub ResetObservations()
    Feedback = ""
    AuditCount = 0
    PublicationCount = 0
    ErrorCount = 0
End Sub

Private Sub Check(ByVal passed As Boolean, ByVal name As String)
    Checks = Checks + 1
    If Not passed Then
        Failures = Failures + 1
        Results = Results & "FAIL: " & name & vbCrLf
    End If
End Sub

Private Sub Finish()
    Dim report As Integer
    report = FreeFile()
    Open App.Path & "\results.txt" For Output As #report
    Print #report, Results & Checks & " checks; " & Failures & " failures"
    Close #report
    Call ExitProcess(Failures)
End Sub

Public Function GetUserRealName(ByVal userIndex As Integer) As String
    GetUserRealName = "TestGM"
End Function

Public Sub LogGM(ByVal name As String, ByVal message As String)
    AuditCount = AuditCount + 1
End Sub

Public Sub WriteConsoleMsg(ByVal userIndex As Integer, ByVal message As String, ByVal channel As Long, ByVal font As Long)
    Feedback = message
End Sub

Public Sub WriteLocaleMsg(ByVal userIndex As Integer, ByVal messageId As Long, ByVal channel As Long, ByVal font As Long, ByVal message As String)
    Feedback = message
End Sub

Public Sub TraceError(ByVal number As Long, ByVal description As String, ByVal source As String, ByVal line As Long)
    ErrorCount = ErrorCount + 1
End Sub

Public Function InMapBounds(ByVal map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    InMapBounds = map = 1 And x = 1 And y = 1
End Function

Public Sub WriteHooTileProperties(ByVal userIndex As Integer, ByVal map As Integer, ByVal x As Integer, ByVal y As Integer, ByVal flags As Long)
    PublicationCount = PublicationCount + 1
End Sub
