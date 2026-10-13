Attribute VB_Name = "modSendData"
' Capture-only transport. Real Aurora Writer is used; no network is opened.
Option Explicit
Public Enum SendTarget
    ToIndex = 1
    ToPCAliveArea = 2
    ToNPCAliveArea = 3
End Enum
Public Sub SendData(ByVal Route As Integer, ByVal Index As Integer, Optional ByVal Prepared As Variant)
    Call CapturePacket()
End Sub
Public Sub SendToAreaByPos(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal Prepared As Variant)
    Call CapturePacket()
End Sub
Private Sub CapturePacket()
    Dim bytes() As Byte, value As Variant
    Call Writer.GetData(bytes)
    Call Writer.Clear()
    If bytes(LBound(bytes)) = 192 Then
        Call Expect(UBound(bytes) - LBound(bytes) + 1 = 16, "fixed packet length")
        If bytes(2) = eTrapUpsert Then
            If Not MapData(bytes(9), bytes(10), bytes(7)).Trap Is Nothing Then
                Call Expect(MapData(bytes(9), bytes(10), bytes(7)).Trap.RemainingMilliseconds = Read32(bytes, 12), "fully initialized before publication")
            End If
        End If
        value = bytes
        Call Packets.Add(value)
    End If
End Sub
