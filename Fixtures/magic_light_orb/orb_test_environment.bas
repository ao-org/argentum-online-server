Attribute VB_Name = "OrbTestEnvironment"
' Copyright (C) 2026 Noland Studios LTD
' Licensed under the GNU Affero General Public License, version 3 or later.
' See licence.txt for the full license.
Option Explicit

' Isolated transport/map fixture. The production orb module is compiled unchanged.
Public Enum e_Class
    Mage = 1
    Cleric = 2
    Warrior = 3
    Assasin = 4
    Bard = 5
    Druid = 6
    Paladin = 7
End Enum
Public Enum e_Block
    ALL_SIDES = 15
End Enum
Public Enum e_TextChannel
    TEXTCHANNEL_COMBAT = 1
End Enum
Public Enum e_FontTypeNames
    FONTTYPE_New_Naranja = 1
End Enum
Public Const MSG_LIGHT_ORB_POSITION As Integer = 2296
Public Type t_WorldPos
    Map As Integer
    x As Byte
    y As Byte
End Type
Public Type t_Spell
    Duration As Integer
    LightOrbObject As Integer
    LightOrbCastRange As Byte
End Type
Public Type t_Object
    CreaLuz As String
End Type
Public Type t_Flags
    TargetMap As Integer
    TargetX As Byte
    TargetY As Byte
End Type
Public Type t_Connection
    ConnIDValida As Boolean
End Type
Public Type t_User
    clase As e_Class
    pos As t_WorldPos
    flags As t_Flags
    ConnectionDetails As t_Connection
End Type
Public Type t_MapTile
    Blocked As Byte
End Type
Public Type t_Group
    CountEntrys As Integer
    UserEntrys(1 To 2) As Integer
End Type
Public Type t_CapturedPacket
    Recipient As Integer
    InstanceId As Long
    Remaining As Long
    x As Byte
    y As Byte
End Type
Public UserList(1 To 2) As t_User
Public Hechizos(1 To 1) As t_Spell
Public ObjData(1 To 1) As t_Object
Public MapData(1 To 100, 1 To 100, 1 To 2) As t_MapTile
Public ConnGroups(1 To 2) As t_Group
Public Supported(1 To 2) As Boolean
Public TestNow As Long
Public Packets(1 To 100) As t_CapturedPacket
Public PacketCount As Integer
Public ErrorCount As Integer
Public FeedbackCount As Integer

Public Function UserSupportsMagicLightOrbs(ByVal UserIndex As Integer) As Boolean
    UserSupportsMagicLightOrbs = Supported(UserIndex)
End Function
Public Function InMapBounds(ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte) As Boolean
    InMapBounds = Map >= 1 And Map <= 2 And x >= 1 And x <= 100 And y >= 1 And y <= 100
End Function
Public Function GetTickCountRaw() As Long
    GetTickCountRaw = TestNow
End Function
Public Function TicksElapsed(ByVal LastTick As Long, ByVal NowRaw As Long) As Double
    Dim first As Double
    Dim second As Double
    first = LastTick
    second = NowRaw
    If first < 0 Then first = first + 4294967296#
    If second < 0 Then second = second + 4294967296#
    TicksElapsed = second - first
    If TicksElapsed < 0 Then TicksElapsed = TicksElapsed + 4294967296#
End Function
Public Function IntervalRemainingMs(ByVal LastTick As Long, ByVal IntervalMs As Long, ByVal NowRaw As Long) As Long
    Dim elapsed As Double
    elapsed = TicksElapsed(LastTick, NowRaw)
    If elapsed < IntervalMs Then IntervalRemainingMs = CLng(IntervalMs - elapsed)
End Function
Public Sub WriteLocaleMsg(ByVal UserIndex As Integer, ByVal MessageId As Integer, ByVal Channel As e_TextChannel, ByVal Font As e_FontTypeNames, ByVal Parameter As String)
    FeedbackCount = FeedbackCount + 1
End Sub
Public Sub WriteHooMagicLightOrb(ByVal UserIndex As Integer, ByVal InstanceId As Long, ByVal Map As Integer, ByVal x As Byte, ByVal y As Byte, ByVal ProfileObject As Integer, ByVal RemainingMs As Long)
    PacketCount = PacketCount + 1
    With Packets(PacketCount)
        .Recipient = UserIndex
        .InstanceId = InstanceId
        .Remaining = RemainingMs
        .x = x
        .y = y
    End With
End Sub
Public Sub TraceError(ByVal Number As Long, ByVal Description As String, ByVal Procedure As String, ByVal LineNumber As Long)
    ErrorCount = ErrorCount + 1
End Sub
