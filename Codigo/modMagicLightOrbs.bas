Attribute VB_Name = "modMagicLightOrbs"
' Argentum 20 Game Server
' Copyright (C) 2026 Noland Studios LTD
' Licensed under the GNU Affero General Public License, version 3 or later.
' See licence.txt for the full license.
Option Explicit

Private Const ORB_UPDATE_INTERVAL_MS As Long = 100
Private Const MAX_ORB_DURATION_SECONDS As Long = 3600
Private Const MAX_ORB_INSTANCE_ID As Long = &H7FFFFFFF

Private Type t_MagicLightOrb
    InstanceId As Long
    Position As t_WorldPos
    ProfileObject As Integer
    CreatedAt As Long
    DurationMs As Long
End Type

Private UserOrbs() As t_MagicLightOrb
Private OrbsInitialized As Boolean
Private NextInstanceId As Long
Private LastUpdate As Long

Public Function CanClassCastLightOrb(ByVal ClassId As e_Class) As Boolean
    CanClassCastLightOrb = (ClassId = Mage Or ClassId = Druid Or ClassId = Bard)
End Function

Public Function IsMagicLightOrbPositionValid(ByVal Map As Integer, ByVal CasterX As Byte, ByVal CasterY As Byte, ByVal x As Byte, ByVal y As Byte, ByVal CastRange As Byte) As Boolean
    If Not InMapBounds(Map, x, y) Then Exit Function
    If CastRange = 0 Then Exit Function
    If Abs(CInt(x) - CInt(CasterX)) > CastRange Or Abs(CInt(y) - CInt(CasterY)) > CastRange Then Exit Function
    If (MapData(x, y, Map).Blocked And e_Block.ALL_SIDES) = e_Block.ALL_SIDES Then Exit Function
    IsMagicLightOrbPositionValid = True
End Function

Public Function CastMagicLightOrb(ByVal UserIndex As Integer, ByVal SpellId As Integer) As Boolean
    On Error GoTo CastMagicLightOrb_Err
    If Not UserSupportsMagicLightOrbs(UserIndex) Then Exit Function
    If Not CanClassCastLightOrb(UserList(UserIndex).clase) Then Exit Function
    With Hechizos(SpellId)
        If .Duration <= 0 Or .Duration > MAX_ORB_DURATION_SECONDS Then Exit Function
        If .LightOrbObject <= 0 Or .LightOrbObject > UBound(ObjData) Then Exit Function
        If Len(ObjData(.LightOrbObject).CreaLuz) = 0 Then Exit Function
        With UserList(UserIndex)
            If .flags.TargetMap <> .pos.Map Or Not IsMagicLightOrbPositionValid(.pos.Map, .pos.x, .pos.y, .flags.TargetX, .flags.TargetY, Hechizos(SpellId).LightOrbCastRange) Then
                Call WriteLocaleMsg(UserIndex, MSG_LIGHT_ORB_POSITION, e_TextChannel.TEXTCHANNEL_COMBAT, e_FontTypeNames.FONTTYPE_New_Naranja, CStr(Hechizos(SpellId).LightOrbCastRange))
                Exit Function
            End If
        End With
    End With
    If NextInstanceId = MAX_ORB_INSTANCE_ID Then Exit Function
    If Not OrbsInitialized Then
        ReDim UserOrbs(1 To UBound(UserList))
        OrbsInitialized = True
    End If
    ' Validation precedes replacement, so a rejected cast preserves the old orb.
    Call RemoveMagicLightOrb(UserIndex)
    NextInstanceId = NextInstanceId + 1
    With UserOrbs(UserIndex)
        .InstanceId = NextInstanceId
        .Position.Map = UserList(UserIndex).pos.Map
        .Position.x = UserList(UserIndex).flags.TargetX
        .Position.y = UserList(UserIndex).flags.TargetY
        .ProfileObject = Hechizos(SpellId).LightOrbObject
        .CreatedAt = GetTickCountRaw()
        .DurationMs = CLng(Hechizos(SpellId).Duration) * 1000
    End With
    Call BroadcastMagicLightOrb(UserIndex, UserOrbs(UserIndex).DurationMs)
    CastMagicLightOrb = True
    Exit Function
CastMagicLightOrb_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.CastMagicLightOrb", Erl)
End Function

Private Sub BroadcastMagicLightOrb(ByVal OwnerIndex As Integer, ByVal RemainingMs As Long)
    On Error GoTo BroadcastMagicLightOrb_Err
    Dim entry As Integer
    Dim recipient As Integer
    Dim mapId As Integer
    mapId = UserOrbs(OwnerIndex).Position.Map
    For entry = 1 To ConnGroups(mapId).CountEntrys
        recipient = ConnGroups(mapId).UserEntrys(entry)
        If recipient > 0 Then Call SendMagicLightOrbToUser(recipient, OwnerIndex, RemainingMs)
    Next entry
    Exit Sub
BroadcastMagicLightOrb_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.BroadcastMagicLightOrb", Erl)
End Sub

Private Sub SendMagicLightOrbToUser(ByVal Recipient As Integer, ByVal OwnerIndex As Integer, ByVal RemainingMs As Long)
    On Error GoTo SendMagicLightOrbToUser_Err
    If Not UserSupportsMagicLightOrbs(Recipient) Then Exit Sub
    With UserOrbs(OwnerIndex)
        If UserList(Recipient).pos.Map <> .Position.Map Then Exit Sub
        Call WriteHooMagicLightOrb(Recipient, .InstanceId, .Position.Map, .Position.x, .Position.y, .ProfileObject, RemainingMs)
    End With
    Exit Sub
SendMagicLightOrbToUser_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.SendMagicLightOrbToUser", Erl)
End Sub

Public Sub RemoveMagicLightOrb(ByVal UserIndex As Integer)
    On Error GoTo RemoveMagicLightOrb_Err
    If Not OrbsInitialized Then Exit Sub
    If UserIndex < 1 Or UserIndex > UBound(UserOrbs) Then Exit Sub
    If UserOrbs(UserIndex).InstanceId = 0 Then Exit Sub
    Call BroadcastMagicLightOrb(UserIndex, 0)
    UserOrbs(UserIndex).InstanceId = 0
    Exit Sub
RemoveMagicLightOrb_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.RemoveMagicLightOrb", Erl)
End Sub

Public Sub UpdateMagicLightOrbs(ByVal NowRaw As Long)
    On Error GoTo UpdateMagicLightOrbs_Err
    If Not OrbsInitialized Then Exit Sub
    If TicksElapsed(LastUpdate, NowRaw) < ORB_UPDATE_INTERVAL_MS Then Exit Sub
    LastUpdate = NowRaw
    Dim owner As Integer
    For owner = 1 To UBound(UserOrbs)
        If UserOrbs(owner).InstanceId <> 0 Then
            If TicksElapsed(UserOrbs(owner).CreatedAt, NowRaw) >= UserOrbs(owner).DurationMs Or _
                Not UserList(owner).ConnectionDetails.ConnIDValida Or _
                UserList(owner).pos.Map <> UserOrbs(owner).Position.Map Then
                Call RemoveMagicLightOrb(owner)
            End If
        End If
    Next owner
    Exit Sub
UpdateMagicLightOrbs_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.UpdateMagicLightOrbs", Erl)
End Sub

Public Sub SendMagicLightOrbSnapshot(ByVal UserIndex As Integer)
    On Error GoTo SendMagicLightOrbSnapshot_Err
    If Not OrbsInitialized Then Exit Sub
    If Not UserSupportsMagicLightOrbs(UserIndex) Then Exit Sub
    Dim owner As Integer
    Dim remaining As Long
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    For owner = 1 To UBound(UserOrbs)
        With UserOrbs(owner)
            If .InstanceId <> 0 And .Position.Map = UserList(UserIndex).pos.Map Then
                remaining = IntervalRemainingMs(.CreatedAt, .DurationMs, nowRaw)
                If remaining > 0 Then Call SendMagicLightOrbToUser(UserIndex, owner, remaining)
            End If
        End With
    Next owner
    Exit Sub
SendMagicLightOrbSnapshot_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.SendMagicLightOrbSnapshot", Erl)
End Sub
