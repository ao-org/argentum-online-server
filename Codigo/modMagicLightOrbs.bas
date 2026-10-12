Attribute VB_Name = "modMagicLightOrbs"
' Argentum 20 Game Server
' Copyright (C) 2026 Noland Studios LTD
' Licensed under the GNU Affero General Public License, version 3 or later.
' See licence.txt for the full license.
Option Explicit

Private Const ORB_UPDATE_INTERVAL_MS As Long = 100
Private Const MAX_ORB_DURATION_SECONDS As Long = 3600

Private Type t_MagicLightOrb
    Active As Boolean
    Position As t_WorldPos
    Color As Long
    Range As Byte
    Particle As Integer
    CreatedAt As Long
    DurationMs As Long
End Type

Private UserOrbs() As t_MagicLightOrb
Private OrbsInitialized As Boolean
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
    If Not CanClassCastLightOrb(UserList(UserIndex).clase) Then Exit Function
    With Hechizos(SpellId)
        If .Duration <= 0 Or .Duration > MAX_ORB_DURATION_SECONDS Then Exit Function
        If .LightOrbColor = 0 Or .LightOrbRange = 0 Or .Particle <= 0 Then Exit Function
        With UserList(UserIndex)
            If .flags.TargetMap <> .pos.Map Or Not IsMagicLightOrbPositionValid(.pos.Map, .pos.x, .pos.y, .flags.TargetX, .flags.TargetY, Hechizos(SpellId).LightOrbCastRange) Then
                Call WriteLocaleMsg(UserIndex, MSG_LIGHT_ORB_POSITION, e_TextChannel.TEXTCHANNEL_COMBAT, e_FontTypeNames.FONTTYPE_New_Naranja, CStr(Hechizos(SpellId).LightOrbCastRange))
                Exit Function
            End If
        End With
    End With
    If Not OrbsInitialized Then
        ReDim UserOrbs(1 To UBound(UserList))
        OrbsInitialized = True
    End If
    ' Validation precedes replacement, so a rejected cast preserves the old orb.
    Call RemoveMagicLightOrb(UserIndex)
    With UserOrbs(UserIndex)
        .Active = True
        .Position.Map = UserList(UserIndex).pos.Map
        .Position.x = UserList(UserIndex).flags.TargetX
        .Position.y = UserList(UserIndex).flags.TargetY
        .Color = Hechizos(SpellId).LightOrbColor
        .Range = Hechizos(SpellId).LightOrbRange
        .Particle = Hechizos(SpellId).Particle
        .CreatedAt = GetTickCountRaw()
        .DurationMs = CLng(Hechizos(SpellId).Duration) * 1000
    End With
    Call BroadcastMagicLightOrbTile(UserOrbs(UserIndex).Position, UserOrbs(UserIndex).Particle)
    CastMagicLightOrb = True
    Exit Function
CastMagicLightOrb_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.CastMagicLightOrb", Erl)
End Function

Private Function OrbOwnerAt(ByRef Position As t_WorldPos, ByVal NowRaw As Long) As Integer
    Dim owner As Integer
    If Not OrbsInitialized Then Exit Function
    For owner = 1 To UBound(UserOrbs)
        With UserOrbs(owner)
            If .Active And .Position.Map = Position.Map And .Position.x = Position.x And .Position.y = Position.y Then
                If TicksElapsed(.CreatedAt, NowRaw) < .DurationMs Then
                    OrbOwnerAt = owner
                    Exit Function
                End If
            End If
        End With
    Next owner
End Function

Private Sub BroadcastMagicLightOrbTile(ByRef Position As t_WorldPos, ByVal RemovedParticle As Integer)
    On Error GoTo BroadcastMagicLightOrbTile_Err
    Dim entry As Integer
    Dim recipient As Integer
    Dim owner As Integer
    owner = OrbOwnerAt(Position, GetTickCountRaw())
    For entry = 1 To ConnGroups(Position.Map).CountEntrys
        recipient = ConnGroups(Position.Map).UserEntrys(entry)
        If recipient > 0 Then Call SendMagicLightOrbTile(recipient, Position, owner, RemovedParticle)
    Next entry
    Exit Sub
BroadcastMagicLightOrbTile_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.BroadcastMagicLightOrbTile", Erl)
End Sub

#If UNIT_TEST = 1 Then
' Runs in the normal server unit suite, not a standalone fixture project.
Public Function test_suite_light_orbs() As Boolean
    On Error GoTo test_suite_light_orbs_Err
    If OrbsInitialized Then
        Call UnitTesting.RunTestError("floor orb ownership", "requires empty unit-test orb state")
        Exit Function
    End If
    ReDim UserOrbs(1 To 3)
    OrbsInitialized = True
    Dim tile As t_WorldPos
    tile.Map = 1: tile.x = 20: tile.y = 30
    Dim owner As Integer
    For owner = 1 To 2
        UserOrbs(owner).Active = True
        UserOrbs(owner).Position = tile
        UserOrbs(owner).CreatedAt = 100
        UserOrbs(owner).DurationMs = 60000
    Next owner
    Call UnitTesting.RunTest("floor orb overlap has one tile presentation", OrbOwnerAt(tile, 200) = 1)
    UserOrbs(1).Active = False
    Call UnitTesting.RunTest("removing first orb preserves second owner", OrbOwnerAt(tile, 200) = 2)
    UserOrbs(2).Active = False
    Call UnitTesting.RunTest("last orb removal clears tile ownership", OrbOwnerAt(tile, 200) = 0)
    UserOrbs(1).Active = True
    Call UnitTesting.RunTest("expired orb not included in entry snapshot", OrbOwnerAt(tile, 60100) = 0)
    UserOrbs(1).CreatedAt = &HFFFFFFF0
    UserOrbs(1).DurationMs = 100
    Call UnitTesting.RunTest("orb remains valid across clock wrap", OrbOwnerAt(tile, &H10) = 1)
    Call UnitTesting.RunTest("orb expires across clock wrap", OrbOwnerAt(tile, &H54) = 0)
    tile.Map = 2
    Call UnitTesting.RunTest("orb ownership does not cross maps", OrbOwnerAt(tile, &H10) = 0)
    Call UnitTesting.RunTest("only mage druid bard can cast", CanClassCastLightOrb(Mage) And CanClassCastLightOrb(Druid) And CanClassCastLightOrb(Bard) And Not CanClassCastLightOrb(Warrior))
    Erase UserOrbs
    OrbsInitialized = False
    test_suite_light_orbs = True
    Exit Function
test_suite_light_orbs_Err:
    Erase UserOrbs
    OrbsInitialized = False
    Call UnitTesting.RunTestError("floor orb ownership suite", Err.Description)
End Function
#End If

Private Sub SendMagicLightOrbTile(ByVal Recipient As Integer, ByRef Position As t_WorldPos, ByVal OwnerIndex As Integer, ByVal RemovedParticle As Integer)
    On Error GoTo SendMagicLightOrbTile_Err
    If UserList(Recipient).pos.Map <> Position.Map Then Exit Sub
    If OwnerIndex > 0 Then
        With UserOrbs(OwnerIndex)
            Call SendData(SendTarget.ToIndex, Recipient, PrepareMessageLightFXToFloor(Position.x, Position.y, .Color, .Range))
            ' -1 is the existing persistent floor-particle lifetime. Only the
            ' server expires the spell; area snapshots never restart its timer.
            Call SendData(SendTarget.ToIndex, Recipient, PrepareMessageParticleFXToFloor(Position.x, Position.y, .Particle, -1))
        End With
    Else
        ' Do not mutate MapData or create a ground object. Restore any current
        ' server floor light when the final overlapping orb disappears.
        With MapData(Position.x, Position.y, Position.Map).Luz
            Call SendData(SendTarget.ToIndex, Recipient, PrepareMessageLightFXToFloor(Position.x, Position.y, .Color, .Rango))
        End With
        Call SendData(SendTarget.ToIndex, Recipient, PrepareMessageParticleFXToFloor(Position.x, Position.y, RemovedParticle, 0))
    End If
    Exit Sub
SendMagicLightOrbTile_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.SendMagicLightOrbTile", Erl)
End Sub

Public Sub RemoveMagicLightOrb(ByVal UserIndex As Integer)
    On Error GoTo RemoveMagicLightOrb_Err
    If Not OrbsInitialized Then Exit Sub
    If UserIndex < 1 Or UserIndex > UBound(UserOrbs) Then Exit Sub
    If Not UserOrbs(UserIndex).Active Then Exit Sub
    ' Remove ownership first. Another caster on this tile keeps it lit.
    UserOrbs(UserIndex).Active = False
    Call BroadcastMagicLightOrbTile(UserOrbs(UserIndex).Position, UserOrbs(UserIndex).Particle)
    Exit Sub
RemoveMagicLightOrb_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.RemoveMagicLightOrb", Erl)
End Sub

Public Sub UpdateMagicLightOrbs()
    On Error GoTo UpdateMagicLightOrbs_Err
    If Not OrbsInitialized Then Exit Sub
    ' Packet handling may create an orb after GlobalFrameTime was sampled.
    ' Read once here so wrap-safe elapsed time never sees an older frame tick.
    Dim NowRaw As Long
    NowRaw = GetTickCountRaw()
    If TicksElapsed(LastUpdate, NowRaw) < ORB_UPDATE_INTERVAL_MS Then Exit Sub
    LastUpdate = NowRaw
    Dim owner As Integer
    For owner = 1 To UBound(UserOrbs)
        If UserOrbs(owner).Active Then
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
    Dim owner As Integer
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    For owner = 1 To UBound(UserOrbs)
        With UserOrbs(owner)
            If .Active And .Position.Map = UserList(UserIndex).pos.Map Then
                ' Exactly one presentation per tile, even with several owners.
                If OrbOwnerAt(.Position, nowRaw) = owner Then Call SendMagicLightOrbTile(UserIndex, .Position, owner, 0)
            End If
        End With
    Next owner
    Exit Sub
SendMagicLightOrbSnapshot_Err:
    Call TraceError(Err.Number, Err.Description, "modMagicLightOrbs.SendMagicLightOrbSnapshot", Erl)
End Sub
