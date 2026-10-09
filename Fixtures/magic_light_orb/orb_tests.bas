Attribute VB_Name = "OrbTests"
' Copyright (C) 2026 Noland Studios LTD
' Licensed under the GNU Affero General Public License, version 3 or later.
' See licence.txt for the full license.
Option Explicit

Private Checks As Integer
Private Sub Check(ByVal Passed As Boolean, ByVal Name As String)
    Checks = Checks + 1
    If Not Passed Then Err.Raise vbObjectError + 1, "orb tests", Name
End Sub

Public Sub Main()
    On Error GoTo Main_Err
    Dim userIndex As Integer
    Dim oldId As Long
    Dim secondId As Long
    Dim output As Integer
    Hechizos(1).Duration = 60
    Hechizos(1).LightOrbObject = 1
    Hechizos(1).LightOrbCastRange = 6
    ObjData(1).CreaLuz = "1"
    TestNow = 1000
    ConnGroups(1).CountEntrys = 2
    ConnGroups(1).UserEntrys(1) = 1
    ConnGroups(1).UserEntrys(2) = 2
    For userIndex = 1 To 2
        With UserList(userIndex)
            .clase = Mage
            .pos.Map = 1
            .pos.x = 20
            .pos.y = 20
            .flags.TargetMap = 1
            .flags.TargetX = 21
            .flags.TargetY = 22
            .ConnectionDetails.ConnIDValida = True
        End With
        Supported(userIndex) = True
    Next userIndex
    Call Check(CanClassCastLightOrb(Mage) And CanClassCastLightOrb(Druid) And CanClassCastLightOrb(Bard), "allowed classes")
    Call Check(Not CanClassCastLightOrb(Cleric) And Not CanClassCastLightOrb(Warrior), "rejected classes")
    Call Check(CastMagicLightOrb(1, 1), "cast")
    Call Check(PacketCount = 2 And Packets(1).Remaining = 60000, "shared spawn")
    oldId = Packets(1).InstanceId
    Call Check(Packets(1).x = 21 And Packets(1).y = 22, "clicked tile")
    PacketCount = 0
    TestNow = 41000
    Call SendMagicLightOrbSnapshot(2)
    Call Check(PacketCount = 1 And Packets(1).Remaining = 20000, "late arrival remaining time")
    PacketCount = 0
    UserList(1).flags.TargetX = 27
    Call Check(Not CastMagicLightOrb(1, 1), "range rejection")
    Call Check(PacketCount = 0 And FeedbackCount = 1, "rejected cast preserves old orb")
    UserList(1).flags.TargetX = 21
    MapData(21, 22, 1).Blocked = ALL_SIDES
    Call Check(Not CastMagicLightOrb(1, 1), "wall rejection")
    MapData(21, 22, 1).Blocked = 0
    Call Check(CastMagicLightOrb(1, 1), "replacement")
    Call Check(PacketCount = 4 And Packets(1).InstanceId = oldId And Packets(1).Remaining = 0, "old orb removed")
    Call Check(Packets(3).InstanceId <> oldId And Packets(3).Remaining = 60000, "new identity")
    oldId = Packets(3).InstanceId
    PacketCount = 0
    Call Check(CastMagicLightOrb(2, 1), "same tile second caster")
    secondId = Packets(1).InstanceId
    Call Check(secondId <> oldId, "independent identities")
    PacketCount = 0
    Call RemoveMagicLightOrb(1)
    Call Check(PacketCount = 2 And Packets(1).InstanceId = oldId, "remove only owner")
    PacketCount = 0
    Call SendMagicLightOrbSnapshot(1)
    Call Check(PacketCount = 1 And Packets(1).InstanceId = secondId, "other caster retained")
    PacketCount = 0
    Supported(1) = False
    Call Check(Not CastMagicLightOrb(1, 1), "capability required")
    Call SendMagicLightOrbSnapshot(1)
    Call Check(PacketCount = 0, "legacy gets no packet")
    TestNow = 101000
    Call UpdateMagicLightOrbs(TestNow)
    Call Check(PacketCount = 1 And Packets(1).Remaining = 0, "expiry removal")
    PacketCount = 0
    Call SendMagicLightOrbSnapshot(2)
    Call Check(PacketCount = 0, "expired snapshot empty")
    TestNow = -100
    Call Check(CastMagicLightOrb(2, 1), "cast near tick wrap")
    PacketCount = 0
    TestNow = 100
    Call SendMagicLightOrbSnapshot(2)
    Call Check(PacketCount = 1 And Packets(1).Remaining = 59800, "wrap safe remaining")
    PacketCount = 0
    UserList(2).ConnectionDetails.ConnIDValida = False
    Call UpdateMagicLightOrbs(TestNow)
    Call Check(PacketCount = 1 And Packets(1).Remaining = 0, "disconnect cleanup")
    UserList(2).ConnectionDetails.ConnIDValida = True
    Call Check(CastMagicLightOrb(2, 1), "cast before changing map")
    PacketCount = 0
    UserList(2).pos.Map = 2
    TestNow = 300
    Call UpdateMagicLightOrbs(TestNow)
    Call SendMagicLightOrbSnapshot(2)
    Call Check(PacketCount = 0, "no old map orb on new map")
    Call Check(ErrorCount = 0, "no runtime errors")
    output = FreeFile
    Open App.Path & "\magic-light-orb-results.txt" For Output As #output
    Print #output, "PASS " & CStr(Checks) & " checks"
    Close #output
    Exit Sub
Main_Err:
    output = FreeFile
    Open App.Path & "\magic-light-orb-results.txt" For Output As #output
    Print #output, "FAIL " & Err.Description & " after " & CStr(Checks) & " checks"
    Close #output
End Sub
