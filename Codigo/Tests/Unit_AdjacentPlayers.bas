Attribute VB_Name = "Unit_AdjacentPlayers"
Option Explicit
#If UNIT_TEST = 1 Then
Private capturing As Boolean
Private packets As Collection
Private accepted As Collection
Private failAt As Long
Private fixtureClan As clsClan
Private readerBytes() As Byte
Private Const TOPOLOGY As String = "{""version"":1,""id_space"":""resource_map_id"",""maps"":[34,35,78,80],""edges"":[[34,35,0,-80],[35,34,0,80],[34,78,74,0],[78,34,-74,0],[35,80,74,0],[80,35,-74,0],[78,80,0,-80],[80,78,0,80]]}"

Public Function CapturePacket(ByVal route As SendTarget, ByVal recipient As Integer, ByRef queued As Boolean) As Boolean
    If Not capturing Then Exit Function
    Dim writer As Network.writer, bytes() As Byte, value As Variant
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(bytes)
    Call writer.Clear()
    queued = True
    If route = ToIndex And recipient = 1 Then
        value = bytes
        Call packets.Add(value)
        queued = packets.Count <> failAt
        Call accepted.Add(queued)
    End If
    CapturePacket = True
End Function

Private Sub Expect(ByVal condition As Boolean, ByVal description As String)
    If Not condition Then Err.Raise vbObjectError + 191, "Adjacent player regression", description
End Sub

Private Sub ClearCapture(Optional ByVal failure As Long = 0)
    Set packets = New Collection
    Set accepted = New Collection
    failAt = failure
End Sub

Private Sub Sample(Optional ByVal failure As Long = 0)
    Call ClearCapture(failure)
    Call AdjacentPlayerTestRefresh(1)
    Call Expect(Protocol_Writes.GetWriterBuffer().GetOffset() = 0, "shared writer cleared")
End Sub

Private Function PacketReader(ByVal index As Long, ByVal id As Integer) As Network.reader
    Dim reader As New Network.reader
    readerBytes = packets(index)
    Call reader.SetData(readerBytes)
    Call Expect(reader.ReadInt16() = id, "packet order/id " & CStr(index))
    Set PacketReader = reader
End Function

Private Function PacketIdAt(ByVal index As Long) As Integer
    Dim bytes() As Byte
    bytes = packets(index)
    PacketIdAt = CInt(CLng(bytes(0)) + 256& * bytes(1))
End Function

Private Function FlagsAt(ByVal index As Long) As Long
    Dim bytes() As Byte
    bytes = packets(index)
    FlagsAt = CLng(bytes(40)) + 256& * bytes(41)
End Function

Private Sub PutUser(ByVal slot As Integer, ByVal map As Integer, ByVal x As Byte, ByVal y As Byte)
    Dim emptyUser As t_User
    UserList(slot) = emptyUser
    With UserList(slot)
        .flags.UserLogged = True: .flags.Privilegios = e_PlayerType.User
        .ConnectionDetails.ConnIDValida = True: .ConnectionDetails.ConnID = slot + 40
        .VersionId = 7: .pos.Map = map: .pos.x = x: .pos.y = y
        .Char.charindex = slot: .Char.body = 10: .Char.head = 20: .Char.Heading = e_Heading.SOUTH
        .Char.WeaponAnim = 30: .Char.ShieldAnim = 40: .Char.CascoAnim = 50
        .Char.CartAnim = 60: .Char.BackpackAnim = 70: .Char.speeding = 1.5
        .Faccion.Status = 1: .Stats.tipoUsuario = 2: .flags.CurrentTeam = 3: .flags.tiene_bandera = 4
        .name = "Jos" & ChrW$(&HE9): .showName = True
        .Char.Arma_Aura = "weapon": .Char.Body_Aura = "body": .Char.Escudo_Aura = "shield"
        .Char.Head_Aura = "head": .Char.Otra_Aura = "other": .Char.DM_Aura = "dm": .Char.RM_Aura = "rm"
        .AreasInfo.AreaReciveX = 1: .AreasInfo.AreaReciveY = 1
        .AreasInfo.AreaPerteneceX = 1: .AreasInfo.AreaPerteneceY = 1
    End With
    CharList(slot) = slot
    MapData(x, y, map).UserIndex = slot
    With ConnGroups(map)
        .CountEntrys = .CountEntrys + 1
        .UserEntrys(.CountEntrys) = slot
    End With
End Sub

Private Sub Fixture()
    Dim slot As Long, map As Variant, x As Long, y As Long, emptyUser As t_User
    Call AdjacentPlayerTestReset()
    Call Expect(LoadAdjacentTopologyJson(TOPOLOGY, False), "topology fixture")
    For Each map In Array(34, 35, 78, 80)
        MapInfo(map).MapResource = CInt(map): MapInfo(map).Seguro = 0
        ConnGroups(map).CountEntrys = 0
        ConnGroups(map).OptValue = 131
        ReDim ConnGroups(map).UserEntrys(1 To 131)
        For x = 1 To 100
            For y = 1 To 100
                MapData(x, y, map).UserIndex = 0
            Next y
        Next x
    Next map
    For slot = 1 To 131
        UserList(slot) = emptyUser
        CharList(slot) = 0
    Next slot
    LastUser = 131
    RequiredGuildLevelSeeInvisible = 5
    Call fixtureClan.SetTestLevel(5)
    Call PutUser(1, 34, 50, 11)
    Call PutUser(2, 35, 50, 90)
    Call ClearCapture()
End Sub

Private Sub WriteHex(ByVal index As Long)
    Dim bytes() As Byte, i As Long, hexText As String, f As Integer
    bytes = packets(index)
    For i = LBound(bytes) To UBound(bytes)
        If i > LBound(bytes) Then hexText = hexText & " "
        hexText = hexText & Right$("0" & Hex$(bytes(i)), 2)
    Next i
    f = FreeFile()
    Open App.Path & "\adjacent-player-upsert.hex" For Output As #f
    Print #f, hexText
    Close #f
End Sub

Private Sub CheckCase(ByVal id As Long)
    Dim value As Variant, oldValue As Variant, reader As Network.reader, i As Long, epoch As Long, bytes() As Byte
    Dim label As String, invisible As Boolean, queued As Boolean
    Select Case id
        Case 1
            Call AdjacentPlayerTestCounters(&H11223343, &H55667788, 0)
            Call Sample()
            Call Expect(packets.Count = 2 And AdjacentPlayerTestCount(1) = 1, "direct reset/upsert")
            #If PYMMO = 1 Then
                Call Expect(eAdjacentNpcReset = 214 And eAdjacentNpcRemove = 216 And eAdjacentPlayerReset = 217 And eAdjacentPlayerRemove = 219, "P1 IDs")
            #Else
                Call Expect(eAdjacentNpcReset = 215 And eAdjacentNpcRemove = 217 And eAdjacentPlayerReset = 218 And eAdjacentPlayerRemove = 220, "P0 IDs")
            #End If
            Set reader = PacketReader(1, eAdjacentPlayerReset)
            Call Expect(reader.ReadInt32() = 1 And reader.ReadInt16() = 34, "reset layout")
            Set reader = PacketReader(2, eAdjacentPlayerUpsert)
            Call Expect(reader.ReadInt32() = 1 And reader.ReadInt16() = 2, "epoch/slot")
            Call Expect(reader.ReadInt32() = &H11223345 And reader.ReadInt32() = &H55667788, "low/high uint64")
            Call Expect(reader.ReadInt16() = 35 And reader.ReadInt8() = 50 And reader.ReadInt8() = 90, "source placement")
            Call Expect(reader.ReadInt16() = 10 And reader.ReadInt16() = 20 And reader.ReadInt8() = e_Heading.SOUTH, "appearance")
            Call Expect(reader.ReadInt16() = 30 And reader.ReadInt16() = 40 And reader.ReadInt16() = 50 And reader.ReadInt16() = 60 And reader.ReadInt16() = 70, "equipment order")
            Call Expect(reader.ReadReal32() = 1.5 And reader.ReadInt8() = 0 And reader.ReadInt16() = 96, "speed/mode/flags")
            Call Expect(reader.ReadInt8() = 1 And reader.ReadInt8() = 1 And reader.ReadInt8() = 2 And reader.ReadInt8() = 3 And reader.ReadInt8() = 4, "metadata")
            Call Expect(reader.ReadString8() = "Jos" & ChrW$(&HE9) & " {}", "UTF8 authorized label")
            For Each value In Array("weapon", "body", "shield", "head", "other", "dm", "rm")
                Call Expect(reader.ReadString8() = CStr(value), "aura order")
            Next value
            Call WriteHex(2)
            Call Sample()
            Call Expect(packets.Count = 0, "unchanged is silent despite initial snap")
        Case 2
            UserList(1).pos.x = 87
            UserList(2).pos.Map = 80: UserList(2).pos.x = 14
            MapData(14, 90, 80).UserIndex = 2
            Call Sample()
            Call Expect(AdjacentPlayerTestCount(1) = 1, "two hop diagonal")
        Case 3
            UserList(2).pos.x = 74: UserList(2).pos.y = 73
            MapData(74, 73, 35).UserIndex = 2
            Call Sample()
            Call Expect(AdjacentPlayerTestCount(1) = 1, "inclusive 24/18 boundary")
            UserList(2).pos.x = 75
            Call Sample()
            Call Expect(AdjacentPlayerTestCount(1) = 0 And PacketIdAt(1) = eAdjacentPlayerRemove, "outside boundary remove")
        Case 4
            For i = 3 To 131
                Call PutUser(CInt(i), 35, 50, 90)
            Next i
            ' Reverse membership order to ensure cap selection does not depend on insertion.
            For i = 1 To 130
                ConnGroups(35).UserEntrys(i) = 132 - i
            Next i
            MapData(50, 90, 35).NpcIndex = 1
            Call Sample()
            Call Expect(packets.Count = 129 And AdjacentPlayerTestCount(1) = 128, "independent cap with overlapping occupants")
            Set reader = PacketReader(129, eAdjacentPlayerUpsert)
            Call reader.ReadInt32()
            Call Expect(reader.ReadInt16() = 129, "stable ascending slot cap")
        Case 5, 6
            If id = 5 Then UserList(2).flags.invisible = 1 Else UserList(2).flags.Oculto = 1
            Call Sample()
            Call Expect(packets.Count = 1 And AdjacentPlayerTestCount(1) = 0, "unauthorized absent before label/location snapshot")
            UserList(2).flags.invisible = 0: UserList(2).flags.Oculto = 0
            Call Sample()
            Call Expect(packets.Count = 1 And PacketIdAt(1) = eAdjacentPlayerUpsert, "stationary reveal")
        Case 7, 8, 9
            UserList(2).flags.AdminInvisible = 1: UserList(2).flags.Privilegios = e_PlayerType.SemiDios
            UserList(2).flags.Navegando = 1
            UserList(1).flags.Privilegios = e_PlayerType.Consejero
            If id = 8 Then UserList(1).flags.Privilegios = e_PlayerType.SemiDios
            If id = 9 Then UserList(1).flags.Privilegios = e_PlayerType.Dios
            UserList(1).flags.Muerto = 1
            Call Sample()
            If id = 7 Then
                Call Expect(AdjacentPlayerTestCount(1) = 0, "lower admin denied before boat")
            Else
                value = AdjacentPlayerTestValue(1, 2)
                Call Expect(value(14) = 1 And (value(15) And 32) <> 0, "authorized admin bypasses dead restriction; staff boat auras")
            End If
        Case 10, 11
            UserList(2).flags.invisible = 1
            UserList(1).flags.Privilegios = e_PlayerType.RoleMaster
            If id = 11 Then UserList(2).flags.Navegando = 1
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(value(14) = 1, "staff representation")
            Call Expect(((value(15) And 32) <> 0) = (id = 11), "independent staff aura flag")
        Case 12, 13, 14, 15, 16
            UserList(1).GuildIndex = 1: UserList(2).GuildIndex = 1
            UserList(2).flags.invisible = 1
            If id = 13 Then Call fixtureClan.SetTestLevel(4)
            If id = 14 Then UserList(2).Char.body = 830
            If id = 15 Then UserList(2).flags.Mimetizado = 1
            If id = 16 Then UserList(2).showName = False
            Call Sample()
            If id = 12 Then
                value = AdjacentPlayerTestValue(1, 2)
                Call Expect(value(14) = 2 And (value(15) And 160) = 160, "clan mode/auras/label tint")
            Else
                Call Expect(AdjacentPlayerTestCount(1) = 0, "clan rendering excluded")
            End If
        Case 17
            UserList(2).flags.invisible = 1: UserList(2).flags.Navegando = 1
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(value(14) = 3 And (value(15) And 34) = 34, "boat mode and auras")
        Case 18, 19, 20, 21
            UserList(1).flags.Muerto = 1
            If id = 19 Then MapInfo(34).Seguro = 1
            If id = 20 Then UserList(2).flags.StatusMask = e_StatusMask.eTalkToDead
            If id = 21 Then UserList(1).flags.StatusMask = e_StatusMask.eTalkToDead
            Call Sample()
            Call Expect((AdjacentPlayerTestCount(1) = 1) = (id = 19 Or id = 20), "dead viewer source-vs-viewer permissions")
        Case 22
            Call Sample()
            UserList(2).flags.Mimetizado = 1: UserList(2).NameMimetizado = "Disguise"
            UserList(1).GuildIndex = 1: UserList(2).GuildIndex = 1
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(packets.Count = 1 And value(21) = "Disguise" And (value(15) And 128) = 0, "stationary mimic exact name no raw clan")
            UserList(2).NameMimetizado = ""
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(value(21) = "", "empty mimic no fallback")
            UserList(2).showName = False
            UserList(2).NameMimetizado = "Secret"
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(value(21) = "", "showName wins")
        Case 23
            UserList(2).flags.EnConsulta = True: UserList(2).Alias = "UnauthorizedAlias"
            UserList(2).Stats.tipoUsuario = 0
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(InStr(value(21), "[CONSULTA]") > 0 And InStr(value(21), "UnauthorizedAlias") = 0, "consultation and gated alias")
        Case 24
            UserList(2).clase = e_Class.Pirat: UserList(2).flags.Oculto = 1
            UserList(1).flags.Privilegios = e_PlayerType.Admin
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(value(21) = "", "hidden pirate label")
        Case 25
            Call Sample()
            UserList(2).Char.BackpackAnim = 99: UserList(2).Char.RM_Aura = "new aura"
            UserList(2).Faccion.Status = 2: UserList(2).flags.CurrentTeam = 4
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(packets.Count = 1 And value(12) = 99 And value(28) = "new aura" And value(16) = 2 And value(19) = 4, "stationary full metadata/equipment/aura refresh")
        Case 26
            Call Sample()
            UserList(2).pos.x = 51
            Call Sample()
            Call Expect((FlagsAt(1) And 64) = 0, "one tile walk interpolates")
            #If PYMMO = 1 Then
                Call RelocateAdjacentPlayer(2)
                UserList(2).pos.x = 52
                Call Sample(1)
                Call Expect((FlagsAt(1) And 64) = 64 And Not accepted(1), "one tile warp snap failed enqueue")
                Call Sample()
                Call Expect((FlagsAt(1) And 64) = 64, "warp snap retained on retry")
                Call Sample()
                Call Expect(packets.Count = 0, "no redundant snap clear")
            #End If
        Case 27
            Call Sample()
            UserList(2).flags.UserLogged = False
            Call Sample()
            Call Expect(packets.Count = 1 And PacketIdAt(1) = eAdjacentPlayerRemove And AdjacentPlayerTestCount(1) = 0, "logout remove")
        Case 28
            Call Sample()
            oldValue = AdjacentPlayerTestValue(1, 2)
            Call ResetAdjacentPlayerSlot(2)
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(packets.Count = 2 And PacketIdAt(1) = eAdjacentPlayerRemove And PacketIdAt(2) = eAdjacentPlayerUpsert And value(0) <> oldValue(0), "same slot/version/connection new lifetime")
        Case 29
            Call Sample()
            Call ResetAdjacentPlayerSlot(1)
            Call Sample()
            Set reader = PacketReader(1, eAdjacentPlayerReset)
            Call Expect(reader.ReadInt32() = 2, "viewer replacement new epoch")
        Case 30
            Call Sample()
            UserList(1).pos.Map = 35
            Call Sample()
            Set reader = PacketReader(1, eAdjacentPlayerReset)
            Call Expect(reader.ReadInt32() = 2 And reader.ReadInt16() = 35 And AdjacentPlayerTestCount(1) = 0, "new map local actor excluded")
        Case 31
            Call Sample(1)
            Call Expect(packets.Count = 1 And AdjacentPlayerTestCount(1) = 0, "failed reset no initialization/upserts")
            Call Sample()
            Call Expect(packets.Count = 2 And AdjacentPlayerTestCount(1) = 1, "failed reset retries")
        Case 32
            Call Sample(2)
            Call Expect(AdjacentPlayerTestCount(1) = 0, "failed initial upsert not cached")
            Call Sample()
            Call Expect(packets.Count = 1 And PacketIdAt(1) = eAdjacentPlayerUpsert, "unchanged initial failure retries")
        Case 33
            Call Sample()
            UserList(2).Char.head = 21
            Call Sample(1)
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(value(6) = 20, "failed refresh retains accepted appearance")
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(packets.Count = 1 And value(6) = 21, "unchanged refresh failure retries")
        Case 34
            Call Sample()
            UserList(2).flags.invisible = 1
            Call Sample(1)
            Call Expect(AdjacentPlayerTestCount(1) = 1, "failed removal retained")
            Call Sample()
            Call Expect(packets.Count = 1 And AdjacentPlayerTestCount(1) = 0, "failed removal retry")
        Case 35, 36
            Call PutUser(3, 35, 51, 90)
            Call Sample(3)
            Call Expect(AdjacentPlayerTestCount(1) = 1, "accepted prefix committed")
            Call Sample()
            Call Expect(packets.Count = 1 And AdjacentPlayerTestCount(1) = 2, "partial upsert recovery")
            If id = 36 Then
                UserList(2).flags.UserLogged = False: UserList(3).flags.UserLogged = False
                Call Sample(2)
                Call Expect(AdjacentPlayerTestCount(1) = 1, "partial remove commit")
                Call Sample()
                Call Expect(packets.Count = 1 And AdjacentPlayerTestCount(1) = 0, "partial remove recovery")
            End If
        Case 37
            Call Sample()
            Call ResetAdjacentPlayerSlot(2)
            Call Sample(1)
            Call Expect(packets.Count = 1 And PacketIdAt(1) = eAdjacentPlayerRemove, "identity remove failure blocks replacement")
            Call Sample()
            Call Expect(packets.Count = 2 And PacketIdAt(2) = eAdjacentPlayerUpsert, "replacement only after accepted remove")
        Case 38
            UserList(2).flags.Mimetizado = 1: UserList(2).NameMimetizado = String$(4096, "a")
            Call Sample()
            Call Expect(AdjacentPlayerTestCount(1) = 1, "4096 byte label allowed")
            UserList(2).NameMimetizado = String$(2049, ChrW$(&HE9))
            Call Sample()
            Call Expect(packets.Count = 0, "UTF8 bytes not UTF16 character limit")
            UserList(2).NameMimetizado = String$(128, "b")
            Call Sample()
            bytes = packets(1)
            Call Expect(bytes(47) = 128 And bytes(48) = 1, "7 bit varint UTF8 prefix")
        Case 39
            UserList(2).Char.RM_Aura = String$(1024, "a")
            Call Sample()
            Call Expect(AdjacentPlayerTestCount(1) = 1, "1024 byte aura allowed")
            UserList(2).Char.RM_Aura = String$(513, ChrW$(&HE9))
            Call Sample()
            Call Expect(packets.Count = 0, "oversized aura rejects and clears writer")
            UserList(2).Char.RM_Aura = "fixed"
            Call Sample()
            Call Expect(packets.Count = 1, "corrected aura retries")
        Case 40
            Call AdjacentPlayerTestCounters(&H7FFFFFFE, 8, 0)
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(value(0) = &H80000000 And value(1) = 8, "signed low word boundary")
        Case 41
            Call AdjacentPlayerTestCounters(-2, 8, 0)
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(value(0) = 0 And value(1) = 9, "unsigned low word carry")
        Case 42
            Call AdjacentPlayerTestCounters(-2, -1, 0)
            Call Sample()
            Call Expect(packets.Count = 1 And AdjacentPlayerTestCount(1) = 0, "max uint64 fail closed")
            Call Sample()
            Call Expect(packets.Count = 0, "exhaustion never reuses ID")
        Case 43
            Call AdjacentPlayerTestCounters(0, 0, &H7FFFFFFE)
            Call Sample()
            Set reader = PacketReader(1, eAdjacentPlayerReset)
            Call Expect(reader.ReadInt32() = &H7FFFFFFF, "last positive epoch")
            Call InvalidateAdjacentPlayerViewer(1)
            Call Sample()
            Call Expect(packets.Count = 0 And AdjacentPlayerTestCount(1) = 0, "epoch cannot wrap")
        Case 44, 45, 46, 47
            #If PYMMO = 1 Then
                Call Sample()
                UserList(2).pos.Map = 34: UserList(2).pos.y = 12
                Call ClearCapture(1)
                Call MakeUserChar(False, 1, 2, 34, 50, 12)
                Call Expect(packets.Count = 1 And PacketIdAt(1) = eAdjacentPlayerRemove, "failed adjacent remove defers only create")
                If id = 45 Then Call ResetAdjacentPlayerSlot(2)
                If id = 46 Then UserList(2).flags.invisible = 1
                If id = 47 Then UserList(2).AreasInfo.AreaPerteneceX = 2
                Call Sample()
                If id = 44 Then
                    Call Expect(packets.Count >= 2 And PacketIdAt(1) = eAdjacentPlayerRemove And PacketIdAt(2) = eCharacterCreate, "retry remove before active create")
                Else
                    For i = 1 To packets.Count
                        Call Expect(PacketIdAt(i) <> eCharacterCreate, "stale session/hidden/out-of-area create discarded")
                    Next i
                End If
            #End If
        Case 48
            #If PYMMO = 1 Then
                UserList(2).pos.Map = 34
                Call ClearCapture()
                Call MakeUserChar(False, 1, 2, 34, 50, 90)
                Call Expect(PacketIdAt(1) = eCharacterCreate, "normal create tracked")
                UserList(2).pos.Map = 35
                Call Sample(2)
                Call Expect(packets.Count = 2 And PacketIdAt(2) = eCharacterRemove And AdjacentPlayerTestCount(1) = 0, "failed normal removal blocks adjacent upsert")
                Call Sample()
                Call Expect(packets.Count = 2 And PacketIdAt(1) = eCharacterRemove And PacketIdAt(2) = eAdjacentPlayerUpsert, "reverse handoff accepted barrier")
            #End If
        Case 49
            Call ClearCapture()
            Call UpdateAdjacentPlayerVisibility(1000)
            #If PYMMO = 1 Then
                Call Expect(packets.Count >= 2, "production tick")
                Call ClearCapture()
                UserList(2).Char.head = 21
                Call UpdateAdjacentPlayerVisibility(1099)
                Call Expect(packets.Count = 0, "100 ms throttle")
                Call UpdateAdjacentPlayerVisibility(1100)
                Call Expect(packets.Count = 1, "100 ms refresh")
            #Else
                Call Expect(packets.Count = 0, "P0 production stream disabled")
            #End If
        Case 50
            CharList(2) = 0
            Call Sample()
            Call Expect(AdjacentPlayerTestCount(1) = 0, "unadmitted target rejected")
            CharList(1) = 0
            Call InvalidateAdjacentPlayerViewer(1)
            Call Sample()
            Call Expect(packets.Count = 0, "unready viewer no reset")
        Case 51
            MapInfo(35).MapResource = 34
            Call Sample()
            Call Expect(AdjacentPlayerTestCount(1) = 0, "instance source rejected")
        Case 52
            UserList(1).flags.Muerto = 1
            UserList(1).GuildIndex = 1: UserList(2).GuildIndex = 1
            Call fixtureClan.SetTestLevel(4)
            Call Sample()
            Call Expect(AdjacentPlayerTestCount(1) = 0, "dead clan level required")
            Call fixtureClan.SetTestLevel(5)
            Call Sample()
            Call Expect(AdjacentPlayerTestCount(1) = 1, "dead clan level admission stationary")
        Case 53
            UserList(2).flags.Montado = 1: UserList(2).flags.Nadando = 1: UserList(2).Char.body = iTrajeAltoNw
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect((value(15) And 28) = 0, "gameplay transport flags are reserved; displayed body is authoritative")
            Call Expect(value(5) = iTrajeAltoNw, "authorized rubber appearance retained")
        Case 54, 55
            #If PYMMO = 1 Then
                Call Sample()
                oldValue = AdjacentPlayerTestValue(1, 2)
                i = 2
                If id = 55 Then i = 1
                ' These fixture slots were not taken from the production free-slot heap.
                UserList(i).flags.IsSlotFree = True
                Call ResetUserSlot(CInt(i))
                Call Expect(Not UserList(i).flags.UserLogged, "real ResetUserSlot completed")
                If id = 54 Then Call PutUser(2, 35, 50, 90) Else Call PutUser(1, 34, 50, 11)
                Call Sample()
                If id = 54 Then
                    value = AdjacentPlayerTestValue(1, 2)
                    Call Expect(value(0) <> oldValue(0) And PacketIdAt(1) = eAdjacentPlayerRemove, "real source reset does not reuse session")
                Else
                    Set reader = PacketReader(1, eAdjacentPlayerReset)
                    Call Expect(reader.ReadInt32() = 2, "real viewer reset starts new scene")
                End If
            #End If
        Case 56
            #If PYMMO = 1 Then
                Call Sample()
                UserList(2).pos.Map = 34: UserList(2).pos.y = 12
                Call ClearCapture(2)
                Call MakeUserChar(False, 1, 2, 34, 50, 12)
                Call Expect(packets.Count = 2 And accepted(1) And Not accepted(2), "remove accepted, create failed")
                Call Sample()
                Call Expect(packets.Count >= 1 And PacketIdAt(1) = eCharacterCreate, "failed active create retried without replaying remove")
                Call Sample()
                Call Expect(packets.Count = 0, "successful deferred create retired")
            #End If
        Case 57
            Call Sample()
            UserList(2).Char.speeding = CSng(1.00000011920929)
            Call Sample()
            Call Expect(packets.Count = 1, "exact float speed change")
        Case 58
            Call Sample()
            UserList(2).flags.StatusMask = e_StatusMask.eTalkToDead
            UserList(2).GuildIndex = 1: UserList(1).GuildIndex = 1
            UserList(1).flags.Muerto = 1
            UserList(2).flags.Mimetizado = 1: UserList(2).NameMimetizado = "ally disguise"
            Call Sample()
            value = AdjacentPlayerTestValue(1, 2)
            Call Expect(value(21) = "ally disguise" And (value(15) And 128) = 0, "dead admission does not leak disguised clan identity")
        Case 59, 60
            #If PYMMO = 1 Then
                UserList(2).pos.Map = 34
                Call ClearCapture()
                Call MakeUserChar(False, 1, 2, 34, 50, 90)
                UserList(2).pos.Map = 35
                UserList(2).Char.charindex = 3: CharList(3) = 2
                If id = 59 Then Call ClearCapture(1) Else Call ClearCapture()
                Call WriteCharacterCreate(1, 10, 20, e_Heading.SOUTH, 2, 50, 90, 30, 40, 60, 70, 0, 0, 50, "npc", 1, 1, 0, _
                    "", "", "", "", "", "", "", 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                Call Sample()
                If id = 59 Then
                    Call Expect(packets.Count = 3 And PacketIdAt(2) = eCharacterRemove, "failed NPC create cannot erase accepted active owner")
                Else
                    Call Expect(packets.Count = 2 And PacketIdAt(2) = eAdjacentPlayerUpsert, "accepted NPC replacement protects reused index from remove")
                End If
            #End If
        Case 61
            #If PYMMO = 1 Then
                NumChars = 2: LastChar = 2
                Call Sample()
                oldValue = AdjacentPlayerTestValue(1, 2)
                Call WarpUserChar(2, 35, 51, 90, False)
                Call Expect(UserList(2).pos.Map = 35 And UserList(2).pos.x = 51, "real one tile warp completed")
                Call Sample()
                value = AdjacentPlayerTestValue(1, 2)
                Call Expect(value(0) = oldValue(0) And value(1) = oldValue(1), "warp retains session identity")
                Call Expect(packets.Count = 1 And (FlagsAt(1) And 64) = 64, "real WarpUserChar forces snap")
                Call Sample()
                Call Expect(packets.Count = 0, "real warp has no redundant followup")
            #End If
        Case 62
            #If PYMMO = 1 Then
                UserList(2).pos.Map = 34
                Call ClearCapture()
                Call MakeUserChar(False, 1, 2, 34, 50, 90)
                ' Model a rejected legacy broadcast remove followed by a normal
                ' create at another index. Both accepted client indices survive.
                UserList(2).Char.charindex = 3: CharList(3) = 2
                Call MakeUserChar(False, 1, 2, 34, 50, 90)
                UserList(2).pos.Map = 35
                Call Sample(3)
                Call Expect(packets.Count = 3 And AdjacentPlayerTestCount(1) = 0, "both old indices must retire before adjacent upsert")
                Set reader = PacketReader(2, eCharacterRemove)
                Call Expect(reader.ReadInt16() = 2 And accepted(2), "first old normal index retired")
                Set reader = PacketReader(3, eCharacterRemove)
                Call Expect(reader.ReadInt16() = 3 And Not accepted(3), "second old normal index retained on failure")
                Call Sample()
                Call Expect(packets.Count = 2 And PacketIdAt(2) = eAdjacentPlayerUpsert, "accepted prefix not replayed")
                Set reader = PacketReader(1, eCharacterRemove)
                Call Expect(reader.ReadInt16() = 3, "retry only remaining normal index")
                Call Sample()
                Call Expect(packets.Count = 0, "reverse handoff settled")
            #End If
    End Select
End Sub

Public Function TestSuite() As Boolean
    On Error GoTo Failed
    Dim oldUsers() As t_User, oldChars(1 To 131) As Integer
    Dim oldMaps(0 To 3) As t_MapInfo, oldGroups(0 To 3) As t_ConnGroup
    Dim oldTiles() As t_MapBlock
    Dim maps As Variant, i As Long, x As Long, y As Long, id As Long, before As Long
    Dim oldLast As Integer, oldRequired As Byte, oldGuildCount As Integer, oldGuild As clsClan, saved As Boolean
    Dim oldRecent As t_RecentKillCache
    Dim oldNumChars As Integer, oldLastChar As Integer
    before = UnitTesting.GetFailedTests()
    ReDim oldUsers(1 To 131)
    ReDim oldTiles(1 To 100, 1 To 100, 0 To 3)
    Call Expect(UBound(UserList) >= 131 And NumMaps >= 80, "fixture capacity")
    maps = Array(34, 35, 78, 80)
    For i = 1 To 131
        oldUsers(i) = UserList(i): oldChars(i) = CharList(i)
    Next i
    For i = 0 To 3
        oldMaps(i) = MapInfo(maps(i)): oldGroups(i) = ConnGroups(maps(i))
        For x = 1 To 100
            For y = 1 To 100
                oldTiles(x, y, i) = MapData(x, y, maps(i))
            Next y
        Next x
    Next i
    oldLast = LastUser: oldRequired = RequiredGuildLevelSeeInvisible: oldGuildCount = CANTIDADDECLANES
    oldRecent = RecentDCUserCache
    oldNumChars = NumChars: oldLastChar = LastChar
    Set fixtureClan = New clsClan
    Call fixtureClan.SetTestLevel(5)
    Set oldGuild = modGuilds.SwapTestGuild(1, fixtureClan)
    If CANTIDADDECLANES < 1 Then CANTIDADDECLANES = 1
    saved = True
    Call InitializeAuxiliaryBuffer()
    capturing = True
    For id = 1 To 62
        Call RunCase(id)
    Next id
    TestSuite = UnitTesting.GetFailedTests() = before
Cleanup:
    capturing = False
    Call AdjacentPlayerTestReset()
    If saved Then
        For i = 1 To 131
            UserList(i) = oldUsers(i): CharList(i) = oldChars(i)
        Next i
        For i = 0 To 3
            MapInfo(maps(i)) = oldMaps(i): ConnGroups(maps(i)) = oldGroups(i)
            For x = 1 To 100
                For y = 1 To 100
                    MapData(x, y, maps(i)) = oldTiles(x, y, i)
                Next y
            Next x
        Next i
        LastUser = oldLast: RequiredGuildLevelSeeInvisible = oldRequired: CANTIDADDECLANES = oldGuildCount
        RecentDCUserCache = oldRecent
        NumChars = oldNumChars: LastChar = oldLastChar
        Set fixtureClan = modGuilds.SwapTestGuild(1, oldGuild)
    End If
    Call LoadAdjacentTopology()
    Exit Function
Failed:
    Call UnitTesting.RunTestError("adjacent player suite", CStr(Err.Number) & ": " & Err.Description)
    Resume Cleanup
End Function

Private Sub RunCase(ByVal id As Long)
    On Error GoTo Failed
    Call Fixture()
    Call CheckCase(id)
    Call UnitTesting.RunTest("adjacent player case " & CStr(id), True)
    Exit Sub
Failed:
    Call UnitTesting.RunTestError("adjacent player case " & CStr(id), CStr(Err.Number) & ": " & Err.Description)
End Sub
#End If
