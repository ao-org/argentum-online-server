Attribute VB_Name = "Unit_NpcBow"
Option Explicit
#If UNIT_TEST = 1 Then
Private capturing As Boolean
Private packets As Collection
Private failPacket As Integer
Private u As Integer, n As Integer, c As Integer, m As Integer
Private nowTick As Long
Private forceMissRoll As Boolean
Private failAmmoAfterShot As Boolean
Private failDamageAfterHit As Boolean
Private readerBytes() As Byte
Private forceHitRoll As Boolean
Private heldHeapCount As Integer

Public Function ForceHit(ByVal userIndex As Integer, ByVal npcIndex As Integer) As Boolean
    ForceHit = capturing And forceHitRoll And userIndex = u And npcIndex = n
End Function

Public Function ForceMiss(ByVal userIndex As Integer, ByVal npcIndex As Integer) As Boolean
    ForceMiss = capturing And forceMissRoll And userIndex = u And npcIndex = n
End Function

Public Sub InitializeTestWorld()
    Call InitializeCircularLogBuffer()
    Set SvrConfig = New ServerConfig
    Call SvrConfig.LoadSettings(App.Path & "\Example.Configuracion.ini")
    Call LoadFeatureToggles()
    ReDim UserList(1 To 2)
    ReDim ObjData(0 To 4000)
    ReDim MapInfo(1 To 300)
    ReDim MapData(1 To 100, 1 To 100, 1 To 300)
    ReDim QuestList(0 To 0)
    ReDim EffectOverTime(1 To 1)
    NumMaps = 300
    Call InitializeNpcIndexHeap()
    Call InitAreas()
    MinXBorder = XMinMapSize + (XWindow \ 2)
    MinYBorder = YMinMapSize + (YWindow \ 2)
    MaxXBorder = XMaxMapSize - (XWindow \ 2)
    MaxYBorder = YMaxMapSize - (YWindow \ 2)
    Call InitializeAuxiliaryBuffer()
End Sub

Public Function CapturePacket(ByVal route As Long, ByVal recipient As Integer, ByRef queued As Boolean) As Boolean
    If Not capturing Then Exit Function
    Dim writer As Network.writer, data() As Byte, value As Variant, packetId As Long
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.GetData(data)
    Call writer.Clear()
    If UBound(data) >= 1 Then packetId = CLng(data(0)) + 256& * data(1)
    queued = packetId <> failPacket
    If failAmmoAfterShot And packetId = ServerPacketID.eCreateProjectile Then
        ' Exercise QuitarUserInvItem's real rejected-removal return after combat.
        UserList(u).CurrentInventorySlots = 1
    End If
    If failDamageAfterHit And packetId = ServerPacketID.ePlayWave Then
        ' Inject a post-commit lookup error into the real damage routine.
        UserList(u).invent.EquippedWeaponObjIndex = UBound(ObjData) + 1
    End If
    If queued Then
        value = data
        Call packets.Add(value)
    End If
    CapturePacket = True
End Function

Private Sub Expect(ByVal condition As Boolean, ByVal text As String)
    If Not condition Then Err.Raise vbObjectError + 181, "NPC bow test", text
End Sub

Private Function IdAt(ByVal index As Long) As Long
    Dim data() As Byte
    data = packets(index)
    IdAt = CLng(data(0)) + 256& * data(1)
End Function

Private Function CountPacket(ByVal packetId As Long) As Long
    Dim i As Long
    For i = 1 To packets.Count
        If IdAt(i) = packetId Then CountPacket = CountPacket + 1
    Next i
End Function

Private Sub ExpectLong(ByRef data() As Byte, ByVal offset As Long, ByVal value As Long)
    Dim unsigned As Double, i As Long, part As Long
    unsigned = value
    If unsigned < 0 Then unsigned = unsigned + 4294967296#
    For i = 0 To 3
        part = unsigned - Fix(unsigned / 256#) * 256#
        Call Expect(data(offset + i) = part, "little-endian field at " & CStr(offset + i))
        unsigned = Fix(unsigned / 256#)
    Next i
End Sub

Private Sub Fixture()
    Call ClearNpcBow(u)
    Dim emptyUser As t_User, emptyNpc As t_Npc, emptyObj As t_ObjData
    UserList(u) = emptyUser
    NpcList(n) = emptyNpc
    ObjData(1) = emptyObj
    ObjData(2) = emptyObj
    failPacket = -1
    forceMissRoll = False
    forceHitRoll = False
    failAmmoAfterShot = False
    failDamageAfterHit = False
    Set packets = New Collection
    nowTick = GetTickCountRaw()
    With UserList(u)
        .Id = 101
        .VersionId = 7
        .name = "bow fixture"
        .flags.UserLogged = True
        .flags.Privilegios = e_PlayerType.User
        .ConnectionDetails.ConnIDValida = True
        .pos.Map = m: .pos.x = 50: .pos.y = 50
        .clase = e_Class.Hunter
        .Stats.MinSta = 100: .Stats.MaxSta = 100
        .Stats.MinHp = 100: .Stats.MaxHp = 100
        .CurrentInventorySlots = 20
        .invent.EquippedWeaponSlot = 1: .invent.EquippedWeaponObjIndex = 1
        .invent.EquippedMunitionSlot = 2: .invent.EquippedMunitionObjIndex = 2
        .invent.Object(1).ObjIndex = 1: .invent.Object(1).amount = 1: .invent.Object(1).Equipped = 1
        .invent.Object(2).ObjIndex = 2: .invent.Object(2).amount = 20: .invent.Object(2).Equipped = 1
        .Counters.TimerLanzarSpell = AddMod32(nowTick, -100000)
        .Counters.TimerGolpeMagia = AddMod32(nowTick, -100000)
        .Counters.TimerPuedeUsarArco = AddMod32(nowTick, -100000)
        .CdTimes(0) = AddMod32(nowTick, -100000)
    End With
    ObjData(1).OBJType = otWeapon: ObjData(1).Proyectil = 1: ObjData(1).Municion = 1
    ObjData(1).MinHit = 1: ObjData(1).MaxHit = 1
    ObjData(2).OBJType = otArrows: ObjData(2).Subtipo = 1: ObjData(2).ProjectileType = 1
    With NpcList(n)
        .flags.NPCActive = True
        .flags.AttackableByEveryone = 1
        .VersionId = 3
        .Char.charindex = c: .Char.body = 1
        ReDim .Char.Animation(0 To 0)
        .pos.Map = m: .pos.x = 51: .pos.y = 50
        .Stats.MinHp = 100000: .Stats.MaxHp = 100000
        .Numero = 1
        .Attackable = 1
        .npcType = DummyTarget
    End With
    CharList(c) = n
    MapData(51, 50, m).NpcIndex = n
    MapData(51, 50, m).Trigger = 0
    MapData(50, 50, m).Trigger = 0
End Sub

Private Function Request(ByVal requestId As Long) As t_NpcBowRequest
    With UserList(u).NpcBow
        Request.Context = .Context
        Request.RequestId = requestId
        Request.CharIndex = c
        Request.Token = .Bindings(.BindingIndex(n)).Token
    End With
End Function

Private Sub Attack(ByRef req As t_NpcBowRequest, ByRef reply As t_NpcBowReply)
    Call ProcessNpcBowAttack(u, req, GetTickCountRaw(), reply)
End Sub

Private Function AttackReader(ByRef req As t_NpcBowRequest, Optional ByVal trailing As Boolean = False) As Network.reader
    Dim writer As New Network.writer, borrowed() As Byte, i As Long
    Call writer.WriteInt32(req.Context.Low)
    Call writer.WriteInt32(req.Context.High)
    Call writer.WriteInt32(req.RequestId)
    Call writer.WriteInt16(req.CharIndex)
    Call writer.WriteInt32(req.Token.Low)
    Call writer.WriteInt32(req.Token.High)
    If trailing Then Call writer.WriteInt8(0)
    ' Both GetData and Reader borrow memory. Copy out of the local writer first.
    Call writer.GetData(borrowed)
    ReDim readerBytes(0 To UBound(borrowed))
    For i = 0 To UBound(borrowed)
        readerBytes(i) = borrowed(i)
    Next i
    Set AttackReader = New Network.reader
    Call AttackReader.SetData(readerBytes)
End Function

Private Sub TestCase(ByVal index As Integer)
    Dim req As t_NpcBowRequest, reply As t_NpcBowReply, other As t_NpcBowRequest
    Dim data() As Byte, i As Long, count As Long, sta As Integer, committed As Boolean, finished As Boolean
    Dim reader As Network.reader, ctx As t_BowIdentity, hp As Long, tick As Long
    Dim oldTile As t_MapBlock, otherNpc As t_Npc, spareNpc As Integer
    Dim oldEffect As t_EffectOverTime, oldClass As t_ModClase, effect As EffectProvoked
    Dim oldLast As Integer, oldCount As Integer, oldLastChar As Integer, released As Boolean
    If index = 10 Then failPacket = ServerPacketID.eCharacterCreate
    Call SubscribeNpcBow(u)
    If index <> 10 Then
        Call Expect(UserList(u).NpcBow.Ready, "subscription ready")
        req = Request(1)
    End If
    Select Case index
        Case 1
            Call Expect(IdAt(1) = ServerPacketID.eHooBowContext, "Begin first")
            Call Expect(IdAt(2) = ServerPacketID.eCharacterRemove, "old client index removed")
            Call Expect(IdAt(3) = ServerPacketID.eCharacterCreate, "authoritative Create")
            Call Expect(IdAt(4) = ServerPacketID.eHooBowBinding, "binding follows Create")
            Call Expect(IdAt(packets.Count) = ServerPacketID.eHooBowContext, "Ready last")
            data = packets(1): Call Expect(UBound(data) = 10 And data(10) = 0, "11 byte Begin")
            Call ExpectLong(data, 2, req.Context.Low): Call ExpectLong(data, 6, req.Context.High)
            data = packets(4): Call Expect(UBound(data) = 20 And data(10) = 1, "21 byte Bind")
            Call ExpectLong(data, 2, req.Context.Low): Call ExpectLong(data, 6, req.Context.High)
            Call Expect(CLng(data(11)) + 256& * data(12) = c, "binding character little endian")
            Call ExpectLong(data, 13, req.Token.Low): Call ExpectLong(data, 17, req.Token.High)
            data = packets(packets.Count): Call Expect(data(10) = 1, "Ready phase")
        Case 2
            count = packets.Count
            Call SubscribeNpcBow(u)
            other = Request(2)
            Call Expect(req.Token.Low = other.Token.Low And req.Token.High = other.Token.High, "duplicate subscription keeps token")
            Call Expect(packets.Count = count + 1, "only Ready replayed")
        Case 3
            Call MakeNPCChar(False, u, n, m, 51, 50)
            other = Request(2)
            Call Expect(req.Token.Low = other.Token.Low And req.Token.High = other.Token.High, "same actor refresh retains binding")
        Case 4
            UserList(u).pos.y = 1
            Call PruneNpcBowUser(u)
            Call Expect(Not UserList(u).NpcBow.BindingIndex.Exists(n), "new area prunes departed actor")
            Call Attack(req, reply): Call Expect(reply.Code = BowInvalidTarget, "departed actor rejected")
        Case 5
            NpcList(n).pos.y = 100
            Call PruneNpcBowActor(n)
            Call Expect(Not UserList(u).NpcBow.BindingIndex.Exists(n), "NPC movement prunes without user movement")
        Case 6
            Call SetUserRef(NpcList(n).MaestroUser, u)
            Call RevokeNpcBow(n)
            Call Attack(req, reply): Call Expect(reply.Code = BowInvalidTarget, "new pet revokes")
            Call MakeNPCChar(False, u, n, m, 51, 50)
            Call Expect(Not UserList(u).NpcBow.BindingIndex.Exists(n), "pet cannot rebind")
        Case 7
            NpcList(n).VersionId = 32761
            Call IncreaseNpcVersionId(n)
            Call Expect(NpcList(n).VersionId = 0, "legacy version wrapped")
            NpcList(n).VersionId = 3
            Call MakeNPCChar(False, u, n, m, 51, 50)
            other = Request(2)
            Call Expect(other.Token.Low <> req.Token.Low Or other.Token.High <> req.Token.High, "fresh token even identical old version")
            Call Attack(req, reply): Call Expect(reply.Code = BowInvalidTarget, "old lifetime cannot attack reused slot")
        Case 8
            NpcList(n).Stats.MinHp = 0
            Call Attack(req, reply): Call Expect(reply.Code = BowInvalidTarget, "dead actor rejected before release")
        Case 9
            NpcList(n).Char.charindex = c - 1
            Call Attack(req, reply): Call Expect(reply.Code = BowInvalidTarget, "live char index rechecked")
        Case 10
            Call Expect(Not UserList(u).NpcBow.Ready, "failed Create cannot complete barrier")
            Call Expect(CountPacket(ServerPacketID.eHooBowBinding) = 0, "no binding without delivered Create")
        Case 11
            req.Context.Low = AddMod32(req.Context.Low, 1)
            Call Attack(req, reply): Call Expect(reply.Code = BowInvalidContext, "foreign context")
        Case 12
            req.Token.High = AddMod32(req.Token.High, 1)
            Call Attack(req, reply): Call Expect(reply.Code = BowInvalidTarget, "wrong token")
        Case 13
            Call Attack(req, reply)
            Call Expect(reply.Code = BowAccepted And reply.Committed, "attempt accepted and committed")
            sta = UserList(u).Stats.MinSta
            count = packets.Count
            Call Attack(req, reply)
            Call Expect(reply.Code = BowAccepted And reply.Committed And UserList(u).Stats.MinSta = sta, "duplicate replays commitment without stamina")
            Call Expect(UserList(u).invent.Object(2).amount = 19 And packets.Count = count, "one ammo and no repeated effects")
        Case 14
            Call Attack(req, reply)
            req.Token.Low = AddMod32(req.Token.Low, 1)
            Call Attack(req, reply): Call Expect(reply.Code = BowInvalidTarget, "altered duplicate rejected")
            Call Expect(UserList(u).invent.Object(2).amount = 19, "altered replay costs nothing")
        Case 15
            other = req
            For i = 1 To 18
                req.RequestId = i
                req.Token.High = -1
                Call Attack(req, reply)
            Next i
            Call Attack(other, reply): Call Expect(reply.Code = BowInvalidTarget, "old uncached request fails closed")
        Case 16
            UserList(u).Stats.MinSta = 9
            Call Attack(req, reply): Call Expect(reply.Code = BowNoEnergy And Not reply.Committed, "no energy before commitment")
        Case 17
            UserList(u).invent.Object(2).amount = 0
            Call Attack(req, reply): Call Expect(reply.Code = BowNoAmmo, "no ammo")
        Case 18
            ObjData(2).Subtipo = 2
            Call Attack(req, reply): Call Expect(reply.Code = BowNoAmmo, "wrong ammo subtype")
        Case 19
            UserList(u).invent.Object(1).Equipped = 0
            Call Attack(req, reply): Call Expect(reply.Code = BowInvalidEquipment, "unequipped weapon")
        Case 20
            UserList(u).flags.Descansar = True
            Call Attack(req, reply): Call Expect(reply.Code = BowActorBlocked, "resting actor")
        Case 21
            NpcList(n).OnlyForGuilds = 1
            Call Attack(req, reply): Call Expect(reply.Code = BowForbiddenTarget, "guild restriction")
            Call UsuarioAtacaNpc(u, n, Ranged, committed, finished, True)
            Call Expect(Not committed And Not finished, "actual early exit reports no attempt")
        Case 22
            NpcList(n).Attackable = 0
            Call Attack(req, reply): Call Expect(reply.Code = BowForbiddenTarget, "immune NPC")
        Case 23
            UserList(u).flags.Montado = 1
            Call Attack(req, reply): Call Expect(reply.Code = BowActorBlocked, "production mounted restriction")
        Case 24
            NpcList(n).pos.x = 50 + RANGO_VISION_X + 1
            MapData(NpcList(n).pos.x, 50, m).NpcIndex = n
            Call Attack(req, reply): Call Expect(reply.Code = BowOutOfRange, "range uses either axis")
            MapData(NpcList(n).pos.x, 50, m).NpcIndex = 0
        Case 25
            UserList(u).Counters.TimerPuedeUsarArco = nowTick
            Call ProcessNpcBowAttack(u, req, nowTick, reply)
            Call Expect(reply.Code = BowRateLimited And reply.RetryAfter = IntervaloFlechasCazadores, "bow delay")
        Case 26
            UserList(u).Counters.TimerPuedeUsarArco = AddMod32(nowTick, -50)
            UserList(u).Counters.TimerLanzarSpell = AddMod32(nowTick, -100)
            UserList(u).Counters.TimerGolpeMagia = AddMod32(nowTick, -150)
            ObjData(1).Cooldown = 1900
            UserList(u).CdTimes(0) = AddMod32(nowTick, -300)
            Call Expect(NpcBowRetryAfter(u, nowTick) = 1600, "maximum actual remaining gate")
        Case 27
            count = packets.Count
            Set reader = AttackReader(req, True)
            Call HandleNpcBowAttack(u, reader)
            Call Expect(packets.Count = count And UserList(u).invent.Object(2).amount = 20, "trailing bytes no mutation")
            Set reader = New Network.reader
            ReDim data(0 To 3)
            Call reader.SetData(data)
            Call HandleNpcBowAttack(u, reader)
            Call Expect(packets.Count = count, "truncated input no mutation")
        Case 28
            Set reader = AttackReader(req)
            Call Expect(reader.ReadInt32() = req.Context.Low, "reader context low matches fixture")
            Call Expect(reader.ReadInt32() = req.Context.High, "reader context high matches fixture")
            Call reader.SetData(readerBytes)
            Call HandleNpcBowAttack(u, reader)
            Call Expect(IdAt(packets.Count) = ServerPacketID.eHooBowResult, "real handler result")
            data = packets(packets.Count)
            Call Expect(UBound(data) = 18 And data(14) = BowAccepted, "19 byte accepted result; upper=" & CStr(UBound(data)) & "; code=" & CStr(data(14)) & "; request=" & CStr(UserList(u).NpcBow.LastRequest))
            Call ExpectLong(data, 2, req.Context.Low): Call ExpectLong(data, 6, req.Context.High)
            Call ExpectLong(data, 10, req.RequestId)
            #If PYMMO = 1 Then
                Call Expect(ClientPacketID.eHooBowSubscribe = 315 And ClientPacketID.eHooBowAttack = 316 And ServerPacketID.eHooBowResult = 213, "HOO IDs")
            #Else
                Call Expect(ClientPacketID.eHooBowSubscribe = 318 And ClientPacketID.eHooBowAttack = 319 And ServerPacketID.eHooBowResult = 214, "conditional IDs")
            #End If
        Case 29
            Set reader = AttackReader(req)
            Call HandleNpcBowAttack(u, reader)
            For i = 1 To 40
                Set reader = AttackReader(req)
                Call HandleNpcBowAttack(u, reader)
            Next i
            Call Expect(CountPacket(ServerPacketID.eHooBowResult) = 32, "duplicates count toward abuse budget")
            Call Expect(UserList(u).invent.Object(2).amount = 19, "abuse never duplicates execution")
        Case 30
            Call WriteChangeMap(u, m)
            Call Expect(Not UserList(u).NpcBow.Subscribed And UserList(u).NpcBow.WantsSubscription, "map retires context and keeps intent")
            Call ResumeNpcBowAfterArea(u)
            other = Request(2)
            Call Expect(other.Context.Low <> req.Context.Low Or other.Context.High <> req.Context.High, "new map context")
            Call Attack(req, reply): Call Expect(reply.Code = BowInvalidContext, "old map request rejected")
        Case 31
            Call SuspendNpcBow(u)
            UserList(u).flags.Muerto = 1
            Call SubscribeNpcBow(u)
            Call Expect(Not UserList(u).NpcBow.Subscribed And UserList(u).NpcBow.WantsSubscription, "dead subscription waits")
            UserList(u).flags.Muerto = 0
            Call ResumeNpcBowAfterArea(u)
            Call Expect(UserList(u).NpcBow.Ready And UserList(u).NpcBow.BindingIndex.Exists(n), "resurrection restores barrier")
        Case 32
            UserList(u).ConnectionDetails.ConnIDValida = False
            Call CloseSocketSL(u)
            Call Expect(Not UserList(u).NpcBow.Subscribed And Not UserList(u).NpcBow.WantsSubscription, "disconnect clears intent")
            Call Attack(req, reply): Call Expect(reply.Code = BowInvalidContext, "old disconnected request")
        Case 33
            Call IncreaseVersionId(u)
            Call SubscribeNpcBow(u)
            Call Attack(req, reply): Call Expect(reply.Code = BowInvalidContext, "new character session rejects old context")
        Case 34
            UserList(u).Counters.TimerLanzarSpell = &H7FFFFFF0
            Call Expect(NpcBowRetryAfter(u, &H80000010) = IntervaloMagiaGolpe - 32, "unsigned tick wrap")
        Case 35
            Call UsuarioAtacaNpc(u, n, Ranged, committed, finished)
            Call Expect(committed And finished, "actual hit or miss is committed")
            Call Expect(UserList(u).Stats.MinSta = 100 And UserList(u).invent.Object(2).amount = 20, "legacy combat caller costs unchanged")
        Case 36
            NpcList(n).flags.invisible = 1
            Call PruneNpcBowActor(n)
            Call Expect(Not UserList(u).NpcBow.BindingIndex.Exists(n), "hidden NPC revoked")
        Case 37
            Call Expect(BowTestIdentityWrap(), "opaque 64 bit carry and exhaustion")
        Case 38
            ' Occupancy at the old tile does not identify the target.
            spareNpc = IIf(n = MaxNPCs, n - 1, n + 1): otherNpc = NpcList(spareNpc)
            NpcList(spareNpc).Stats.MinHp = 12345
            oldTile = MapData(52, 50, m)
            MapData(51, 50, m).NpcIndex = spareNpc
            MapData(51, 50, m).UserIndex = u
            MapData(52, 50, m).NpcIndex = n
            NpcList(n).pos.x = 52
            Call PruneNpcBowActor(n)
            Call Attack(req, reply)
            hp = NpcList(spareNpc).Stats.MinHp
            MapData(52, 50, m) = oldTile
            NpcList(spareNpc) = otherNpc
            Call Expect(reply.Code = BowAccepted And hp = 12345, "identity follows moved NPC, not old occupants")
            For i = 1 To packets.Count
                If IdAt(i) = ServerPacketID.eCreateProjectile Then
                    data = packets(i)
                    Call Expect(data(4) = 52 And data(5) = 50, "projectile uses current NPC tile")
                End If
            Next i
        Case 39
            UserList(u).Stats.MinSta = 9
            Call Attack(req, reply): Call Expect(reply.Code = BowNoEnergy, "initial rejection")
            UserList(u).Stats.MinSta = 100
            Call Attack(req, reply): Call Expect(reply.Code = BowNoEnergy, "rejected replay unchanged")
            Call Expect(UserList(u).invent.Object(2).amount = 20 And UserList(u).Stats.MinSta = 100, "rejected replay free")
            req.RequestId = 2
            Call Attack(req, reply): Call Expect(reply.Code = BowAccepted, "fresh request can succeed")
        Case 40
            UserList(u).invent.Object(2).amount = 1
            Call Attack(req, reply): Call Expect(reply.Code = BowAccepted, "last arrow accepted")
            Call Expect(UserList(u).invent.Object(2).amount = 0 And UserList(u).invent.Object(2).ObjIndex = 0, "last arrow removed")
            Call Expect(UserList(u).invent.EquippedMunitionSlot = 0 And UserList(u).invent.EquippedMunitionObjIndex = 0, "ammo equipment cleared")
            count = packets.Count: sta = UserList(u).Stats.MinSta
            Call Attack(req, reply)
            Call Expect(reply.Code = BowAccepted And packets.Count = count And UserList(u).Stats.MinSta = sta, "depleted duplicate harmless")
        Case 41
            ObjData(1).Municion = 0
            UserList(u).invent.EquippedMunitionSlot = 0
            UserList(u).invent.EquippedMunitionObjIndex = 0
            Call Attack(req, reply): Call Expect(reply.Code = BowAccepted, "ammo-free bow accepted")
            tick = UserList(u).CdTimes(0): count = packets.Count
            Call Attack(req, reply)
            Call Expect(UserList(u).invent.Object(2).amount = 20 And UserList(u).CdTimes(0) = tick And packets.Count = count, "ammo-free replay preserves inventory and cooldown")
        Case 42
            req.RequestId = &H7FFFFFFF
            UserList(u).Stats.MinSta = 9
            Call Attack(req, reply): Call Expect(reply.Code = BowNoEnergy, "maximum signed request ID")
            UserList(u).Stats.MinSta = 100
            Call Attack(req, reply): Call Expect(reply.Code = BowNoEnergy, "maximum ID replays")
            req.RequestId = 1
            Call Attack(req, reply): Call Expect(reply.Code = BowInvalidTarget, "wrapped request cannot execute")
            Call Expect(UserList(u).invent.Object(2).amount = 20, "wrapped request costs nothing")
        Case 43
            forceMissRoll = True
            hp = NpcList(n).Stats.MinHp
            Call Attack(req, reply): Call Expect(reply.Code = BowAccepted, "guaranteed miss is accepted attempt")
            Call Expect(NpcList(n).Stats.MinHp = hp And UserList(u).invent.Object(2).amount = 19, "miss preserves HP and consumes one arrow")
            count = packets.Count: sta = UserList(u).Stats.MinSta
            Call Attack(req, reply)
            Call Expect(packets.Count = count And UserList(u).Stats.MinSta = sta, "miss replay no effects or cost")
        Case 44
            failAmmoAfterShot = True
            Call Attack(req, reply): Call Expect(reply.Code = BowInternalError And reply.Committed, "unconfirmed ammo removal is uncertain after commitment")
            Call Expect(UserList(u).invent.Object(2).amount = 20, "failed helper did not consume")
            count = packets.Count: sta = UserList(u).Stats.MinSta: hp = NpcList(n).Stats.MinHp
            UserList(u).CurrentInventorySlots = 20
            failAmmoAfterShot = False
            Call Attack(req, reply)
            Call Expect(reply.Code = BowInternalError And packets.Count = count And UserList(u).Stats.MinSta = sta And NpcList(n).Stats.MinHp = hp, "post-commit failure is terminal and cached")
        Case 45
            reply.Request.Context.Low = &H89ABCDEF
            reply.Request.Context.High = &HFEDCBA98
            reply.Request.RequestId = &H12345678
            reply.Code = BowRateLimited
            reply.RetryAfter = &H1020304
            Call WriteNpcBowResult(u, reply)
            data = packets(packets.Count)
            Call Expect(UBound(data) = 18 And IdAt(packets.Count) = ServerPacketID.eHooBowResult And data(14) = 1, "explicit result fixture")
            Call ExpectLong(data, 2, &H89ABCDEF): Call ExpectLong(data, 6, &HFEDCBA98)
            Call ExpectLong(data, 10, &H12345678): Call ExpectLong(data, 15, &H1020304)
        Case 46
            UserList(u).flags.Inmunidad = 1
            Call Attack(req, reply): Call Expect(reply.Code = BowActorBlocked, "actor immunity is recoverable actor state")
        Case 47
            UserList(u).flags.EnConsulta = True
            Call Attack(req, reply): Call Expect(reply.Code = BowActorBlocked, "consultation is recoverable actor state")
        Case 48
            oldTile = MapData(51, 50, MAP_CAPTURE_THE_FLAG_1)
            UserList(u).pos.Map = MAP_CAPTURE_THE_FLAG_1
            NpcList(n).pos.Map = MAP_CAPTURE_THE_FLAG_1
            MapData(51, 50, MAP_CAPTURE_THE_FLAG_1).NpcIndex = n
            Call SubscribeNpcBow(u)
            req = Request(1)
            Call Attack(req, reply)
            MapData(51, 50, MAP_CAPTURE_THE_FLAG_1) = oldTile
            Call Expect(reply.Code = BowAccepted And UserList(u).invent.Object(2).amount = 20, "free consumable area keeps arrows")
            Call Expect(UserList(u).Stats.MinSta < 100, "free area still pays stamina")
            count = packets.Count: sta = UserList(u).Stats.MinSta
            Call Attack(req, reply)
            Call Expect(packets.Count = count And UserList(u).Stats.MinSta = sta, "free area duplicate no costs")
        Case 49
            oldEffect = EffectOverTime(1): oldClass = ModClase(e_Class.Hunter)
            oldLast = LastNPC: oldCount = NumNPCs: oldLastChar = LastChar
            EffectOverTime(1).PhysicalDamageDone = -0.5
            EffectOverTime(1).MagicDamageDone = -0.25
            EffectOverTime(1).ClientEffectTypeId = 0
            Set effect = New EffectProvoked
            Call effect.Setup(u, eUser, u, eUser, 1, 1)
            With UserList(u).EffectOverTime
                ReDim .EffectList(0 To 0)
                Set .EffectList(0) = effect
                .EffectCount = 1
                .CallbaclMask = eTargetWillAttack Or eTargetDidHit Or eTargetFailedAttack
            End With
            ModClase(e_Class.Hunter).DañoProyectiles = 1
            ObjData(1).MinHit = 100: ObjData(1).MaxHit = 100
            NpcList(n).npcType = Comun
            NpcList(n).flags.Respawn = 1
            NpcList(n).Stats.MinHp = 1
            forceHitRoll = True
            Call Attack(req, reply)
            released = Not NpcList(n).flags.NPCActive
            EffectOverTime(1) = oldEffect: ModClase(e_Class.Hunter) = oldClass
            LastNPC = oldLast: NumNPCs = oldCount: LastChar = oldLastChar
            If GetAvailableNpcIndex() > heldHeapCount Then Call Expect(GetNextAvailableNpc() = n, "reclaim isolated killed fixture index")
            Call Expect(reply.Code = BowAccepted And released, "real killing arrow retires NPC; code=" & CStr(reply.Code) & "; HP=" & CStr(NpcList(n).Stats.MinHp) & "; physical=" & CStr(UserList(u).Modifiers.PhysicalDamageBonus))
            Call Expect(Abs(UserList(u).Modifiers.PhysicalDamageBonus) < 0.00001 And Abs(UserList(u).Modifiers.MagicDamageBonus) < 0.00001, "killing hit restores provocation modifiers")
            Call Expect(UserList(u).invent.Object(2).amount = 19, "killing arrow consumed once")
        Case 50
            UserList(u).flags.Muerto = 1
            For i = 1 To 40
                If AllowNpcBowMessage(u, nowTick) Then Call SubscribeNpcBow(u)
            Next i
            Call Expect(UserList(u).NpcBow.WindowCount = 32, "dead resubscriptions retain abuse budget")
        Case 51
            forceHitRoll = True
            failDamageAfterHit = True
            Call Attack(req, reply)
            UserList(u).invent.EquippedWeaponObjIndex = 1
            failDamageAfterHit = False
            Call Expect(reply.Code = BowInternalError, "damage helper error cannot report Accepted")
            count = packets.Count: sta = UserList(u).Stats.MinSta: hp = NpcList(n).Stats.MinHp
            Call Attack(req, reply)
            Call Expect(reply.Code = BowInternalError And count = packets.Count And sta = UserList(u).Stats.MinSta And hp = NpcList(n).Stats.MinHp, "uncertain damage failure replays without execution")
            Call Expect(UserList(u).invent.Object(2).amount = 20, "damage failure does not continue ammo completion")
        Case 52
            With UserList(u).NpcBow
                .Context.Low = &H89ABCDEF: .Context.High = &HFEDCBA98
                .Bindings(.BindingIndex(n)).Token.Low = &H87654321
                .Bindings(.BindingIndex(n)).Token.High = &HABCDEF01
            End With
            Call SubscribeNpcBow(u)
            data = packets(packets.Count)
            Call Expect(UBound(data) = 10 And IdAt(packets.Count) = ServerPacketID.eHooBowContext And data(10) = 1, "explicit ready fixture")
            Call ExpectLong(data, 2, &H89ABCDEF): Call ExpectLong(data, 6, &HFEDCBA98)
            Call MakeNPCChar(False, u, n, m, 51, 50)
            For i = packets.Count To 1 Step -1
                If IdAt(i) = ServerPacketID.eHooBowBinding Then Exit For
            Next i
            data = packets(i)
            Call Expect(UBound(data) = 20 And data(10) = 1, "explicit binding fixture")
            Call ExpectLong(data, 2, &H89ABCDEF): Call ExpectLong(data, 6, &HFEDCBA98)
            Call Expect(CLng(data(11)) + 256& * data(12) = c, "explicit binding char")
            Call ExpectLong(data, 13, &H87654321): Call ExpectLong(data, 17, &HABCDEF01)
            Call RevokeNpcBow(n)
            data = packets(packets.Count)
            Call Expect(UBound(data) = 20 And IdAt(packets.Count) = ServerPacketID.eHooBowBinding And data(10) = 2, "explicit revoke fixture")
            Call ExpectLong(data, 2, &H89ABCDEF): Call ExpectLong(data, 6, &HFEDCBA98)
            Call ExpectLong(data, 13, &H87654321): Call ExpectLong(data, 17, &HABCDEF01)
        Case 53
            #If DEBUGGING = 1 Then
                Dim diagnostic As String
                reply.Request = req: reply.Code = BowInternalError: reply.Committed = True
                UserList(u).NpcBow.WindowCount = 1
                diagnostic = NpcBowDiagnostic(u, reply)
                Call Expect(InStr(diagnostic, "context=") > 0 And InStr(diagnostic, "request=1") > 0 And InStr(diagnostic, "target=" & CStr(c)) > 0, "trace correlates identity")
                Call Expect(InStr(diagnostic, "token=") > 0 And InStr(diagnostic, "result=10 retry_ms=0 committed=1") > 0, "trace records uncertain committed outcome")
                UserList(u).NpcBow.WindowCount = 5
                Call Expect(LenB(NpcBowDiagnostic(u, reply)) = 0, "trace bounded independently of gameplay")
            #End If
    End Select
    If index >= 16 And index <= 25 Then
        Call Expect(UserList(u).invent.Object(2).amount = IIf(index = 17, 0, 20), "rejection preserves ammo")
        Call Expect(UserList(u).Stats.MinSta = IIf(index = 16, 9, 100), "rejection preserves energy")
    End If
End Sub

Private Sub RunCase(ByVal index As Integer)
    On Error GoTo Failed
    Call Fixture()
    Call TestCase(index)
    Call UnitTesting.RunTest("NPC bow " & CStr(index), True)
    Exit Sub
Failed:
    Call UnitTesting.RunTestError("NPC bow " & CStr(index), CStr(Err.Number) & ": " & Err.Description & vbCrLf & Join(GetLastMessages(), vbCrLf))
End Sub

Public Function test_suite_npc_bow() As Boolean
    On Error GoTo Failed
    Dim oldUser As t_User, oldNpc As t_Npc, old1 As t_ObjData, old2 As t_ObjData
    Dim oldMap As t_MapInfo, oldTiles() As t_MapBlock
    Dim saved As Boolean, x As Integer, y As Integer, i As Long, before As Long
    Dim oldBow As Long, oldMG As Long, oldGM As Long, oldChar As Integer
    before = UnitTesting.GetFailedTests()
    For i = UBound(UserList) To 1 Step -1
        If Not UserList(i).flags.UserLogged And Not UserList(i).ConnectionDetails.ConnIDValida Then u = i: Exit For
    Next i
    n = GetNextAvailableNpc()
    heldHeapCount = GetAvailableNpcIndex()
    For i = UBound(CharList) To 1 Step -1
        If CharList(i) = 0 Then c = i: Exit For
    Next i
    m = NumMaps
    Call Expect(u > 0 And n > 0 And c > 0 And m > 0, "isolated fixture slots")
    oldUser = UserList(u): oldNpc = NpcList(n): oldChar = CharList(c)
    old1 = ObjData(1): old2 = ObjData(2): oldMap = MapInfo(m)
    ReDim oldTiles(1 To 100, 1 To 100)
    For x = 1 To 100
        For y = 1 To 100
            oldTiles(x, y) = MapData(x, y, m)
        Next y
    Next x
    oldBow = IntervaloFlechasCazadores: oldMG = IntervaloMagiaGolpe: oldGM = IntervaloGolpeMagia
    saved = True
    IntervaloFlechasCazadores = 500: IntervaloMagiaGolpe = 700: IntervaloGolpeMagia = 800
    capturing = True
    For i = 1 To 53
        Call RunCase(i)
    Next i
    test_suite_npc_bow = UnitTesting.GetFailedTests() = before
Cleanup:
    If saved Then
        Call ClearNpcBow(u)
        If GetAvailableNpcIndex() = heldHeapCount Then
            NpcList(n).flags.NPCActive = True
            Call ReleaseNpc(n, eDie)
        End If
        UserList(u) = oldUser: NpcList(n) = oldNpc: CharList(c) = oldChar
        ObjData(1) = old1: ObjData(2) = old2: MapInfo(m) = oldMap
        For x = 1 To 100
            For y = 1 To 100
                MapData(x, y, m) = oldTiles(x, y)
            Next y
        Next x
        IntervaloFlechasCazadores = oldBow: IntervaloMagiaGolpe = oldMG: IntervaloGolpeMagia = oldGM
    End If
    capturing = False
    Exit Function
Failed:
    Call UnitTesting.RunTestError("NPC bow suite", CStr(Err.Number) & ": " & Err.Description)
    Resume Cleanup
End Function
#End If
