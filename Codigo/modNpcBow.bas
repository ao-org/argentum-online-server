Attribute VB_Name = "modNpcBow"
Option Explicit

Private nextIdentity As t_BowIdentity
Private identitiesExhausted As Boolean
Private subscribers As New Dictionary
Private Const REQUESTS_PER_SECOND As Integer = 32

' One process-wide unsigned 64-bit sequence covers contexts and binding tokens.
' A restarted server has no surviving transport sessions; exhaustion fails closed.
Public Function AllocateBowIdentity(ByRef value As t_BowIdentity) As Boolean
    If identitiesExhausted Then Exit Function
    If nextIdentity.Low = -1 And nextIdentity.High = -1 Then
        identitiesExhausted = True
        Exit Function
    End If
    nextIdentity.Low = AddMod32(nextIdentity.Low, 1)
    If nextIdentity.Low = 0 Then nextIdentity.High = AddMod32(nextIdentity.High, 1)
    value = nextIdentity
    AllocateBowIdentity = True
End Function

Private Function SameIdentity(ByRef a As t_BowIdentity, ByRef b As t_BowIdentity) As Boolean
    SameIdentity = a.Low = b.Low And a.High = b.High
End Function

Public Sub ClearNpcBow(ByVal userIndex As Integer, Optional ByVal PreserveBudget As Boolean = False)
    Dim emptyState As t_NpcBowState
    If PreserveBudget Then
        emptyState.WindowStart = UserList(userIndex).NpcBow.WindowStart
        emptyState.WindowCount = UserList(userIndex).NpcBow.WindowCount
    End If
    UserList(userIndex).NpcBow = emptyState
    If subscribers.Exists(userIndex) Then Call subscribers.Remove(userIndex)
End Sub

#If UNIT_TEST = 1 Then
Public Function BowTestIdentityWrap() As Boolean
    Dim saved As t_BowIdentity, wasExhausted As Boolean, value As t_BowIdentity
    saved = nextIdentity
    wasExhausted = identitiesExhausted
    identitiesExhausted = False
    nextIdentity.Low = &H7FFFFFFF: nextIdentity.High = 5
    If Not AllocateBowIdentity(value) Then GoTo Done
    If value.Low <> &H80000000 Or value.High <> 5 Then GoTo Done
    nextIdentity.Low = -1: nextIdentity.High = 5
    If Not AllocateBowIdentity(value) Then GoTo Done
    If value.Low <> 0 Or value.High <> 6 Then GoTo Done
    nextIdentity.Low = -1: nextIdentity.High = -1
    If AllocateBowIdentity(value) Then GoTo Done
    If Not identitiesExhausted Then GoTo Done
    BowTestIdentityWrap = True
Done:
    nextIdentity = saved
    identitiesExhausted = wasExhausted
End Function
#End If

' Map/death invalidate identity now, but retain only the stream subscription.
' CheckUpdateNeededUser restarts the barrier after authoritative area creation.
Public Sub SuspendNpcBow(ByVal userIndex As Integer)
    Dim wanted As Boolean
    wanted = UserList(userIndex).NpcBow.WantsSubscription Or UserList(userIndex).NpcBow.Subscribed
    Call ClearNpcBow(userIndex, True)
    UserList(userIndex).NpcBow.WantsSubscription = wanted
End Sub

Public Sub ResumeNpcBowAfterArea(ByVal userIndex As Integer)
    If UserList(userIndex).NpcBow.WantsSubscription And Not UserList(userIndex).NpcBow.Subscribed Then
        If UserList(userIndex).flags.Muerto = 0 Then Call SubscribeNpcBow(userIndex)
    End If
End Sub

Private Function CurrentSession(ByVal userIndex As Integer) As Boolean
    With UserList(userIndex)
        CurrentSession = .NpcBow.Subscribed And .flags.UserLogged And .ConnectionDetails.ConnIDValida And _
            .NpcBow.OwnerId = .Id And .NpcBow.OwnerVersion = .VersionId And .NpcBow.Map = .pos.Map
    End With
End Function

Private Function VisibleNpc(ByVal userIndex As Integer, ByVal npcIndex As Integer) As Boolean
    If npcIndex < 1 Or npcIndex > MaxNPCs Then Exit Function
    With NpcList(npcIndex)
        If Not .flags.NPCActive Or .Char.charindex < 1 Or .Char.body = 0 Then Exit Function
        If .Stats.MinHp <= 0 Then Exit Function
        If .flags.invisible <> 0 Then Exit Function
        If .MaestroUser.ArrayIndex <> 0 Or .MaestroNPC.ArrayIndex <> 0 Then Exit Function
        If UserList(userIndex).flags.Muerto <> 0 Then Exit Function
        If Not InMapBounds(.pos.Map, .pos.x, .pos.y) Then Exit Function
        If MapData(.pos.x, .pos.y, .pos.Map).NpcIndex <> npcIndex Then Exit Function
    End With
    VisibleNpc = NpcBowInLoadedArea(userIndex, npcIndex)
End Function

Private Function WriteContext(ByVal userIndex As Integer, ByVal phase As Byte) As Boolean
    Dim writer As Network.writer, queued As Boolean
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.WriteInt16(ServerPacketID.eHooBowContext)
    Call writer.WriteInt32(UserList(userIndex).NpcBow.Context.Low)
    Call writer.WriteInt32(UserList(userIndex).NpcBow.Context.High)
    Call writer.WriteInt8(phase)
    Call SendData(ToIndex, userIndex, , False, queued)
    WriteContext = queued
End Function

Private Function WriteBinding(ByVal userIndex As Integer, ByVal kind As Byte, ByRef binding As t_NpcBowBinding) As Boolean
    Dim writer As Network.writer, queued As Boolean
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.WriteInt16(ServerPacketID.eHooBowBinding)
    Call writer.WriteInt32(UserList(userIndex).NpcBow.Context.Low)
    Call writer.WriteInt32(UserList(userIndex).NpcBow.Context.High)
    Call writer.WriteInt8(kind)
    Call writer.WriteInt16(binding.CharIndex)
    Call writer.WriteInt32(binding.Token.Low)
    Call writer.WriteInt32(binding.Token.High)
    Call SendData(ToIndex, userIndex, , False, queued)
    WriteBinding = queued
End Function

Public Sub WriteNpcBowResult(ByVal userIndex As Integer, ByRef reply As t_NpcBowReply)
    On Error GoTo Failed
    Dim writer As Network.writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.WriteInt16(ServerPacketID.eHooBowResult)
    Call writer.WriteInt32(reply.Request.Context.Low)
    Call writer.WriteInt32(reply.Request.Context.High)
    Call writer.WriteInt32(reply.Request.RequestId)
    Call writer.WriteInt8(CByte(reply.Code))
    Call writer.WriteInt32(reply.RetryAfter)
    Call SendData(ToIndex, userIndex)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "NpcBow.WriteResult", Erl)
End Sub

Private Sub RevokeForUser(ByVal userIndex As Integer, ByVal npcIndex As Integer)
    With UserList(userIndex).NpcBow
        If .BindingIndex Is Nothing Then Exit Sub
        If Not .BindingIndex.Exists(npcIndex) Then Exit Sub
        Dim slot As Integer, old As t_NpcBowBinding, emptyBinding As t_NpcBowBinding
        slot = .BindingIndex(npcIndex)
        old = .Bindings(slot)
        ' Retire first: a failed network send cannot keep authorization alive.
        Call .BindingIndex.Remove(npcIndex)
        .Bindings(slot) = emptyBinding
        If CurrentSession(userIndex) Then Call WriteBinding(userIndex, 2, old)
    End With
End Sub

Public Sub RevokeNpcBow(ByVal npcIndex As Integer)
    On Error GoTo Failed
    Dim key As Variant
    For Each key In subscribers.Keys
        Call RevokeForUser(CInt(key), npcIndex)
    Next key
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "NpcBow.RevokeNpc", Erl)
End Sub

Public Sub PruneNpcBowUser(ByVal userIndex As Integer)
    On Error GoTo Failed
    With UserList(userIndex).NpcBow
        If Not .Subscribed Then Exit Sub
        If Not CurrentSession(userIndex) Then
            Call ClearNpcBow(userIndex)
            Exit Sub
        End If
        Dim key As Variant
        For Each key In .BindingIndex.Keys
            If Not VisibleNpc(userIndex, CInt(key)) Then Call RevokeForUser(userIndex, CInt(key))
        Next key
    End With
    Exit Sub
Failed:
    Call ClearNpcBow(userIndex)
    Call TraceError(Err.Number, Err.Description, "NpcBow.PruneUser", Erl)
End Sub

' Called on every NPC movement, even when its area-cell identity is unchanged.
Public Sub PruneNpcBowActor(ByVal npcIndex As Integer)
    On Error GoTo Failed
    Dim key As Variant
    For Each key In subscribers.Keys
        If Not VisibleNpc(CInt(key), npcIndex) Then Call RevokeForUser(CInt(key), npcIndex)
    Next key
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "NpcBow.PruneActor", Erl)
End Sub

' Only MakeNPCChar calls this, after its owner-specific Create was queued.
Public Sub NpcBowCharacterCreated(ByVal userIndex As Integer, ByVal npcIndex As Integer, ByVal body As Integer)
    On Error GoTo Failed
    If Not CurrentSession(userIndex) Then Exit Sub
    If body = 0 Or Not VisibleNpc(userIndex, npcIndex) Then
        Call RevokeForUser(userIndex, npcIndex)
        Exit Sub
    End If
    Dim slot As Integer, i As Integer, binding As t_NpcBowBinding
    With UserList(userIndex).NpcBow
        If .BindingIndex.Exists(npcIndex) Then
            slot = .BindingIndex(npcIndex)
            binding = .Bindings(slot)
            If binding.CharIndex = NpcList(npcIndex).Char.charindex And binding.NpcVersion = NpcList(npcIndex).VersionId Then
                ' Repeated Create refreshes the same live actor, not its identity.
                If Not WriteBinding(userIndex, 1, binding) Then Call RevokeForUser(userIndex, npcIndex)
                Exit Sub
            End If
            Call RevokeForUser(userIndex, npcIndex)
        End If
        For i = 1 To .BindingCount
            If .Bindings(i).NpcIndex = 0 Then
                slot = i
                Exit For
            End If
        Next i
        If slot = 0 Then
            If .BindingCount >= MaxNPCs Then Exit Sub
            .BindingCount = .BindingCount + 1
            slot = .BindingCount
            ReDim Preserve .Bindings(1 To .BindingCount)
        End If
        If Not AllocateBowIdentity(binding.Token) Then
            Call ClearNpcBow(userIndex, True)
            Exit Sub
        End If
        binding.NpcIndex = npcIndex
        binding.NpcVersion = NpcList(npcIndex).VersionId
        binding.CharIndex = NpcList(npcIndex).Char.charindex
        If Not WriteBinding(userIndex, 1, binding) Then Exit Sub
        .Bindings(slot) = binding
        Call .BindingIndex.Add(npcIndex, slot)
    End With
    Exit Sub
Failed:
    Call ClearNpcBow(userIndex, True)
    Call TraceError(Err.Number, Err.Description, "NpcBow.CharacterCreated", Erl)
End Sub

Public Sub SubscribeNpcBow(ByVal userIndex As Integer)
    On Error GoTo Failed
    If Not UserList(userIndex).flags.UserLogged Or Not UserList(userIndex).ConnectionDetails.ConnIDValida Then Exit Sub
    If UserList(userIndex).flags.Muerto <> 0 Then
        Call ClearNpcBow(userIndex, True)
        UserList(userIndex).NpcBow.WantsSubscription = True
        Exit Sub
    End If
    If CurrentSession(userIndex) Then
        If UserList(userIndex).NpcBow.Ready Then Call WriteContext(userIndex, 1)
        Exit Sub
    End If
    Call ClearNpcBow(userIndex, True)
    With UserList(userIndex).NpcBow
        If Not AllocateBowIdentity(.Context) Then Exit Sub
        .OwnerId = UserList(userIndex).Id
        .OwnerVersion = UserList(userIndex).VersionId
        .Map = UserList(userIndex).pos.Map
        Set .BindingIndex = New Dictionary
        .Subscribed = True
        .WantsSubscription = True
        Call subscribers.Add(userIndex, True)
    End With
    If Not WriteContext(userIndex, 0) Then GoTo Abort
    Dim npcIndex As Integer, queued As Boolean
    For npcIndex = 1 To MaxNPCs
        If VisibleNpc(userIndex, npcIndex) Then
            ' A pre-subscription client actor may be an old reuse of this index.
            ' Remove/Create proves a fresh local generation before binding it.
            Call SendData(ToIndex, userIndex, PrepareMessageCharacterRemove(5, NpcList(npcIndex).Char.charindex, True), False, queued)
            If Not queued Then GoTo Abort
            With NpcList(npcIndex)
                Call MakeNPCChar(False, userIndex, npcIndex, .pos.Map, .pos.x, .pos.y)
            End With
            If Not UserList(userIndex).NpcBow.Subscribed Then Exit Sub
            If Not UserList(userIndex).NpcBow.BindingIndex.Exists(npcIndex) Then GoTo Abort
        End If
    Next npcIndex
    If Not WriteContext(userIndex, 1) Then GoTo Abort
    UserList(userIndex).NpcBow.Ready = True
    Exit Sub
Abort:
    Call ClearNpcBow(userIndex, True)
    Exit Sub
Failed:
    Call ClearNpcBow(userIndex, True)
    Call TraceError(Err.Number, Err.Description, "NpcBow.Subscribe", Erl)
End Sub

Private Sub MaxRemaining(ByRef remaining As Long, ByVal startTick As Long, ByVal interval As Long, ByVal nowTick As Long)
    Dim value As Double
    value = CDbl(interval) - TicksElapsed(startTick, nowTick)
    If value > remaining Then remaining = CLng(value)
End Sub

' Same four gates as the ordinary projectile path, without mutating any timer.
Public Function NpcBowRetryAfter(ByVal userIndex As Integer, ByVal nowTick As Long) As Long
    Dim result As Long, weapon As Integer
    With UserList(userIndex)
        Call MaxRemaining(result, .Counters.TimerLanzarSpell, IntervaloMagiaGolpe, nowTick)
        Call MaxRemaining(result, .Counters.TimerGolpeMagia, IntervaloGolpeMagia, nowTick)
        Call MaxRemaining(result, .Counters.TimerPuedeUsarArco, IntervaloFlechasCazadores, nowTick)
        weapon = .invent.EquippedWeaponObjIndex
        If weapon > 0 And weapon <= UBound(ObjData) Then
            Call MaxRemaining(result, .CdTimes(ObjData(weapon).cdType), ObjData(weapon).Cooldown, nowTick)
        End If
    End With
    NpcBowRetryAfter = result
End Function

Private Function ValidateEquipment(ByVal userIndex As Integer) As e_NpcBowResult
    ValidateEquipment = BowInvalidEquipment
    With UserList(userIndex).invent
        Dim slot As Integer, weapon As Integer, ammo As Integer
        slot = .EquippedWeaponSlot
        weapon = .EquippedWeaponObjIndex
        If slot < 1 Or slot > UserList(userIndex).CurrentInventorySlots Then Exit Function
        If weapon < 1 Or weapon > UBound(ObjData) Then Exit Function
        If .Object(slot).ObjIndex <> weapon Or .Object(slot).Equipped <> 1 Or .Object(slot).amount < 1 Then Exit Function
        If ObjData(weapon).Proyectil <> 1 Then Exit Function
        If ObjData(weapon).Municion > 0 Then
            ValidateEquipment = BowNoAmmo
            slot = .EquippedMunitionSlot
            ammo = .EquippedMunitionObjIndex
            If slot < 1 Or slot > UserList(userIndex).CurrentInventorySlots Then Exit Function
            If ammo < 1 Or ammo > UBound(ObjData) Then Exit Function
            If .Object(slot).ObjIndex <> ammo Or .Object(slot).amount < 1 Or .Object(slot).Equipped <> 1 Then Exit Function
            If ObjData(ammo).OBJType <> otArrows Or ObjData(ammo).Subtipo <> ObjData(weapon).Municion Then Exit Function
        End If
    End With
    ValidateEquipment = BowAccepted
End Function

Private Function SameRequest(ByRef a As t_NpcBowRequest, ByRef b As t_NpcBowRequest) As Boolean
    SameRequest = SameIdentity(a.Context, b.Context) And a.RequestId = b.RequestId And _
                  a.CharIndex = b.CharIndex And SameIdentity(a.Token, b.Token)
End Function

' Invoked only after UsuarioAtacaNpc's actual permission/guild checks pass.
' Legacy callers already paid through WorkLeftClick and do not use this path.
Public Function CommitNpcBowCosts(ByVal userIndex As Integer) As Boolean
    With UserList(userIndex)
        If .Stats.MinSta < 10 Then Exit Function
        If NpcBowRetryAfter(userIndex, GetTickCountRaw()) > 0 Then Exit Function
        If Not IntervaloPermiteUsarArcos(userIndex) Then Exit Function
        If .flags.Meditando Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(ToPCAliveArea, userIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
        End If
        Call CancelExit(userIndex)
        Call QuitarSta(userIndex, RandomNumber(1, 10))
    End With
    CommitNpcBowCosts = True
End Function

Public Sub ProcessNpcBowAttack(ByVal userIndex As Integer, ByRef request As t_NpcBowRequest, ByVal nowTick As Long, ByRef reply As t_NpcBowReply)
    On Error GoTo Failed
    Dim emptyReply As t_NpcBowReply, i As Integer, cacheSlot As Integer, cached As Boolean
    Dim npcIndex As Integer, binding As t_NpcBowBinding, interaction As t_AttackInteractionResult
    Dim weapon As Integer, projectile As Integer, committed As Boolean, completed As Boolean
    reply = emptyReply
    reply.Request = request
    reply.Code = BowInvalidContext
    If Not CurrentSession(userIndex) Then Exit Sub
    With UserList(userIndex).NpcBow
        If Not .Ready Or Not SameIdentity(.Context, request.Context) Then Exit Sub
        reply.Code = BowInvalidTarget
        If request.RequestId <= 0 Then Exit Sub
        For i = 0 To 15
            If .Replies(i).Request.RequestId = request.RequestId Then
                If SameRequest(request, .Replies(i).Request) Then reply = .Replies(i)
                Exit Sub
            End If
        Next i
        If request.RequestId <= .LastRequest Then Exit Sub
        .LastRequest = request.RequestId
        cacheSlot = .NextReply
        .NextReply = (.NextReply + 1) Mod 16
        ' Cache a terminal uncertain result before any combat side effect.
        reply.Code = BowInternalError
        .Replies(cacheSlot) = reply
        cached = True
        reply.Code = BowInvalidTarget
        If request.CharIndex < 1 Or request.CharIndex > UBound(CharList) Then GoTo Done
        npcIndex = CharList(request.CharIndex)
        If npcIndex < 1 Or npcIndex > MaxNPCs Then GoTo Done
        If Not .BindingIndex.Exists(npcIndex) Then GoTo Done
        binding = .Bindings(.BindingIndex(npcIndex))
        If binding.CharIndex <> request.CharIndex Or Not SameIdentity(binding.Token, request.Token) Then GoTo Done
        If NpcList(npcIndex).Char.charindex <> request.CharIndex Then
            Call RevokeForUser(userIndex, npcIndex)
            GoTo Done
        End If
        If binding.NpcVersion <> NpcList(npcIndex).VersionId Or Not VisibleNpc(userIndex, npcIndex) Then
            Call RevokeForUser(userIndex, npcIndex)
            GoTo Done
        End If
    End With
    With UserList(userIndex)
        reply.Code = BowActorBlocked
        If .flags.Muerto <> 0 Or .flags.Descansar Or .flags.YaGuardo Or .flags.Comerciando Then GoTo Done
        If UserMod.IsStun(.flags, .Counters) Then GoTo Done
        reply.Code = BowOutOfRange
        If Abs(.pos.x - NpcList(npcIndex).pos.x) > RANGO_VISION_X Or Abs(.pos.y - NpcList(npcIndex).pos.y) > RANGO_VISION_Y Then GoTo Done
        reply.Code = ValidateEquipment(userIndex)
        If reply.Code <> BowAccepted Then GoTo Done
        reply.Code = BowForbiddenTarget
        If NpcList(npcIndex).OnlyForGuilds = 1 And .GuildIndex <= 0 Then GoTo Done
        interaction = UserCanAttackNpc(userIndex, npcIndex)
        If Not interaction.CanAttack Then
            Select Case interaction.Result
                Case eMounted, eDeathAttacker, eTalkWithMaster, eCreatureInmunity
                    reply.Code = BowActorBlocked
            End Select
            GoTo Done
        End If
        reply.Code = BowNoEnergy
        If .Stats.MinSta < 10 Then GoTo Done
        reply.RetryAfter = NpcBowRetryAfter(userIndex, nowTick)
        If reply.RetryAfter > 0 Then
            reply.Code = BowRateLimited
            GoTo Done
        End If
        weapon = .invent.EquippedWeaponObjIndex
        projectile = GetProjectileView(UserList(userIndex))
        reply.Code = BowInternalError
        If Not ResolveNpcBowShot(userIndex, npcIndex, CByte(projectile), committed, True) Then GoTo Done
        Call FinishProjectileAttempt(userIndex, ObjData(weapon), True, completed, True)
        If Not completed Then GoTo Done
        reply.Code = BowAccepted
        reply.RetryAfter = NpcBowRetryAfter(userIndex, GetTickCountRaw())
    End With
Done:
    reply.Committed = committed
    If cached And CurrentSession(userIndex) Then
        If SameIdentity(UserList(userIndex).NpcBow.Context, request.Context) Then UserList(userIndex).NpcBow.Replies(cacheSlot) = reply
    End If
    Exit Sub
Failed:
    reply.Code = BowInternalError
    reply.RetryAfter = 0
    reply.Committed = committed
    If cached And CurrentSession(userIndex) Then
        If SameIdentity(UserList(userIndex).NpcBow.Context, request.Context) Then UserList(userIndex).NpcBow.Replies(cacheSlot) = reply
    End If
    Call TraceError(Err.Number, Err.Description, "NpcBow.Attack", Erl)
End Sub

Public Sub HandleNpcBowSubscribe(ByVal userIndex As Integer, ByVal message As Network.reader)
    If Not AllowNpcBowMessage(userIndex, GetTickCountRaw()) Then Exit Sub
    If message.GetAvailable() = 0 Then Call SubscribeNpcBow(userIndex)
End Sub

' Transport abuse budget is separate from gameplay cooldowns and replay cache.
' Over-budget packets are ignored, never converted into a different cached
' outcome. Once admitted again, an exact duplicate still gets its old result.
Public Function AllowNpcBowMessage(ByVal userIndex As Integer, ByVal nowTick As Long) As Boolean
    With UserList(userIndex).NpcBow
        If TicksElapsed(.WindowStart, nowTick) >= 1000 Then
            .WindowStart = nowTick
            .WindowCount = 0
        End If
        If .WindowCount >= REQUESTS_PER_SECOND Then Exit Function
        .WindowCount = .WindowCount + 1
    End With
    AllowNpcBowMessage = True
End Function

#If DEBUGGING = 1 Then
' Reuse the admitted-message window: at most four native debug log entries per
' user/second, including replayed requests. No trace work in production builds.
Public Function NpcBowDiagnostic(ByVal userIndex As Integer, ByRef reply As t_NpcBowReply) As String
    With UserList(userIndex).NpcBow
        If .WindowCount < 1 Or .WindowCount > 4 Then Exit Function
    End With
    With reply
        NpcBowDiagnostic = "NPC bow user=" & CStr(userIndex) & _
            " context=" & Hex$(.Request.Context.High) & ":" & Hex$(.Request.Context.Low) & _
            " request=" & CStr(.Request.RequestId) & " target=" & CStr(.Request.CharIndex) & _
            " token=" & Hex$(.Request.Token.High) & ":" & Hex$(.Request.Token.Low) & _
            " result=" & CStr(.Code) & " retry_ms=" & CStr(.RetryAfter) & _
            " committed=" & IIf(.Committed, "1", "0")
    End With
End Function
#End If

Public Sub HandleNpcBowAttack(ByVal userIndex As Integer, ByVal message As Network.reader)
    On Error GoTo Failed
    If Not AllowNpcBowMessage(userIndex, GetTickCountRaw()) Then Exit Sub
    If message.GetAvailable() <> 22 Then Exit Sub
    Dim request As t_NpcBowRequest, reply As t_NpcBowReply
    request.Context.Low = message.ReadInt32()
    request.Context.High = message.ReadInt32()
    request.RequestId = message.ReadInt32()
    request.CharIndex = message.ReadInt16()
    request.Token.Low = message.ReadInt32()
    request.Token.High = message.ReadInt32()
    If request.RequestId <= 0 Then Exit Sub
    Call ProcessNpcBowAttack(userIndex, request, GetTickCountRaw(), reply)
    #If DEBUGGING = 1 Then
        Dim diagnostic As String
        diagnostic = NpcBowDiagnostic(userIndex, reply)
        If LenB(diagnostic) > 0 Then Call LogInfoServidor(diagnostic)
    #End If
    Call WriteNpcBowResult(userIndex, reply)
    Exit Sub
Failed:
    Call TraceError(Err.Number, Err.Description, "NpcBow.HandleAttack", Erl)
End Sub
