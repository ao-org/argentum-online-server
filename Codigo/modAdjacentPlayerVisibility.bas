Attribute VB_Name = "modAdjacentPlayerVisibility"
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit

Public Const ADJACENT_PLAYER_LIMIT As Long = 128
Private Const RADIUS_X As Long = 24
Private Const RADIUS_Y As Long = 18
Private Const UPDATE_MS As Long = 100
Private Const LOG_MS As Long = 10000
Private Const MAX_LONG As Long = &H7FFFFFFF
Private Const MIN_LONG As Long = &H80000000
Private Const FLAG_NAV As Long = 2
Private Const FLAG_AURAS As Long = 32
Private Const FLAG_SNAP As Long = 64
Private Const FLAG_CLAN As Long = 128
Private Const HANDOFF_REPLACEMENT As Long = 7
Private Enum e_HandoffStage
    HANDOFF_FRESH = 0
    HANDOFF_BLANK = 1
    HANDOFF_VISIBILITY = 2
    HANDOFF_FULL = 3
    HANDOFF_CANCEL = 4
    HANDOFF_RESTART = 5
End Enum
Private mLastUpdate As Long
Private mLastLog As Long
Private mLogged As Boolean
Private mGeneration As Long
Private mLow As Long
Private mHigh As Long
Private mExhausted As Boolean
Private mIdentities As Scripting.Dictionary
Private mKnown As Scripting.Dictionary
Private mViewers As Scripting.Dictionary
Private mEpochs As Scripting.Dictionary
Private mPending As Scripting.Dictionary
Private mActive As Scripting.Dictionary

Private Sub EnsureState()
    If Not mKnown Is Nothing Then Exit Sub
    Set mKnown = New Scripting.Dictionary
    Set mViewers = New Scripting.Dictionary
    Set mEpochs = New Scripting.Dictionary
    Set mIdentities = New Scripting.Dictionary
    Set mPending = New Scripting.Dictionary
    Set mActive = New Scripting.Dictionary
End Sub

Private Function LiveUser(ByVal slot As Integer) As Boolean
    If slot < 1 Or slot > UBound(UserList) Then Exit Function
    LiveUser = UserList(slot).flags.UserLogged And UserList(slot).ConnectionDetails.ConnIDValida
End Function

Private Function ReadyUser(ByVal slot As Integer) As Boolean
    If Not LiveUser(slot) Then Exit Function
    With UserList(slot)
        If .Char.charindex < 1 Or .Char.charindex > UBound(CharList) Then Exit Function
        If CharList(.Char.charindex) <> slot Then Exit Function
        If Not InMapBounds(.pos.Map, .pos.x, .pos.y) Then Exit Function
    End With
    ReadyUser = True
End Function

Private Function Identity(ByVal slot As Integer) As Variant
    Call EnsureState()
    Dim key As String, signature As String, value As Variant
    key = CStr(slot)
    signature = CStr(UserList(slot).VersionId) & ":" & CStr(UserList(slot).ConnectionDetails.ConnID)
    If mIdentities.Exists(key) Then
        value = mIdentities(key)
        If value(2) = signature Then
            Identity = value
            Exit Function
        End If
        Call ResetAdjacentPlayerSlot(slot)
    End If
    If mExhausted Then Exit Function
    ' Increment unsigned 64 bits without signed overflow or floating-point loss.
    If mLow = -1 Then
        If mHigh = -1 Then
            mExhausted = True
            Call ReportFailure("session ID exhausted")
            Exit Function
        End If
        mLow = 0
        If mHigh = MAX_LONG Then mHigh = MIN_LONG Else mHigh = mHigh + 1
    ElseIf mLow = MAX_LONG Then
        mLow = MIN_LONG
    Else
        mLow = mLow + 1
    End If
    value = Array(mLow, mHigh, signature, 0&)
    Call mIdentities.Add(key, value)
    Identity = value
End Function

Private Function SameIdentity(ByVal first As Variant, ByVal second As Variant) As Boolean
    If IsEmpty(first) Or IsEmpty(second) Then Exit Function
    SameIdentity = first(0) = second(0) And first(1) = second(1)
End Function

Public Sub RelocateAdjacentPlayer(ByVal subject As Integer)
    #If PYMMO = 0 Then
        Exit Sub
    #End If
    On Error GoTo Failed
    If mIdentities Is Nothing Then Exit Sub
    If Not mIdentities.Exists(CStr(subject)) Then Exit Sub
    Dim value As Variant
    value = mIdentities(CStr(subject))
    If value(3) = MAX_LONG Or value(3) = -1 Then
        value(3) = -1&
    Else
        value(3) = CLng(value(3)) + 1
    End If
    mIdentities(CStr(subject)) = value
    Exit Sub
Failed:
    Call ReportFailure("relocation", Err.Number, Err.Description)
End Sub

Private Function ViewerSignature(ByVal slot As Integer) As String
    Dim value As Variant
    value = Identity(slot)
    If IsEmpty(value) Then Exit Function
    ViewerSignature = CStr(value(0)) & ":" & CStr(value(1)) & ":" & CStr(UserList(slot).pos.Map)
End Function

Public Sub ResetAdjacentPlayerSlot(ByVal slot As Integer)
    On Error GoTo Failed
    Call InvalidateAdjacentPlayerViewer(slot)
    If Not mIdentities Is Nothing Then
        If mIdentities.Exists(CStr(slot)) Then Call mIdentities.Remove(CStr(slot))
    End If
    Exit Sub
Failed:
    Call ReportFailure("reset slot", Err.Number, Err.Description)
End Sub

Public Sub InvalidateAdjacentPlayerViewer(ByVal viewer As Integer)
    On Error GoTo Failed
    If mKnown Is Nothing Then Exit Sub
    Dim key As String
    key = CStr(viewer)
    If mKnown.Exists(key) Then Call mKnown.Remove(key)
    If mViewers.Exists(key) Then Call mViewers.Remove(key)
    If mEpochs.Exists(key) Then Call mEpochs.Remove(key)
    If mPending.Exists(key) Then Call mPending.Remove(key)
    If mActive.Exists(key) Then Call mActive.Remove(key)
    Exit Sub
Failed:
    Call ReportFailure("invalidate viewer", Err.Number, Err.Description)
End Sub

Public Sub UpdateAdjacentPlayerVisibility(ByVal currentTime As Long)
    #If PYMMO = 0 Then
        Exit Sub
    #End If
    On Error GoTo Failed
    If mLastUpdate <> 0 Then
        If TicksElapsed(mLastUpdate, currentTime) < UPDATE_MS Then Exit Sub
    End If
    mLastUpdate = currentTime
    Call EnsureState()
    Dim key As Variant, viewer As Integer
    For Each key In mIdentities.Keys
        If Not LiveUser(CInt(key)) Or CLng(key) > LastUser Then Call ResetAdjacentPlayerSlot(CInt(key))
    Next key
    For viewer = 1 To LastUser
        If LiveUser(viewer) Then Call RefreshViewer(viewer)
    Next viewer
    Exit Sub
Failed:
    Call ReportFailure("update", Err.Number, Err.Description)
End Sub

Private Sub RefreshViewer(ByVal viewer As Integer)
    On Error GoTo Failed
    If Not ReadyUser(viewer) Then Exit Sub
    Dim key As String, signature As String, epoch As Long
    Dim previous As Scripting.Dictionary, current As New Scripting.Dictionary
    Dim actor As Variant, oldValue As Variant, value As Variant
    key = CStr(viewer)
    signature = ViewerSignature(viewer)
    If Len(signature) = 0 Then Exit Sub
    If mViewers.Exists(key) Then
        If mViewers(key) <> signature Then Call InvalidateAdjacentPlayerViewer(viewer)
    End If
    If Not mViewers.Exists(key) Then Call mViewers.Add(key, signature)
    Call RetryHandoffs(viewer)
    If Not CollectVisible(viewer, current) Then Exit Sub
    If Not mKnown.Exists(key) Then
        If mGeneration = MAX_LONG Then
            Call ReportFailure("scene epoch exhausted")
            Exit Sub
        End If
        mGeneration = mGeneration + 1
        If Not SendReset(viewer, mGeneration) Then Exit Sub
        Set previous = New Scripting.Dictionary
        Call mKnown.Add(key, previous)
        Call mEpochs.Add(key, mGeneration)
    End If
    Set previous = mKnown(key)
    epoch = CLng(mEpochs(key))
    For Each actor In previous.Keys
        oldValue = previous(actor)
        If current.Exists(actor) Then
            value = current(actor)
            If Not SameIdentity(oldValue, value) Then
                If Not SendRemove(viewer, epoch, CInt(actor), oldValue) Then Exit Sub
                Call previous.Remove(actor)
            End If
        Else
            If Not SendRemove(viewer, epoch, CInt(actor), oldValue) Then Exit Sub
            Call previous.Remove(actor)
        End If
    Next actor
    For Each actor In current.Keys
        value = current(actor)
        If Not RetireActive(viewer, CInt(actor)) Then Exit Sub
        If previous.Exists(actor) Then
            oldValue = previous(actor)
            If Fingerprint(oldValue) <> Fingerprint(value) Then
                If Not SendUpsert(viewer, epoch, CInt(actor), value, NeedsSnap(oldValue, value)) Then Exit Sub
                previous(actor) = value
            End If
        Else
            If Not SendUpsert(viewer, epoch, CInt(actor), value, True) Then Exit Sub
            Call previous.Add(actor, value)
        End If
    Next actor
    Exit Sub
Failed:
    ' Retain accepted state on failure; the next sample can retry the missing suffix.
    Call ReportFailure("refresh", Err.Number, Err.Description)
End Sub

Private Function CollectVisible(ByVal viewer As Integer, ByVal result As Scripting.Dictionary) As Boolean
    On Error GoTo Failed
    Dim map As Integer, placements() As t_AdjacentNpcPlacement, count As Long
    Dim i As Long, j As Long, x As Long, y As Long, x0 As Long, x1 As Long, y0 As Long, y1 As Long
    Dim neighbor As Integer, candidates As New Scripting.Dictionary, slot As Integer
    Dim keys As Variant, swap As Variant, value As Variant, available As Long
    map = UserList(viewer).pos.Map
    CollectVisible = True
    available = ADJACENT_PLAYER_LIMIT
    If mPending.Exists(CStr(viewer)) Then available = available - mPending(CStr(viewer)).Count
    If available <= 0 Then Exit Function
    If map < 1 Or map > NumMaps Then Exit Function
    If MapInfo(map).MapResource <> map Then Exit Function
    If Not TryGetAdjacentVisibilityPlacements(map, placements, count) Then
        CollectVisible = False
        Exit Function
    End If
    For i = 1 To count
        neighbor = placements(i).MapId
        If neighbor > 0 And neighbor <= NumMaps And neighbor <> map Then
            If MapInfo(neighbor).MapResource = neighbor Then
                x0 = max(XMinMapSize, CLng(UserList(viewer).pos.x) - placements(i).dx - RADIUS_X)
                x1 = min(XMaxMapSize, CLng(UserList(viewer).pos.x) - placements(i).dx + RADIUS_X)
                y0 = max(YMinMapSize, CLng(UserList(viewer).pos.y) - placements(i).dy - RADIUS_Y)
                y1 = min(YMaxMapSize, CLng(UserList(viewer).pos.y) - placements(i).dy + RADIUS_Y)
                If x0 <= x1 And y0 <= y1 Then
                    For y = y0 To y1
                        For x = x0 To x1
                            slot = MapData(x, y, neighbor).UserIndex
                            Call AddCandidate(candidates, slot, neighbor, x, x, y, y)
                        Next x
                    Next y
                    ' A tile stores only one occupant. Map membership retains overlapping
                    ' ghosts/staff; do not scan the global player table for every viewer.
                    For j = 1 To ConnGroups(neighbor).CountEntrys
                        slot = ConnGroups(neighbor).UserEntrys(j)
                        Call AddCandidate(candidates, slot, neighbor, x0, x1, y0, y1)
                    Next j
                End If
            End If
        End If
    Next i
    If candidates.Count = 0 Then Exit Function
    keys = candidates.Keys
    For i = 1 To UBound(keys)
        swap = keys(i): j = i - 1
        Do While j >= 0
            If CLng(keys(j)) <= CLng(swap) Then Exit Do
            keys(j + 1) = keys(j): j = j - 1
        Loop
        keys(j + 1) = swap
    Next i
    For i = 0 To UBound(keys)
        slot = CInt(keys(i))
        If BuildSnapshot(viewer, slot, value) Then
            Call result.Add(CStr(slot), value)
            If result.Count >= available Then Exit Function
        End If
    Next i
    Exit Function
Failed:
    CollectVisible = False
    Call ReportFailure("collect", Err.Number, Err.Description)
End Function

Private Sub AddCandidate(ByVal candidates As Scripting.Dictionary, ByVal slot As Integer, ByVal map As Integer, _
                         ByVal x0 As Long, ByVal x1 As Long, ByVal y0 As Long, ByVal y1 As Long)
    If Not ReadyUser(slot) Then Exit Sub
    With UserList(slot)
        If .Char.charindex <= 0 Or .Char.body <= 0 Or .pos.Map <> map Then Exit Sub
        If .pos.x < x0 Or .pos.x > x1 Or .pos.y < y0 Or .pos.y > y1 Then Exit Sub
    End With
    If Not candidates.Exists(CStr(slot)) Then Call candidates.Add(CStr(slot), True)
End Sub

Private Function Admitted(ByVal viewer As Integer, ByVal subject As Integer) As Boolean
    ' NotifyUser's admission is deliberately stricter than transient event routing.
    With UserList(subject)
        If .flags.AdminInvisible = 1 Then
            Admitted = CanUserSeeAdminInvisible(viewer, subject)
            Exit Function
        End If
        If UserList(viewer).flags.Muerto = 1 And MapInfo(UserList(viewer).pos.Map).Seguro = 0 Then
            If Not IsSet(.flags.StatusMask, e_StatusMask.eTalkToDead) Then
                If .GuildIndex = 0 Or UserList(viewer).GuildIndex <> .GuildIndex Then Exit Function
                If modGuilds.NivelDeClan(.GuildIndex) < RequiredGuildLevelSeeInvisible Then Exit Function
            End If
        End If
    End With
    Admitted = True
End Function

Private Function Representation(ByVal viewer As Integer, ByVal subject As Integer, ByRef mode As Byte, ByRef flags As Long) As Boolean
    If Not Admitted(viewer, subject) Then Exit Function
    Dim guildVisible As Boolean, clanLevel As Byte, sameClan As Boolean
    With UserList(subject)
        sameClan = .GuildIndex > 0 And .GuildIndex = UserList(viewer).GuildIndex
        ' MakeUserChar only sends a clan level for undisguised, named players.
        If .showName And .flags.Mimetizado = e_EstadoMimetismo.Desactivado Then
            If .GuildIndex > 0 Then clanLevel = modGuilds.NivelDeClan(.GuildIndex)
            If sameClan And Not (.clase = e_Class.Pirat And .flags.Oculto = 1) Then flags = flags Or FLAG_CLAN
        End If
        guildVisible = sameClan And clanLevel >= RequiredGuildLevelSeeInvisible And _
            .Char.body <> iCuerpoMuerto And .Char.body <> 830 And .Char.body <> 87
        If .flags.Navegando <> 0 Then flags = flags Or FLAG_NAV
        ' Mount/swim/rubber are derived by HOO from the displayed body, like
        ' CharacterCreate. Actual gameplay flags can disclose a mimic's identity.
        If .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminInvisible = 0 Then
            mode = 0: flags = flags Or FLAG_AURAS
        ElseIf UserList(viewer).flags.Privilegios > 1 And .flags.Privilegios <= UserList(viewer).flags.Privilegios Then
            mode = 1
            If guildVisible Or .flags.Navegando <> 0 Then flags = flags Or FLAG_AURAS
        ElseIf guildVisible Then
            mode = 2: flags = flags Or FLAG_AURAS
        ElseIf .flags.Navegando <> 0 Then
            mode = 3: flags = flags Or FLAG_AURAS
        Else
            Exit Function
        End If
    End With
    Representation = True
End Function

Private Function AuthorizedName(ByVal subject As Integer) As String
    With UserList(subject)
        If Not .showName Then Exit Function
        ' RefreshCharStatus explicitly hides a concealed pirate's label.
        If .clase = e_Class.Pirat And .flags.Oculto = 1 Then Exit Function
        If .flags.Mimetizado <> e_EstadoMimetismo.Desactivado Then
            AuthorizedName = .NameMimetizado
        Else
            AuthorizedName = GetUserDisplayNameOrReal(subject)
            If .GuildIndex = 0 And .flags.EnConsulta Then AuthorizedName = AuthorizedName & " [CONSULTA]"
            AuthorizedName = AuthorizedName & " {" & GetCharacterAlias(subject) & "}"
            If .GuildIndex > 0 Then AuthorizedName = AuthorizedName & " <" & modGuilds.GuildName(.GuildIndex) & ">"
        End If
    End With
End Function

Private Function BuildSnapshot(ByVal viewer As Integer, ByVal subject As Integer, ByRef value As Variant) As Boolean
    Dim mode As Byte, flags As Long, session As Variant
    If Not Representation(viewer, subject, mode, flags) Then Exit Function
    session = Identity(subject)
    If IsEmpty(session) Then Exit Function
    If session(3) = -1 Then Exit Function
    With UserList(subject)
        value = Array(session(0), session(1), .pos.Map, .pos.x, .pos.y, .Char.body, .Char.head, .Char.Heading, _
            .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim, .Char.speeding, _
            mode, flags, .Faccion.Status, .flags.Privilegios, .Stats.tipoUsuario, .flags.CurrentTeam, .flags.tiene_bandera, _
            AuthorizedName(subject), .Char.Arma_Aura, .Char.Body_Aura, .Char.Escudo_Aura, .Char.Head_Aura, .Char.Otra_Aura, .Char.DM_Aura, .Char.RM_Aura, session(3))
    End With
    BuildSnapshot = True
End Function

Private Function Fingerprint(ByVal value As Variant) As String
    Dim i As Long, field As String
    For i = 0 To UBound(value)
        If i = 13 Then field = CStr(CDbl(value(i))) Else field = CStr(value(i))
        Fingerprint = Fingerprint & Len(field) & ":" & field
    Next i
End Function

Private Function NeedsSnap(ByVal previous As Variant, ByVal current As Variant) As Boolean
    NeedsSnap = previous(2) <> current(2) Or previous(29) <> current(29) Or _
        Abs(CLng(previous(3)) - current(3)) + Abs(CLng(previous(4)) - current(4)) > 1
End Function

Private Function SendReset(ByVal viewer As Integer, ByVal epoch As Long) As Boolean
    On Error GoTo Failed
    Dim writer As Network.writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.WriteInt16(ServerPacketID.eAdjacentPlayerReset)
    Call writer.WriteInt32(epoch)
    Call writer.WriteInt16(UserList(viewer).pos.Map)
    Call modSendData.SendData(ToIndex, viewer, , False, SendReset, True)
    If Not SendReset Then Call ReportFailure("reset enqueue")
    Exit Function
Failed:
    Call ReportFailure("reset writer", Err.Number, Err.Description)
    If Not writer Is Nothing Then Call writer.Clear()
End Function

Private Function SendRemove(ByVal viewer As Integer, ByVal epoch As Long, ByVal subject As Integer, ByVal value As Variant) As Boolean
    On Error GoTo Failed
    Dim writer As Network.writer
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.WriteInt16(ServerPacketID.eAdjacentPlayerRemove)
    Call writer.WriteInt32(epoch)
    Call writer.WriteInt16(subject)
    Call writer.WriteInt32(CLng(value(0)))
    Call writer.WriteInt32(CLng(value(1)))
    Call modSendData.SendData(ToIndex, viewer, , False, SendRemove, True)
    If Not SendRemove Then Call ReportFailure("remove enqueue")
    Exit Function
Failed:
    Call ReportFailure("remove writer", Err.Number, Err.Description)
    If Not writer Is Nothing Then Call writer.Clear()
End Function

Private Sub WriteBoundedString(ByVal writer As Network.writer, ByVal value As String, ByVal limit As Long)
    Dim offset As Long, size As Long, prefix As Long
    ' Checking UTF-16 length first also bounds temporary writer allocation.
    If Len(value) > limit Then Err.Raise 5, , "Adjacent player text exceeds limit"
    offset = writer.GetOffset()
    Call writer.WriteString8(value)
    size = writer.GetOffset() - offset
    prefix = 1
    If size >= 130 Then prefix = 2
    If size - prefix > limit Then Err.Raise 5, , "Adjacent player UTF-8 text exceeds limit"
End Sub

Private Function SendUpsert(ByVal viewer As Integer, ByVal epoch As Long, ByVal subject As Integer, ByVal value As Variant, ByVal snap As Boolean) As Boolean
    On Error GoTo Failed
    Dim writer As Network.writer, i As Long, flags As Long
    Set writer = Protocol_Writes.GetWriterBuffer()
    Call writer.WriteInt16(ServerPacketID.eAdjacentPlayerUpsert)
    Call writer.WriteInt32(epoch)
    Call writer.WriteInt16(subject)
    Call writer.WriteInt32(CLng(value(0)))
    Call writer.WriteInt32(CLng(value(1)))
    Call writer.WriteInt16(CInt(value(2)))
    Call writer.WriteInt8(CByte(value(3)))
    Call writer.WriteInt8(CByte(value(4)))
    Call writer.WriteInt16(CInt(value(5)))
    Call writer.WriteInt16(CInt(value(6)))
    Call writer.WriteInt8(CByte(value(7)))
    For i = 8 To 12
        Call writer.WriteInt16(CInt(value(i)))
    Next i
    Call writer.WriteReal32(CSng(value(13)))
    Call writer.WriteInt8(CByte(value(14)))
    flags = CLng(value(15))
    If snap Then flags = flags Or FLAG_SNAP
    Call writer.WriteInt16(CInt(flags))
    For i = 16 To 20
        Call writer.WriteInt8(CByte(value(i)))
    Next i
    Call WriteBoundedString(writer, CStr(value(21)), 4096)
    For i = 22 To 28
        Call WriteBoundedString(writer, CStr(value(i)), 1024)
    Next i
    Call modSendData.SendData(ToIndex, viewer, , False, SendUpsert, True)
    If Not SendUpsert Then Call ReportFailure("upsert enqueue")
    Exit Function
Failed:
    Call ReportFailure("upsert writer", Err.Number, Err.Description)
    If Not writer Is Nothing Then Call writer.Clear()
End Function

Public Function PrepareAdjacentPlayerHandoff(ByVal viewer As Integer, ByVal subject As Integer, Optional ByRef visibilityHandled As Boolean = False) As Boolean
    visibilityHandled = False
    PrepareAdjacentPlayerHandoff = True
    #If PYMMO = 0 Then
        Exit Function
    #End If
    On Error GoTo Failed
    Call EnsureState()
    Dim known As Scripting.Dictionary, pending As Scripting.Dictionary, key As String, actor As String
    Dim value As Variant, hidden As Boolean, queued As Boolean
    key = CStr(viewer)
    actor = CStr(subject)
    If mPending.Exists(key) Then
        Set pending = mPending(key)
        If pending.Exists(actor) Then
            value = pending(actor)
            If value(5) = HANDOFF_CANCEL Or Not HandoffIsCurrent(viewer, subject, value) Then
                visibilityHandled = True
                PrepareAdjacentPlayerHandoff = False
                ' Only this explicit create request may replace obsolete work.
                ' Keep the old accepted owner intact until its removal succeeds.
                value(HANDOFF_REPLACEMENT) = NewHandoffRequest(viewer, subject)
                If Not CancelHandoff(viewer, subject, pending, value) Then Exit Function
                If Not pending.Exists(actor) Then Exit Function
            End If
        End If
    End If
    If mKnown.Exists(key) Then
        Set known = mKnown(key)
        If known.Exists(actor) Then
            ' Reserve before removing the adjacent copy, including on first failure.
            visibilityHandled = True
            PrepareAdjacentPlayerHandoff = False
            If Not DeferCreate(viewer, subject) Then Exit Function
            If Not SendRemove(viewer, CLng(mEpochs(key)), subject, known(actor)) Then Exit Function
            Call known.Remove(actor)
        End If
    End If
    If Not mPending.Exists(key) Then Exit Function
    Set pending = mPending(key)
    If Not pending.Exists(actor) Then Exit Function
    visibilityHandled = True
    PrepareAdjacentPlayerHandoff = False
    value = pending(actor)
    If Not HandoffIsCurrent(viewer, subject, value) Then
        Call CancelHandoff(viewer, subject, pending, value)
        Exit Function
    End If
    If value(5) = HANDOFF_RESTART Then
        If Not RemoveHandoffRecord(viewer, subject, value) Then Exit Function
        value(5) = HANDOFF_BLANK: value(6) = -1&
        pending(actor) = value
    End If
    If Not Admitted(viewer, subject) Then
        ' A legacy admin/death removal may have erased an accepted blank record.
        ' Retire any record we still own before restarting; denied subjects must
        ' not leave a collision-only ghost or lose their bounded retry entry.
        value(5) = HANDOFF_RESTART: value(6) = -1&
        pending(actor) = value
        If RemoveHandoffRecord(viewer, subject, value) Then
            value(5) = HANDOFF_BLANK
            pending(actor) = value
        End If
        Exit Function
    End If
    With UserList(subject)
        hidden = .flags.invisible <> 0 Or .flags.Oculto <> 0
    End With
    If value(5) = HANDOFF_FRESH And hidden Then value(5) = HANDOFF_BLANK
    If value(5) = HANDOFF_BLANK Then
        pending(actor) = value
        If Not SendHandoffBlank(viewer, subject) Then Exit Function
        value(5) = HANDOFF_VISIBILITY
        pending(actor) = value
    End If
    If value(5) = HANDOFF_VISIBILITY Or (value(5) = HANDOFF_FULL And value(6) <> Abs(CLng(hidden))) Then
        With UserList(subject)
            Call modSendData.SendData(ToIndex, viewer, PrepareMessageSetInvisible(.Char.charindex, hidden, .pos.x, .pos.y), False, queued, True)
        End With
        If Not queued Then
            Call ReportFailure("handoff visibility enqueue")
            Exit Function
        End If
        value(5) = HANDOFF_FULL: value(6) = Abs(CLng(hidden))
        pending(actor) = value
    End If
    ' MakeUserChar sends the full snapshot only after the required state is accepted.
    PrepareAdjacentPlayerHandoff = True
    Exit Function
Failed:
    PrepareAdjacentPlayerHandoff = False
    Call ReportFailure("handoff", Err.Number, Err.Description)
End Function

Private Function DeferCreate(ByVal viewer As Integer, ByVal subject As Integer) As Boolean
    Dim pending As Scripting.Dictionary, request As Variant, key As String
    key = CStr(viewer)
    request = NewHandoffRequest(viewer, subject)
    If IsEmpty(request) Then Exit Function
    If Not mPending.Exists(key) Then
        Set pending = New Scripting.Dictionary
        Call mPending.Add(key, pending)
    End If
    Set pending = mPending(key)
    If pending.Exists(CStr(subject)) Then
        DeferCreate = True
        Exit Function
    End If
    If pending.Count >= ADJACENT_PLAYER_LIMIT And Not pending.Exists(CStr(subject)) Then
        Call ReportFailure("handoff capacity exhausted")
        Exit Function
    End If
    pending(CStr(subject)) = request
    DeferCreate = True
End Function

Private Function NewHandoffRequest(ByVal viewer As Integer, ByVal subject As Integer) As Variant
    Dim session As Variant, request As Variant
    If Not ReadyUser(viewer) Or Not ReadyUser(subject) Then Exit Function
    session = Identity(subject)
    If IsEmpty(session) Then Exit Function
    request = Array(session(0), session(1), UserList(subject).pos.Map, UserList(subject).Char.charindex, ViewerSignature(viewer), HANDOFF_FRESH, -1&, Empty)
    If HandoffIsCurrent(viewer, subject, request) Then NewHandoffRequest = request
End Function

Private Function SendHandoffBlank(ByVal viewer As Integer, ByVal subject As Integer) As Boolean
    On Error GoTo Failed
    ' HOO treats this exact zero-appearance PC as initially invisible. Ordinary
    ' CharacterChange traffic therefore cannot reveal it if SetInvisible fails.
    With UserList(subject)
        Call WriteCharacterCreate(viewer, 0, 0, .Char.Heading, .Char.charindex, .pos.x, .pos.y, 0, 0, 0, 0, 0, 0, 0, _
            vbNullString, .Faccion.Status, .flags.Privilegios, 0, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, _
            .Char.speeding, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, 0, 0, 0, SendHandoffBlank)
    End With
    If Not SendHandoffBlank Then Call ReportFailure("handoff blank enqueue")
    Exit Function
Failed:
    Call ReportFailure("handoff blank writer", Err.Number, Err.Description)
    Call Protocol_Writes.GetWriterBuffer().Clear()
End Function

Private Function HandoffIsCurrent(ByVal viewer As Integer, ByVal subject As Integer, ByVal value As Variant) As Boolean
    If Not ReadyUser(viewer) Or Not ReadyUser(subject) Then Exit Function
    If Not SameIdentity(value, Identity(subject)) Or value(4) <> ViewerSignature(viewer) Then Exit Function
    With UserList(subject)
        If .pos.Map <> value(2) Or .pos.Map <> UserList(viewer).pos.Map Or .Char.charindex <> value(3) Then Exit Function
        If (UserList(viewer).AreasInfo.AreaReciveX And .AreasInfo.AreaPerteneceX) = 0 Then Exit Function
        If (UserList(viewer).AreasInfo.AreaReciveY And .AreasInfo.AreaPerteneceY) = 0 Then Exit Function
    End With
    HandoffIsCurrent = True
End Function

Private Function RemoveHandoffRecord(ByVal viewer As Integer, ByVal subject As Integer, ByVal value As Variant) As Boolean
    On Error GoTo Failed
    RemoveHandoffRecord = True
    If Not mActive.Exists(CStr(viewer)) Then Exit Function
    Dim active As Scripting.Dictionary, owner As Variant, key As String, queued As Boolean
    Set active = mActive(CStr(viewer))
    key = CStr(value(3))
    If Not active.Exists(key) Then Exit Function
    owner = active(key)
    ' Reused normal indices belong to their last accepted create, not the old
    ' pending user slot. Never remove a newer player/NPC occupying this index.
    If Not SameIdentity(owner, value) Or owner(4) <> subject Or owner(2) <> value(2) Then Exit Function
    RemoveHandoffRecord = False
    Call modSendData.SendData(ToIndex, viewer, PrepareMessageCharacterRemove(4, CInt(value(3)), False, True), False, queued, True)
    If Not queued Then
        Call ReportFailure("handoff cleanup enqueue")
        Exit Function
    End If
    Call active.Remove(key)
    RemoveHandoffRecord = True
    Exit Function
Failed:
    RemoveHandoffRecord = False
    Call ReportFailure("handoff cleanup", Err.Number, Err.Description)
End Function

Private Function CancelHandoff(ByVal viewer As Integer, ByVal subject As Integer, ByVal pending As Scripting.Dictionary, ByVal value As Variant) As Boolean
    value(5) = HANDOFF_CANCEL
    pending(CStr(subject)) = value
    If Not RemoveHandoffRecord(viewer, subject, value) Then Exit Function
    Call pending.Remove(CStr(subject))
    ' At most one explicitly requested successor shares this bounded entry.
    ' A retry of cleanup alone must never discover a new occupant of the slot.
    Dim replacement As Variant
    replacement = value(HANDOFF_REPLACEMENT)
    If Not IsEmpty(replacement) Then
        If HandoffIsCurrent(viewer, subject, replacement) Then pending(CStr(subject)) = replacement
    End If
    CancelHandoff = True
End Function

Public Sub CompleteAdjacentPlayerCreate(ByVal viewer As Integer, ByVal subject As Integer, ByVal queued As Boolean)
    #If PYMMO = 0 Then
        Exit Sub
    #End If
    On Error GoTo Failed
    Call EnsureState()
    If Not queued Then Exit Sub
    If mPending.Exists(CStr(viewer)) Then
        Dim pending As Scripting.Dictionary
        Set pending = mPending(CStr(viewer))
        If pending.Exists(CStr(subject)) Then Call pending.Remove(CStr(subject))
    End If
    Exit Sub
Failed:
    Call ReportFailure("complete handoff", Err.Number, Err.Description)
End Sub

Public Sub RecordAdjacentPlayerActive(ByVal viewer As Integer, ByVal charindex As Integer, ByVal isNpc As Boolean)
    #If PYMMO = 0 Then
        Exit Sub
    #End If
    On Error GoTo Failed
    If charindex < 1 Or charindex > UBound(CharList) Then Exit Sub
    Call EnsureState()
    Dim active As Scripting.Dictionary, subject As Integer, session As Variant
    If Not mActive.Exists(CStr(viewer)) Then
        Set active = New Scripting.Dictionary
        Call mActive.Add(CStr(viewer), active)
    End If
    Set active = mActive(CStr(viewer))
    ' Key by the bounded normal-character namespace, not user slot: a rejected
    ' normal remove can leave several old indices for one subject on the client.
    If active.Exists(CStr(charindex)) Then Call active.Remove(CStr(charindex))
    If isNpc Then Exit Sub
    subject = CharList(charindex)
    If Not LiveUser(subject) Then Exit Sub
    session = Identity(subject)
    If IsEmpty(session) Then Exit Sub
    active(CStr(charindex)) = Array(session(0), session(1), UserList(subject).pos.Map, charindex, subject)
    Exit Sub
Failed:
    Call ReportFailure("record active", Err.Number, Err.Description)
End Sub

Private Function RetireActive(ByVal viewer As Integer, ByVal subject As Integer) As Boolean
    On Error GoTo Failed
    RetireActive = True
    If Not mActive.Exists(CStr(viewer)) Then Exit Function
    Dim active As Scripting.Dictionary, value As Variant, queued As Boolean, key As Variant
    Set active = mActive(CStr(viewer))
    For Each key In active.Keys
        value = active(key)
        If value(4) = subject Then
            RetireActive = False
            Call modSendData.SendData(ToIndex, viewer, PrepareMessageCharacterRemove(4, CInt(value(3)), False, True), False, queued, True)
            If Not queued Then
                Call ReportFailure("active remove enqueue")
                Exit Function
            End If
            Call active.Remove(key)
        End If
    Next key
    RetireActive = True
    Exit Function
Failed:
    RetireActive = False
    Call ReportFailure("retire active", Err.Number, Err.Description)
End Function

Private Sub RetryHandoffs(ByVal viewer As Integer)
    If Not mPending.Exists(CStr(viewer)) Then Exit Sub
    Dim pending As Scripting.Dictionary, key As Variant, value As Variant, subject As Integer, valid As Boolean
    Set pending = mPending(CStr(viewer))
    For Each key In pending.Keys
        subject = CInt(key): value = pending(key): valid = False
        valid = value(5) <> HANDOFF_CANCEL And HandoffIsCurrent(viewer, subject, value)
        If Not valid Then
            If CancelHandoff(viewer, subject, pending, value) Then
                If pending.Exists(key) Then valid = HandoffIsCurrent(viewer, subject, pending(key))
            End If
        End If
        If valid Then
            ' Normal-map admission is checked by the shared create hook. Hidden
            ' ordinary PCs need normal records for collision and later reveal.
            Call MakeUserChar(False, viewer, subject, UserList(subject).pos.Map, UserList(subject).pos.x, UserList(subject).pos.y, 0)
        End If
    Next key
End Sub

Private Sub ReportFailure(ByVal operation As String, Optional ByVal number As Long = 0, Optional ByVal description As String = "")
    Dim now As Long
    now = GetTickCountRaw()
    If mLogged Then
        If TicksElapsed(mLastLog, now) < LOG_MS Then Exit Sub
    End If
    mLogged = True: mLastLog = now
    Call TraceError(number, operation & ": " & description, "modAdjacentPlayerVisibility", 0)
End Sub

#If UNIT_TEST = 1 Then
Public Sub AdjacentPlayerTestReset()
    Set mKnown = Nothing
    Call EnsureState()
    mGeneration = 0: mLow = 0: mHigh = 0: mExhausted = False: mLastUpdate = 0
End Sub

Public Sub AdjacentPlayerTestRefresh(ByVal viewer As Integer)
    Call EnsureState()
    Call RefreshViewer(viewer)
End Sub

Public Function AdjacentPlayerTestCount(ByVal viewer As Integer) As Long
    If mKnown.Exists(CStr(viewer)) Then AdjacentPlayerTestCount = mKnown(CStr(viewer)).Count
End Function

Public Function AdjacentPlayerTestPendingCount(ByVal viewer As Integer) As Long
    If mPending.Exists(CStr(viewer)) Then AdjacentPlayerTestPendingCount = mPending(CStr(viewer)).Count
End Function

Public Function AdjacentPlayerTestValue(ByVal viewer As Integer, ByVal subject As Integer) As Variant
    AdjacentPlayerTestValue = mKnown(CStr(viewer))(CStr(subject))
End Function

Public Sub AdjacentPlayerTestCounters(ByVal low As Long, ByVal high As Long, ByVal epoch As Long)
    mLow = low: mHigh = high: mGeneration = epoch
End Sub
#End If
