Attribute VB_Name = "modPartyEffects"
' Argentum 20 Game Server
' Copyright (C) 2023-2026 Noland Studios LTD
' Licensed under the GNU Affero General Public License, version 3 or later.

Option Explicit

Private Const MaxPartyEffects As Integer = 16

Private Function PartyEffectAllowed(ByVal ResourceId As Integer, ByVal Category As Byte) As Boolean
    ' Effects.ini 1-42 are public EOT effects; 43-49 are the supported statuses.
    PartyEffectAllowed = ResourceId >= 1 And ResourceId <= 49 And _
        (Category = eBuff Or Category = eDebuff)
End Function

Private Function EffectRemaining(ByRef Effect As t_PartyDisplayEffect) As Long
    If Effect.RemainingMs = -1 Then
        EffectRemaining = -1
        Exit Function
    End If
    Dim Elapsed As Double
    Elapsed = CDbl(GetTickCountRaw()) - Effect.ReceiptTick
    If Elapsed < 0 Then Elapsed = Elapsed + 4294967296#
    If Elapsed >= Effect.RemainingMs Then Exit Function
    EffectRemaining = CLng(Effect.RemainingMs - Elapsed)
End Function

Private Function PartySlotOf(ByVal LeaderIndex As Integer, ByVal MemberIndex As Integer) As Byte
    Dim Slot As Byte
    For Slot = 1 To UserList(LeaderIndex).Grupo.CantidadMiembros
        If IsValidUserRef(UserList(LeaderIndex).Grupo.Miembros(Slot)) Then
            If UserList(LeaderIndex).Grupo.Miembros(Slot).ArrayIndex = MemberIndex Then
                PartySlotOf = Slot
                Exit Function
            End If
        End If
    Next Slot
End Function

Private Sub WritePartyEffectFields(ByRef Effect As t_PartyDisplayEffect, ByVal RemainingMs As Long)
    Call Writer.WriteInt16(Effect.ResourceId)
    Call Writer.WriteInt32(Effect.InstanceId)
    Call Writer.WriteInt32(RemainingMs)
    Call Writer.WriteInt32(Effect.TotalMs)
    Call Writer.WriteInt8(Effect.Category)
    Call Writer.WriteInt16(Effect.Stacks)
End Sub

Private Sub SendPartyEffectUpdate(ByVal RecipientIndex As Integer, ByVal MemberSlot As Byte, _
                                  ByRef Effect As t_PartyDisplayEffect, ByVal RemainingMs As Long)
    On Error GoTo SendPartyEffectUpdate_Err
    Call Writer.WriteInt16(ServerPacketID.ePartyMemberEffectUpdate)
    Call Writer.WriteInt8(MemberSlot)
    Call WritePartyEffectFields(Effect, RemainingMs)
    Call modSendData.SendData(ToIndex, RecipientIndex)
    Exit Sub
SendPartyEffectUpdate_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "modPartyEffects.SendPartyEffectUpdate", Erl)
End Sub

Public Sub TrackPartyEffect(ByVal UserIndex As Integer, ByVal ResourceId As Integer, _
                            ByVal InstanceId As Long, ByVal RemainingMs As Long, _
                            ByVal TotalMs As Long, ByVal Category As Byte, ByVal Stacks As Integer)
    If Not PartyEffectAllowed(ResourceId, Category) Then Exit Sub
    Dim Index As Integer
    Dim FreeIndex As Integer
    For Index = 1 To MaxPartyEffects
        With UserList(UserIndex).PartyDisplayEffects(Index)
            If .ResourceId = ResourceId And .InstanceId = InstanceId Then
                FreeIndex = Index
                Exit For
            End If
            If .ResourceId <> 0 Then
                If EffectRemaining(UserList(UserIndex).PartyDisplayEffects(Index)) = 0 Then .ResourceId = 0
            End If
            If .ResourceId = 0 And FreeIndex = 0 Then FreeIndex = Index
        End With
    Next Index
    If FreeIndex = 0 Then Exit Sub
    Dim Effect As t_PartyDisplayEffect
    With Effect
        .ResourceId = ResourceId
        .InstanceId = InstanceId
        .RemainingMs = RemainingMs
        .TotalMs = TotalMs
        .Category = Category
        .Stacks = Stacks
        .ReceiptTick = CDbl(GetTickCountRaw())
    End With
    If RemainingMs = 0 Then
        UserList(UserIndex).PartyDisplayEffects(FreeIndex).ResourceId = 0
    Else
        UserList(UserIndex).PartyDisplayEffects(FreeIndex) = Effect
    End If
    If Not UserList(UserIndex).Grupo.EnGrupo Then Exit Sub
    If Not IsValidUserRef(UserList(UserIndex).Grupo.Lider) Then Exit Sub
    Dim LeaderIndex As Integer
    LeaderIndex = UserList(UserIndex).Grupo.Lider.ArrayIndex
    Dim MemberSlot As Byte
    MemberSlot = PartySlotOf(LeaderIndex, UserIndex)
    If MemberSlot = 0 Then Exit Sub
    Dim RecipientSlot As Byte
    Dim RecipientIndex As Integer
    For RecipientSlot = 1 To UserList(LeaderIndex).Grupo.CantidadMiembros
        If IsValidUserRef(UserList(LeaderIndex).Grupo.Miembros(RecipientSlot)) Then
            RecipientIndex = UserList(LeaderIndex).Grupo.Miembros(RecipientSlot).ArrayIndex
            If UserSupportsPartyEffects(RecipientIndex) Then
                Call SendPartyEffectUpdate(RecipientIndex, MemberSlot, Effect, RemainingMs)
            End If
        End If
    Next RecipientSlot
End Sub

Private Sub SendPartyEffectSnapshot(ByVal RecipientIndex As Integer, ByVal MemberIndex As Integer, _
                                    ByVal MemberSlot As Byte)
    On Error GoTo SendPartyEffectSnapshot_Err
    Dim Count As Byte
    Dim Index As Integer
    Dim Remaining(1 To MaxPartyEffects) As Long
    For Index = 1 To MaxPartyEffects
        With UserList(MemberIndex).PartyDisplayEffects(Index)
            If PartyEffectAllowed(.ResourceId, .Category) Then
                Remaining(Index) = EffectRemaining( _
                    UserList(MemberIndex).PartyDisplayEffects(Index))
                If Remaining(Index) <> 0 Then Count = Count + 1
            End If
        End With
    Next Index
    Call Writer.WriteInt16(ServerPacketID.ePartyMemberEffectsSnapshot)
    Call Writer.WriteInt8(MemberSlot)
    Call Writer.WriteInt8(Count)
    For Index = 1 To MaxPartyEffects
        With UserList(MemberIndex).PartyDisplayEffects(Index)
            If Remaining(Index) <> 0 Then
                Call WritePartyEffectFields( _
                    UserList(MemberIndex).PartyDisplayEffects(Index), Remaining(Index))
            End If
        End With
    Next Index
    Call modSendData.SendData(ToIndex, RecipientIndex)
    Exit Sub
SendPartyEffectSnapshot_Err:
    Call Writer.Clear
    Call TraceError(Err.Number, Err.Description, "modPartyEffects.SendPartyEffectSnapshot", Erl)
End Sub

Public Sub SendPartyEffectSnapshotsToUser(ByVal RecipientIndex As Integer)
    If Not UserSupportsPartyEffects(RecipientIndex) Then Exit Sub
    If Not UserList(RecipientIndex).Grupo.EnGrupo Then Exit Sub
    If Not IsValidUserRef(UserList(RecipientIndex).Grupo.Lider) Then Exit Sub
    Dim LeaderIndex As Integer
    LeaderIndex = UserList(RecipientIndex).Grupo.Lider.ArrayIndex
    Dim Slot As Byte
    For Slot = 1 To UserList(LeaderIndex).Grupo.CantidadMiembros
        If IsValidUserRef(UserList(LeaderIndex).Grupo.Miembros(Slot)) Then
            Call SendPartyEffectSnapshot(RecipientIndex, _
                UserList(LeaderIndex).Grupo.Miembros(Slot).ArrayIndex, Slot)
        End If
    Next Slot
End Sub

Public Sub SendPartyEffectSnapshotsForGroup(ByVal MemberIndex As Integer)
    If Not UserList(MemberIndex).Grupo.EnGrupo Then Exit Sub
    If Not IsValidUserRef(UserList(MemberIndex).Grupo.Lider) Then Exit Sub
    Dim LeaderIndex As Integer
    LeaderIndex = UserList(MemberIndex).Grupo.Lider.ArrayIndex
    Dim Slot As Byte
    For Slot = 1 To UserList(LeaderIndex).Grupo.CantidadMiembros
        If IsValidUserRef(UserList(LeaderIndex).Grupo.Miembros(Slot)) Then
            Call SendPartyEffectSnapshotsToUser( _
                UserList(LeaderIndex).Grupo.Miembros(Slot).ArrayIndex)
        End If
    Next Slot
End Sub

Public Sub ClearPartyEffectState(ByVal UserIndex As Integer)
    Dim Index As Integer
    For Index = 1 To MaxPartyEffects
        UserList(UserIndex).PartyDisplayEffects(Index).ResourceId = 0
    Next Index
End Sub
