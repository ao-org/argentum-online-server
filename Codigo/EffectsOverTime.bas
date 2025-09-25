Attribute VB_Name = "EffectsOverTime"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
Private LastUpdateTime        As Long
Private UniqueIdCounter       As Long
Const ACTIVE_EFFECTS_MIN_SIZE As Integer = 500
Private ActiveEffects         As t_EffectOverTimeList
Const UnequipEffectId = 23
Const INITIAL_POOL_SIZE = 200
Private EffectPools() As t_EffectOverTimeList

Public Enum e_EffectCallbackMask
    eTargetUseMagic = 1
    eTartgetWillAtack = 2
    eTartgetDidHit = 4
    eTargetFailedAttack = 8
    eTargetWasDamaged = 16
    eTargetWillAttackPosition = 32
    eTargetApplyDamageReduction = 64
    eTargetChangeTerrain = 128
End Enum

Public Sub InitializePools()
    On Error GoTo InitializePools_Err
    Dim i           As Integer
    Dim j           As Integer
    Dim InitialSize As Integer
    If RunningInVB() Then
        InitialSize = 2
    Else
        InitialSize = INITIAL_POOL_SIZE
    End If
    ReDim EffectPools(1 To e_EffectOverTimeType.EffectTypeCount - 1) As t_EffectOverTimeList
    For i = 1 To e_EffectOverTimeType.EffectTypeCount - 1
        ReDim EffectPools(i).EffectList(InitialSize) As IBaseEffectOverTime
        For j = 0 To InitialSize
            Call AddEffect(EffectPools(i), InstantiateEOT(i))
        Next j
    Next i
    Exit Sub
InitializePools_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.InitializePools", Erl)
End Sub

Public Sub UpdateEffectOverTime()
    On Error GoTo Update_Err
    Dim CurrTime    As Long
    Dim ElapsedTime As Long
    CurrTime = GetTickCount()
    If CurrTime < LastUpdateTime Then ' GetTickCount can overflow se we take care of that
        ElapsedTime = 0
    Else
        ElapsedTime = CurrTime - LastUpdateTime
    End If
    LastUpdateTime = CurrTime
    Dim i As Integer
    Do While i < ActiveEffects.EffectCount
        If UpdateEffect(i, ElapsedTime) Then
            i = i + 1
        End If
    Loop
    Exit Sub
Update_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.Update", Erl)
End Sub

Private Function UpdateEffect(ByVal Index As Integer, ByVal ElapsedTime As Long) As Boolean
    On Error GoTo UpdateEffect_Err
    'this should never happend but it covers us for breaking all effects if something goes wrong
    If ActiveEffects.EffectList(Index) Is Nothing Then
        UpdateEffect = True
        Exit Function
    End If
    Dim CurrentEffect As IBaseEffectOverTime
    Set CurrentEffect = ActiveEffects.EffectList(Index)
    CurrentEffect.Update (ElapsedTime)
    If CurrentEffect.RemoveMe Then
        If CurrentEffect.TargetIsValid Then
            If CurrentEffect.TargetRefType = eUser Then
                Call RemoveEffect(UserList(CurrentEffect.TargetArrayIndex).EffectOverTime, CurrentEffect)
            ElseIf CurrentEffect.TargetRefType = eNpc Then
                Call RemoveEffect(NpcList(CurrentEffect.TargetArrayIndex).EffectOverTime, CurrentEffect)
            End If
        End If
        Call RemoveEffectAtPos(ActiveEffects, Index)
        Call RecycleEffect(CurrentEffect)
        UpdateEffect = False
    Else
        UpdateEffect = True
    End If
    Exit Function
UpdateEffect_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.UpdateEffect", Erl)
    Set ActiveEffects.EffectList(Index) = Nothing
    UpdateEffect = True
End Function

Private Function GetNextId() As Long
    UniqueIdCounter = (UniqueIdCounter + 1) And &H7FFFFFFF
    GetNextId = UniqueIdCounter
End Function

Public Sub CreateEffect(ByVal SourceIndex As Integer, _
                        ByVal SourceType As e_ReferenceType, _
                        ByVal TargetIndex As Integer, _
                        ByVal TargetType As e_ReferenceType, _
                        ByVal EffectIndex As Integer)
    On Error GoTo CreateEffect_Err
    Dim EffectType As e_EffectOverTimeType
    EffectType = EffectOverTime(EffectIndex).Type
    Select Case EffectType
        Case e_EffectOverTimeType.eHealthModifier
            Dim Dot As UpdateHpOverTime
            Set Dot = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call Dot.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(Dot)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, Dot)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, Dot)
            End If
        Case e_EffectOverTimeType.eApplyModifiers
            Dim StatDot As StatModifier
            Set StatDot = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call StatDot.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(StatDot)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, StatDot)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, StatDot)
            End If
        Case e_EffectOverTimeType.eProvoke
            Dim Provoke As EffectProvoke
            Set Provoke = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call Provoke.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(Provoke)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, Provoke)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, Provoke)
            End If
        Case e_EffectOverTimeType.eProvoked
            Dim StatProvoked As EffectProvoked
            Set StatProvoked = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call StatProvoked.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(StatProvoked)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, StatProvoked)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, StatProvoked)
            End If
        Case e_EffectOverTimeType.eDrunk
            Dim Drunk As DrunkEffect
            Set Drunk = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call Drunk.Setup(SourceIndex, SourceType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(Drunk)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, Drunk)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, Drunk)
            End If
        Case e_EffectOverTimeType.eTranslation
            Dim TE As TranslationEffect
            Set TE = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call TE.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(TE)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, TE)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, TE)
            End If
        Case e_EffectOverTimeType.eApplyEffectOnHit
            Dim EOH As ApplyEffectOnHit
            Set EOH = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call EOH.Setup(SourceIndex, SourceType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(EOH)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, EOH)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, EOH)
            End If
        Case e_EffectOverTimeType.eManaModifier
            Dim Mot As UpdateManaOverTime
            Set Mot = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call Mot.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(Mot)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, Mot)
                'npc doesn't have mana
            End If
        Case e_EffectOverTimeType.ePartyBonus
            Dim PartyEffect As ApplyEffectToParty
            Set PartyEffect = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call PartyEffect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(PartyEffect)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, PartyEffect)
                'npc doesn't have groups
            End If
        Case e_EffectOverTimeType.ePullTarget
            Dim PullEffect As AttrackEffect
            Set PullEffect = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call PullEffect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(PullEffect)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, PullEffect)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, PullEffect)
            End If
        Case e_EffectOverTimeType.eMultipleAttacks
            Dim MultiAttacks As MultipleAttacks
            Set MultiAttacks = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call MultiAttacks.Setup(SourceIndex, SourceType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(MultiAttacks)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, MultiAttacks)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, MultiAttacks)
            End If
        Case e_EffectOverTimeType.eProtection
            Dim Protect As ProtectEffect
            Set Protect = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call Protect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(Protect)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, Protect)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, Protect)
            End If
        Case e_EffectOverTimeType.eTransform
            Dim Transform As TransformEffect
            Set Transform = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call Transform.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(Transform)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, Transform)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, Transform)
            End If
        Case e_EffectOverTimeType.eBonusDamage
            Dim BonusDamage As BonusDamageEffect
            Set BonusDamage = GetEOT(EffectType)
            UniqueIdCounter = GetNextId()
            Call BonusDamage.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
            Call AddEffectToUpdate(BonusDamage)
            If TargetType = eUser Then
                Call AddEffect(UserList(TargetIndex).EffectOverTime, BonusDamage)
            ElseIf TargetType = eNpc Then
                Call AddEffect(NpcList(TargetIndex).EffectOverTime, BonusDamage)
            End If
        Case Else
            Debug.Assert False
    End Select
    Exit Sub
CreateEffect_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateEffect EffectIndex:" & EffectIndex, Erl)
End Sub

Public Sub CreateTrap(ByVal SourceIndex As Integer, _
                      ByVal SourceType As e_ReferenceType, _
                      ByVal Map As Integer, _
                      ByVal TileX As Integer, _
                      ByVal TileY As Integer, _
                      ByVal EffectTypeId As Integer)
    On Error GoTo CreateTrap_Err
    Dim EffectType As e_EffectOverTimeType
    EffectType = e_EffectOverTimeType.eTrap
    Dim Trap As clsTrap
    Set Trap = GetEOT(EffectType)
    UniqueIdCounter = GetNextId()
    Call Trap.Setup(SourceIndex, SourceType, EffectTypeId, UniqueIdCounter, Map, TileX, TileY)
    Call AddEffectToUpdate(Trap)
    If SourceType = eUser Then
        Call AddEffect(UserList(SourceIndex).EffectOverTime, Trap)
    ElseIf SourceType = eNpc Then
        Call AddEffect(NpcList(SourceIndex).EffectOverTime, Trap)
    End If
    Exit Sub
CreateTrap_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateTrap", Erl)
End Sub

Public Sub CreateDelayedBlast(ByVal SourceIndex As Integer, _
                              ByVal SourceType As e_ReferenceType, _
                              ByVal Map As Integer, _
                              ByVal TileX As Integer, _
                              ByVal TileY As Integer, _
                              ByVal EffectTypeId As Integer, _
                              ByVal SourceObjIndex As Integer)
    On Error GoTo CreateDelayedBlast_Err
    Dim EffectType As e_EffectOverTimeType
    EffectType = e_EffectOverTimeType.eDelayedBlast
    Dim Blast As DelayedBlast
    Set Blast = GetEOT(EffectType)
    UniqueIdCounter = GetNextId()
    Call Blast.Setup(SourceIndex, SourceType, EffectTypeId, UniqueIdCounter, Map, TileX, TileY, SourceObjIndex)
    Call AddEffectToUpdate(Blast)
    If SourceType = eUser Then
        Call AddEffect(UserList(SourceIndex).EffectOverTime, Blast)
    ElseIf SourceType = eNpc Then
        Call AddEffect(NpcList(SourceIndex).EffectOverTime, Blast)
    End If
    Exit Sub
CreateDelayedBlast_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateTrap", Erl)
End Sub

Public Sub CreateUnequip(ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, ByVal ItemSlotType As Long)
    On Error GoTo CreateDelayedBlast_Err
    If Not IsFeatureEnabled("bandit_unequip_bonus") Then Exit Sub
    Dim EffectType As e_EffectOverTimeType
    EffectType = e_EffectOverTimeType.eUnequip
    Dim Unequip As UnequipItem
    Set Unequip = GetEOT(EffectType)
    UniqueIdCounter = GetNextId()
    Call Unequip.Setup(TargetIndex, TargetType, UnequipEffectId, UniqueIdCounter, ItemSlotType)
    Call AddEffectToUpdate(Unequip)
    If TargetType = eUser Then
        Call AddEffect(UserList(TargetIndex).EffectOverTime, Unequip)
    End If
    Exit Sub
CreateDelayedBlast_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateTrap", Erl)
End Sub

Private Function InstantiateEOT(ByVal EffectType As e_EffectOverTimeType) As IBaseEffectOverTime
    Select Case EffectType
        Case e_EffectOverTimeType.eHealthModifier
            Set InstantiateEOT = New UpdateHpOverTime
        Case e_EffectOverTimeType.eApplyModifiers
            Set InstantiateEOT = New StatModifier
        Case e_EffectOverTimeType.eProvoke
            Set InstantiateEOT = New EffectProvoke
        Case e_EffectOverTimeType.eProvoked
            Set InstantiateEOT = New EffectProvoked
        Case e_EffectOverTimeType.eTrap
            Set InstantiateEOT = New clsTrap
        Case e_EffectOverTimeType.eDrunk
            Set InstantiateEOT = New DrunkEffect
        Case e_EffectOverTimeType.eTranslation
            Set InstantiateEOT = New TranslationEffect
        Case e_EffectOverTimeType.eApplyEffectOnHit
            Set InstantiateEOT = New ApplyEffectOnHit
        Case e_EffectOverTimeType.eManaModifier
            Set InstantiateEOT = New UpdateManaOverTime
        Case e_EffectOverTimeType.ePartyBonus
            Set InstantiateEOT = New ApplyEffectToParty
        Case e_EffectOverTimeType.ePullTarget
            Set InstantiateEOT = New AttrackEffect
        Case e_EffectOverTimeType.eDelayedBlast
            Set InstantiateEOT = New DelayedBlast
        Case e_EffectOverTimeType.eUnequip
            Set InstantiateEOT = New UnequipItem
        Case e_EffectOverTimeType.eMultipleAttacks
            Set InstantiateEOT = New MultipleAttacks
        Case e_EffectOverTimeType.eProtection
            Set InstantiateEOT = New ProtectEffect
        Case e_EffectOverTimeType.eTransform
            Set InstantiateEOT = New TransformEffect
        Case e_EffectOverTimeType.eBonusDamage
            Set InstantiateEOT = New BonusDamageEffect
        Case Else
            Debug.Assert False
    End Select
End Function

Private Function GetEOT(ByVal EffectType As e_EffectOverTimeType) As IBaseEffectOverTime
    On Error GoTo GetEOT_Err
    Set GetEOT = Nothing
    If EffectPools(EffectType).EffectCount = 0 Then
        Set GetEOT = InstantiateEOT(EffectType)
        Exit Function
    End If
    Set GetEOT = EffectPools(EffectType).EffectList(EffectPools(EffectType).EffectCount - 1)
    Set EffectPools(EffectType).EffectList(EffectPools(EffectType).EffectCount - 1) = Nothing
    EffectPools(EffectType).EffectCount = EffectPools(EffectType).EffectCount - 1
    Exit Function
GetEOT_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.GetEOT", Erl)
End Function

Private Sub RecycleEffect(ByRef Effect As IBaseEffectOverTime)
    Call AddEffect(EffectPools(Effect.TypeId), Effect)
End Sub

Public Sub AddEffectToUpdate(ByRef Effect As IBaseEffectOverTime)
    On Error GoTo AddEffectToUpdate_Err
    Call AddEffect(ActiveEffects, Effect)
    Exit Sub
AddEffectToUpdate_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.AddEffectToUpdate", Erl)
End Sub

Public Sub AddEffect(ByRef EffectList As t_EffectOverTimeList, ByRef Effect As IBaseEffectOverTime)
    On Error GoTo AddEffect_Err
    If Not IsArrayInitialized(EffectList.EffectList) Then
        ReDim EffectList.EffectList(ACTIVE_EFFECT_LIST_SIZE) As IBaseEffectOverTime
    ElseIf EffectList.EffectCount >= UBound(EffectList.EffectList) Then
        ReDim Preserve EffectList.EffectList(EffectList.EffectCount * 1.2) As IBaseEffectOverTime
    End If
    Set EffectList.EffectList(EffectList.EffectCount) = Effect
    Call SetMask(EffectList.CallbaclMask, Effect.CallBacksMask)
    EffectList.EffectCount = EffectList.EffectCount + 1
    Exit Sub
AddEffect_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.AddEffect", Erl)
End Sub

Public Sub RemoveEffect(ByRef EffectList As t_EffectOverTimeList, ByRef Effect As IBaseEffectOverTime, Optional ByVal CallRemove As Boolean = True)
    On Error GoTo RemoveEffect_Err
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        If EffectList.EffectList(i).UniqueId() = Effect.UniqueId() Then
            Call RemoveEffectAtPos(EffectList, i, CallRemove)
            Exit Sub
        End If
    Next i
    Exit Sub
RemoveEffect_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.RemoveEffect", Erl)
End Sub

Public Function FindEffectOfTypeOnTarget(ByRef EffectList As t_EffectOverTimeList, ByVal TargetType As e_EffectType) As IBaseEffectOverTime
    On Error GoTo FindEffectOfTypeOnTarget_Err
    Set FindEffectOfTypeOnTarget = Nothing
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        If EffectList.EffectList(i).EffectType = TargetType Then
            Set FindEffectOfTypeOnTarget = EffectList.EffectList(i)
            Exit Function
        End If
    Next i
    Exit Function
FindEffectOfTypeOnTarget_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.FindEffectOnTarget", Erl)
End Function

Public Function FindEffectOnTarget(ByVal CasterIndex As Integer, ByRef EffectList As t_EffectOverTimeList, ByVal EffectId As Integer) As IBaseEffectOverTime
    On Error GoTo FindEffectOnTarget_Err
    Set FindEffectOnTarget = Nothing
    Dim EffectLimit As e_EOTTargetLimit
    EffectLimit = EffectOverTime(EffectId).Limit
    Dim i As Integer
    If EffectLimit = e_EOTTargetLimit.eAny Then
        Exit Function
    End If
    For i = 0 To EffectList.EffectCount - 1
        If EffectLimit = eSingle Or EffectLimit = eSingleByCaster Then
            If EffectList.EffectList(i).EotId = EffectId Then
                If EffectLimit = eSingle Then
                    Set FindEffectOnTarget = EffectList.EffectList(i)
                    Exit Function
                Else
                    If EffectList.EffectList(i).CasterRefType = eUser Then
                        If EffectList.EffectList(i).CasterUserId = UserList(CasterIndex).Id Then
                            Set FindEffectOnTarget = EffectList.EffectList(i)
                            Exit Function
                        End If
                    ElseIf EffectList.EffectList(i).CasterRefType = eNpc Then
                        If EffectList.EffectList(i).CasterIsValid And EffectList.EffectList(i).CasterArrayIndex = CasterIndex Then
                            Set FindEffectOnTarget = EffectList.EffectList(i)
                            Exit Function
                        End If
                    End If
                End If
            End If
        ElseIf EffectLimit = eSingleByType Then
            If EffectList.EffectList(i).TypeId = EffectOverTime(EffectId).Type Then
                Set FindEffectOnTarget = EffectList.EffectList(i)
                Exit Function
            End If
        ElseIf EffectLimit = eSingleByTypeId Then
            If EffectList.EffectList(i).SharedTypeId = EffectOverTime(EffectId).SharedTypeId Then
                Set FindEffectOnTarget = EffectList.EffectList(i)
                Exit Function
            End If
        End If
    Next i
    Exit Function
FindEffectOnTarget_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.FindEffectOnTarget", Erl)
End Function

Public Sub ClearEffectList(ByRef EffectList As t_EffectOverTimeList, Optional ByVal Filter As e_EffectType = e_EffectType.eAny, Optional ByVal ClearForDeath As Boolean = False)
    On Error GoTo ClearEffectList_Err
    Dim i As Integer
    Do While i < EffectList.EffectCount
        If (Filter = e_EffectType.eAny Or Filter = EffectList.EffectList(i).EffectType) And Not (ClearForDeath And EffectList.EffectList(i).KeepAfterDead()) Then
            EffectList.EffectList(i).RemoveMe = True
            Call RemoveEffectAtPos(EffectList, i)
        Else
            i = i + 1
        End If
    Loop
    Exit Sub
ClearEffectList_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.ClearEffectList", Erl)
End Sub

Public Sub RemoveEffectAtPos(ByRef EffectList As t_EffectOverTimeList, ByVal Position As Integer, Optional ByVal CallRemove As Boolean = True)
    On Error GoTo RemoveEffectAtPos_Err
    Dim RegenerateMask As Boolean
    RegenerateMask = EffectList.EffectList(Position).CallBacksMask > 0
    If CallRemove Then Call EffectList.EffectList(Position).OnRemove
    Dim i As Integer
    For i = Position To EffectList.EffectCount - 1
        Set EffectList.EffectList(i) = EffectList.EffectList(i + 1)
    Next i
    Set EffectList.EffectList(EffectList.EffectCount - 1) = Nothing
    EffectList.EffectCount = EffectList.EffectCount - 1
    If RegenerateMask Then
        EffectList.CallbaclMask = 0
        For i = 0 To EffectList.EffectCount - 1
            Call SetMask(EffectList.CallbaclMask, EffectList.EffectList(i).CallBacksMask)
        Next i
    End If
    Exit Sub
RemoveEffectAtPos_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.RemoveEffectAtPos", Erl)
End Sub

Public Sub TargetUseMagic(ByRef EffectList As t_EffectOverTimeList, ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetUseMagic) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TargetUseMagic(TargetUserId, SourceType, MagicId)
    Next i
End Sub

Public Sub TartgetWillAtack(ByRef EffectList As t_EffectOverTimeList, ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTartgetWillAtack) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TartgetWillAtack(TargetUserId, SourceType, AttackType)
    Next i
End Sub

Public Sub TartgetDidHit(ByRef EffectList As t_EffectOverTimeList, ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTartgetDidHit) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TartgetDidHit(TargetUserId, SourceType, AttackType)
    Next i
End Sub

Public Sub TargetFailedAttack(ByRef EffectList As t_EffectOverTimeList, ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetFailedAttack) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TargetFailedAttack(TargetUserId, SourceType, AttackType)
    Next i
End Sub

Public Function TargetApplyDamageReduction(ByRef EffectList As t_EffectOverTimeList, _
                                           ByVal Damage As Long, _
                                           ByVal SourceUserId As Integer, _
                                           ByVal SourceType As e_ReferenceType, _
                                           ByVal AttackType As e_DamageSourceType) As Long
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetApplyDamageReduction) Then
        TargetApplyDamageReduction = Damage
        Exit Function
    End If
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Damage = EffectList.EffectList(i).ApplyDamageReduction(Damage, SourceUserId, SourceType, AttackType)
        If Damage >= 0 Then
            Exit Function
        End If
    Next i
    TargetApplyDamageReduction = Damage
End Function

Public Sub TargetWasDamaged(ByRef EffectList As t_EffectOverTimeList, ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetWasDamaged) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TargetWasDamaged(SourceUserId, SourceType, AttackType)
    Next i
End Sub

Public Sub TargetWillAttackPosition(ByRef EffectList As t_EffectOverTimeList, ByRef Position As t_WorldPos)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetWillAttackPosition) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TargetWillAttackPosition(Position.Map, Position.x, Position.y)
    Next i
End Sub

Public Sub TargetUpdateTerrain(ByRef EffectList As t_EffectOverTimeList)
    If Not IsSet(EffectList.CallbaclMask, e_EffectCallbackMask.eTargetChangeTerrain) Then Exit Sub
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
        Call EffectList.EffectList(i).TargetChangeTerrain
    Next i
End Sub

Public Sub ChangeOwner(ByVal CurrentOwner As Integer, _
                       ByVal CurrentOwnerType As e_ReferenceType, _
                       ByVal NewOwner As Integer, _
                       ByVal NewOwnerType As e_ReferenceType, _
                       ByRef Effect As IBaseEffectOverTime)
    If CurrentOwnerType = eUser Then
        Call RemoveEffect(UserList(CurrentOwner).EffectOverTime, Effect, False)
    Else
        Call RemoveEffect(NpcList(CurrentOwner).EffectOverTime, Effect, False)
    End If
    Dim PrevEffect As IBaseEffectOverTime
    If NewOwnerType = eUser Then
        Set PrevEffect = FindEffectOnTarget(Effect.CasterArrayIndex, UserList(NewOwner).EffectOverTime, Effect.EotId)
        If Not PrevEffect Is Nothing Then
            PrevEffect.RemoveMe = True
        End If
        If Effect.ChangeTarget(NewOwner, NewOwnerType) Then
            Call AddEffect(UserList(NewOwner).EffectOverTime, Effect)
        Else
            Effect.RemoveMe = True
        End If
    Else
        Set PrevEffect = FindEffectOnTarget(Effect.CasterArrayIndex, NpcList(NewOwner).EffectOverTime, Effect.EotId)
        If Not PrevEffect Is Nothing Then
            If Not EffectOverTime(Effect.EotId).Override Then
                Effect.RemoveMe = True
                Exit Sub
            End If
        End If
        If Effect.ChangeTarget(NewOwner, NewOwnerType) Then
            Call AddEffect(NpcList(NewOwner).EffectOverTime, Effect)
        Else
            Effect.RemoveMe = True
        End If
    End If
End Sub

Public Function ConvertToClientBuff(ByVal buffType As e_EffectType) As e_EffectType
    Select Case buffType
        Case e_EffectType.eInformativeBuff
            ConvertToClientBuff = eBuff
        Case e_EffectType.eInformativeDebuff
            ConvertToClientBuff = eDebuff
        Case Else
            ConvertToClientBuff = buffType
    End Select
End Function

Public Function ApplyEotModifier(ByRef TargetRef As t_AnyReference, ByRef EffectStats As t_EffectOverTime, Optional ByVal Modifier As Single = 0)
    If IsValidRef(TargetRef) Then
        Call UpdateIncreaseModifier(TargetRef, MagicBonus, EffectStats.MagicDamageDone + EffectStats.MagicDamageDone * Modifier)
        Call UpdateIncreaseModifier(TargetRef, PhysiccalBonus, EffectStats.PhysicalDamageDone + EffectStats.PhysicalDamageDone * Modifier)
        Call UpdateIncreaseModifier(TargetRef, MagicReduction, EffectStats.MagicDamageReduction + EffectStats.MagicDamageReduction * Modifier)
        Call UpdateIncreaseModifier(TargetRef, PhysicalReduction, EffectStats.PhysicalDamageReduction + EffectStats.PhysicalDamageReduction * Modifier)
        Call UpdateIncreaseModifier(TargetRef, MovementSpeed, EffectStats.SpeedModifier + EffectStats.SpeedModifier * Modifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.HitBonus, EffectStats.HitModifier + EffectStats.HitModifier * Modifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.EvasionBonus, EffectStats.EvasionModifier + EffectStats.EvasionModifier * Modifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.SelfHealingBonus, EffectStats.SelfHealingBonus + EffectStats.SelfHealingBonus * Modifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.MagicHealingBonus, EffectStats.MagicHealingBonus + EffectStats.MagicHealingBonus * Modifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.PhysicalLinearBonus, EffectStats.PhysicalLinearBonus + EffectStats.PhysicalLinearBonus * Modifier)
        If IsSet(EffectStats.ApplyStatusMask, eCCInmunity) Then
            If TargetRef.RefType = eUser Then
                If UserList(TargetRef.ArrayIndex).flags.Inmovilizado = 1 Then
                    UserList(TargetRef.ArrayIndex).flags.Inmovilizado = 0
                    Call WriteInmovilizaOK(TargetRef.ArrayIndex)
                End If
                UserList(TargetRef.ArrayIndex).flags.Inmovilizado = 0
                UserList(TargetRef.ArrayIndex).Counters.Inmovilizado = 0
                UserList(TargetRef.ArrayIndex).Counters.Paralisis = 0
                UserList(TargetRef.ArrayIndex).flags.Paralizado = 0
            End If
            Call SetStatusMask(TargetRef, eCCInmunity)
        End If
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.DefenseBonus, EffectStats.DefenseBonus + EffectStats.DefenseBonus * Modifier)
    End If
End Function

Public Function RemoveEotModifier(ByRef TargetRef As t_AnyReference, ByRef EffectStats As t_EffectOverTime, Optional ByVal Modifier As Single = 0)
    If IsValidRef(TargetRef) Then
        Call UpdateIncreaseModifier(TargetRef, MagicBonus, -(EffectStats.MagicDamageDone + EffectStats.MagicDamageDone * Modifier))
        Call UpdateIncreaseModifier(TargetRef, PhysiccalBonus, -(EffectStats.PhysicalDamageDone + EffectStats.PhysicalDamageDone * Modifier))
        Call UpdateIncreaseModifier(TargetRef, MagicReduction, -(EffectStats.MagicDamageReduction + EffectStats.MagicDamageReduction * Modifier))
        Call UpdateIncreaseModifier(TargetRef, PhysicalReduction, -(EffectStats.PhysicalDamageReduction + EffectStats.PhysicalDamageReduction * Modifier))
        Call UpdateIncreaseModifier(TargetRef, MovementSpeed, -(EffectStats.SpeedModifier + EffectStats.SpeedModifier * Modifier))
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.HitBonus, -(EffectStats.HitModifier + EffectStats.HitModifier * Modifier))
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.EvasionBonus, -(EffectStats.EvasionModifier + EffectStats.EvasionModifier * Modifier))
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.SelfHealingBonus, -(EffectStats.SelfHealingBonus + EffectStats.SelfHealingBonus * Modifier))
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.MagicHealingBonus, -(EffectStats.MagicHealingBonus + EffectStats.MagicHealingBonus * Modifier))
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.PhysicalLinearBonus, -(EffectStats.PhysicalLinearBonus + EffectStats.PhysicalLinearBonus * Modifier))
        If IsSet(EffectStats.ApplyStatusMask, eCCInmunity) Then
            Call UnsetStatusMask(TargetRef, eCCInmunity)
        End If
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.DefenseBonus, -(EffectStats.DefenseBonus + EffectStats.DefenseBonus * Modifier))
    End If
End Function
