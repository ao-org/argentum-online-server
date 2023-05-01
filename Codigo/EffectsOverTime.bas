Attribute VB_Name = "EffectsOverTime"
Option Explicit

Private LastUpdateTime As Long
Private UniqueIdCounter As Long
Const ACTIVE_EFFECTS_MIN_SIZE As Integer = 500
Private ActiveEffects As t_EffectOverTimeList
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
    Dim i As Integer
    Dim j As Integer
    Dim InitialSize As Integer
    If RunningInVB() Then
        InitialSize = 2
    Else
        InitialSize = INITIAL_POOL_SIZE
    End If
100 ReDim EffectPools(1 To e_EffectOverTimeType.EffectTypeCount - 1) As t_EffectOverTimeList
102 For i = 1 To e_EffectOverTimeType.EffectTypeCount - 1
104     ReDim EffectPools(i).EffectList(InitialSize) As IBaseEffectOverTime
106     For j = 0 To InitialSize
108         Call AddEffect(EffectPools(i), InstantiateEOT(i))
110     Next j
    Next i
    Exit Sub
InitializePools_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.InitializePools", Erl)
End Sub

Public Sub UpdateEffectOverTime()
On Error GoTo Update_Err
    Dim CurrTime As Long
    Dim ElapsedTime As Long
100 CurrTime = GetTickCount()
102 If CurrTime < LastUpdateTime Then ' GetTickCount can overflow se we take care of that
104     ElapsedTime = 0
    Else
106     ElapsedTime = CurrTime - LastUpdateTime
    End If
108 LastUpdateTime = CurrTime
    
    
    Dim i As Integer
200 Do While i < ActiveEffects.EffectCount
202     If UpdateEffect(i, ElapsedTime) Then
204         i = i + 1
        End If
    Loop
    Exit Sub
Update_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.Update", Erl)
End Sub

Private Function UpdateEffect(ByVal Index As Integer, ByVal ElapsedTime As Long) As Boolean
On Error GoTo UpdateEffect_Err
    'this should never happend but it covers us for breaking all effects if something goes wrong
100 If ActiveEffects.EffectList(index) Is Nothing Then
102     UpdateEffect = True
        Exit Function
    End If
    Dim CurrentEffect As IBaseEffectOverTime
    Set CurrentEffect = ActiveEffects.EffectList(index)
104 CurrentEffect.Update (ElapsedTime)
106 If CurrentEffect.RemoveMe Then
108     If CurrentEffect.TargetIsValid Then
110         If CurrentEffect.TargetRefType = eUser Then
112             Call RemoveEffect(UserList(CurrentEffect.TargetArrayIndex).EffectOverTime, CurrentEffect)
114         ElseIf CurrentEffect.TargetRefType = eNpc Then
116             Call RemoveEffect(NpcList(CurrentEffect.TargetArrayIndex).EffectOverTime, CurrentEffect)
            End If
        End If
        Call RemoveEffectAtPos(ActiveEffects, index)
120     Call RecycleEffect(CurrentEffect)
134     UpdateEffect = False
    Else
138     UpdateEffect = True
    End If
    Exit Function
UpdateEffect_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.UpdateEffect", Erl)
    Set ActiveEffects.EffectList(index) = Nothing
    UpdateEffect = True
End Function

Private Function GetNextId() As Long
    UniqueIdCounter = (UniqueIdCounter + 1) And &H7FFFFFFF
    GetNextId = UniqueIdCounter
End Function

Public Sub CreateEffect(ByVal sourceIndex As Integer, ByVal sourceType As e_ReferenceType, _
                                  ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
                                  ByVal EffectIndex As Integer)
On Error GoTo CreateEffect_Err
    Dim EffectType As e_EffectOverTimeType
100 EffectType = EffectOverTime(EffectIndex).Type
    Select Case EffectType
        Case e_EffectOverTimeType.eHealthModifier
102         Dim Dot As UpdateHpOverTime
104         Set Dot = GetEOT(EffectType)
106         UniqueIdCounter = GetNextId()
108         Call Dot.Setup(sourceIndex, sourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
110         Call AddEffectToUpdate(Dot)
112         If TargetType = eUser Then
114             Call AddEffect(UserList(TargetIndex).EffectOverTime, Dot)
116         ElseIf TargetType = eNpc Then
118             Call AddEffect(NpcList(TargetIndex).EffectOverTime, Dot)
            End If
        Case e_EffectOverTimeType.eApplyModifiers
130         Dim StatDot As StatModifier
132         Set StatDot = GetEOT(EffectType)
134         UniqueIdCounter = GetNextId()
136         Call StatDot.Setup(sourceIndex, sourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
138         Call AddEffectToUpdate(StatDot)
140         If TargetType = eUser Then
142             Call AddEffect(UserList(TargetIndex).EffectOverTime, StatDot)
144         ElseIf TargetType = eNpc Then
146             Call AddEffect(NpcList(TargetIndex).EffectOverTime, StatDot)
            End If
        Case e_EffectOverTimeType.eProvoke
150         Dim Provoke As EffectProvoke
152         Set Provoke = GetEOT(EffectType)
154         UniqueIdCounter = GetNextId()
156         Call Provoke.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
158         Call AddEffectToUpdate(Provoke)
160         If TargetType = eUser Then
162             Call AddEffect(UserList(TargetIndex).EffectOverTime, Provoke)
164         ElseIf TargetType = eNpc Then
166             Call AddEffect(NpcList(TargetIndex).EffectOverTime, Provoke)
            End If
        Case e_EffectOverTimeType.eProvoked
170         Dim StatProvoked As EffectProvoked
172         Set StatProvoked = GetEOT(EffectType)
174         UniqueIdCounter = GetNextId()
176         Call StatProvoked.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
178         Call AddEffectToUpdate(StatProvoked)
180         If TargetType = eUser Then
182             Call AddEffect(UserList(TargetIndex).EffectOverTime, StatProvoked)
184         ElseIf TargetType = eNpc Then
186             Call AddEffect(NpcList(TargetIndex).EffectOverTime, StatProvoked)
            End If
        Case e_EffectOverTimeType.eDrunk
190         Dim Drunk As DrunkEffect
192         Set Drunk = GetEOT(EffectType)
194         UniqueIdCounter = GetNextId()
196         Call Drunk.Setup(SourceIndex, SourceType, EffectIndex, UniqueIdCounter)
198         Call AddEffectToUpdate(Drunk)
200         If TargetType = eUser Then
202             Call AddEffect(UserList(TargetIndex).EffectOverTime, Drunk)
204         ElseIf TargetType = eNpc Then
206             Call AddEffect(NpcList(TargetIndex).EffectOverTime, Drunk)
            End If
        Case e_EffectOverTimeType.eTranslation
230      Dim TE As TranslationEffect
232         Set TE = GetEOT(EffectType)
236         UniqueIdCounter = GetNextId()
238         Call TE.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
240         Call AddEffectToUpdate(TE)
242         If TargetType = eUser Then
244             Call AddEffect(UserList(TargetIndex).EffectOverTime, TE)
246         ElseIf TargetType = eNpc Then
248             Call AddEffect(NpcList(TargetIndex).EffectOverTime, TE)
            End If
        Case e_EffectOverTimeType.eApplyEffectOnHit
390         Dim EOH As ApplyEffectOnHit
392         Set EOH = GetEOT(EffectType)
394         UniqueIdCounter = GetNextId()
396         Call EOH.Setup(SourceIndex, SourceType, EffectIndex, UniqueIdCounter)
398         Call AddEffectToUpdate(EOH)
400         If TargetType = eUser Then
402             Call AddEffect(UserList(TargetIndex).EffectOverTime, EOH)
404         ElseIf TargetType = eNpc Then
406             Call AddEffect(NpcList(TargetIndex).EffectOverTime, EOH)
            End If
        Case e_EffectOverTimeType.eManaModifier
420         Dim Mot As UpdateManaOverTime
422         Set Mot = GetEOT(EffectType)
426         UniqueIdCounter = GetNextId()
428         Call Mot.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
430         Call AddEffectToUpdate(Mot)
432         If TargetType = eUser Then
434             Call AddEffect(UserList(TargetIndex).EffectOverTime, Mot)
438             'npc doesn't have mana
            End If
        Case e_EffectOverTimeType.ePartyBonus
450         Dim PartyEffect As ApplyEffectToParty
452         Set PartyEffect = GetEOT(EffectType)
456         UniqueIdCounter = GetNextId()
458         Call PartyEffect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
460         Call AddEffectToUpdate(PartyEffect)
462         If TargetType = eUser Then
464             Call AddEffect(UserList(TargetIndex).EffectOverTime, PartyEffect)
468             'npc doesn't have groups
            End If
        Case e_EffectOverTimeType.ePullTarget
490         Dim PullEffect As AttrackEffect
492         Set PullEffect = GetEOT(EffectType)
494         UniqueIdCounter = GetNextId()
496         Call PullEffect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
498         Call AddEffectToUpdate(PullEffect)
500         If TargetType = eUser Then
502             Call AddEffect(UserList(TargetIndex).EffectOverTime, PullEffect)
504         ElseIf TargetType = eNpc Then
506             Call AddEffect(NpcList(TargetIndex).EffectOverTime, PullEffect)
            End If
        Case e_EffectOverTimeType.eMultipleAttacks
590         Dim MultiAttacks As MultipleAttacks
592         Set MultiAttacks = GetEOT(EffectType)
594         UniqueIdCounter = GetNextId()
596         Call MultiAttacks.Setup(SourceIndex, SourceType, EffectIndex, UniqueIdCounter)
598         Call AddEffectToUpdate(MultiAttacks)
600         If TargetType = eUser Then
602             Call AddEffect(UserList(TargetIndex).EffectOverTime, MultiAttacks)
604         ElseIf TargetType = eNpc Then
606             Call AddEffect(NpcList(TargetIndex).EffectOverTime, MultiAttacks)
            End If
        Case e_EffectOverTimeType.eProtection
610         Dim Protect As ProtectEffect
612         Set Protect = GetEOT(EffectType)
614         UniqueIdCounter = GetNextId()
616         Call Protect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
618         Call AddEffectToUpdate(Protect)
620         If TargetType = eUser Then
622             Call AddEffect(UserList(TargetIndex).EffectOverTime, Protect)
624         ElseIf TargetType = eNpc Then
626             Call AddEffect(NpcList(TargetIndex).EffectOverTime, Protect)
            End If
        Case e_EffectOverTimeType.eTransform
630         Dim Transform As TransformEffect
632         Set Transform = GetEOT(EffectType)
634         UniqueIdCounter = GetNextId()
636         Call Transform.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
638         Call AddEffectToUpdate(Transform)
640         If TargetType = eUser Then
642             Call AddEffect(UserList(TargetIndex).EffectOverTime, Transform)
644         ElseIf TargetType = eNpc Then
646             Call AddEffect(NpcList(TargetIndex).EffectOverTime, Transform)
            End If
        Case e_EffectOverTimeType.eBonusDamage
650         Dim BonusDamage As BonusDamageEffect
652         Set BonusDamage = GetEOT(EffectType)
654         UniqueIdCounter = GetNextId()
656         Call BonusDamage.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
658         Call AddEffectToUpdate(BonusDamage)
660         If TargetType = eUser Then
662             Call AddEffect(UserList(TargetIndex).EffectOverTime, BonusDamage)
664         ElseIf TargetType = eNpc Then
666             Call AddEffect(NpcList(TargetIndex).EffectOverTime, BonusDamage)
            End If
        Case Else
            Debug.Assert False
    End Select
    Exit Sub
CreateEffect_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateEffect", Erl)
End Sub

Public Sub CreateTrap(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal map As Integer, ByVal TileX As Integer, ByVal TileY As Integer, ByVal EffectTypeId As Integer)
On Error GoTo CreateTrap_Err
    Dim EffectType As e_EffectOverTimeType
100 EffectType = e_EffectOverTimeType.eTrap
    Dim Trap As clsTrap
104 Set Trap = GetEOT(EffectType)
106 UniqueIdCounter = GetNextId()
108 Call Trap.Setup(SourceIndex, SourceType, EffectTypeId, UniqueIdCounter, map, TileX, TileY)
110 Call AddEffectToUpdate(Trap)
112 If SourceType = eUser Then
114     Call AddEffect(UserList(SourceIndex).EffectOverTime, Trap)
116 ElseIf SourceType = eNpc Then
118     Call AddEffect(NpcList(SourceIndex).EffectOverTime, Trap)
    End If
    Exit Sub
CreateTrap_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateTrap", Erl)
End Sub

Public Sub CreateDelayedBlast(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal Map As Integer, ByVal TileX As Integer, _
                              ByVal TileY As Integer, ByVal EffectTypeId As Integer, ByVal SourceObjIndex As Integer)
On Error GoTo CreateDelayedBlast_Err
    Dim EffectType As e_EffectOverTimeType
100 EffectType = e_EffectOverTimeType.eDelayedBlast
    Dim Blast As DelayedBlast
104 Set Blast = GetEOT(EffectType)
106 UniqueIdCounter = GetNextId()
108 Call Blast.Setup(SourceIndex, SourceType, EffectTypeId, UniqueIdCounter, Map, TileX, TileY, SourceObjIndex)
110 Call AddEffectToUpdate(Blast)
112 If SourceType = eUser Then
114     Call AddEffect(UserList(SourceIndex).EffectOverTime, Blast)
116 ElseIf SourceType = eNpc Then
118     Call AddEffect(NpcList(SourceIndex).EffectOverTime, Blast)
    End If
    Exit Sub
CreateDelayedBlast_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateTrap", Erl)
End Sub

Public Sub CreateUnequip(ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, ByVal ItemSlotType As Long)
On Error GoTo CreateDelayedBlast_Err
    If Not IsFeatureEnabled("bandit_unequip_bonus") Then Exit Sub
    Dim EffectType As e_EffectOverTimeType
100 EffectType = e_EffectOverTimeType.eUnequip
    Dim Unequip As UnequipItem
104 Set Unequip = GetEOT(EffectType)
106 UniqueIdCounter = GetNextId()
108 Call Unequip.Setup(TargetIndex, TargetType, UnequipEffectId, UniqueIdCounter, ItemSlotType)
110 Call AddEffectToUpdate(Unequip)
112 If TargetType = eUser Then
114     Call AddEffect(UserList(TargetIndex).EffectOverTime, Unequip)
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
100 Set GetEOT = Nothing
102 If EffectPools(EffectType).EffectCount = 0 Then
104     Set GetEOT = InstantiateEOT(EffectType)
        Exit Function
    End If
108 Set GetEOT = EffectPools(EffectType).EffectList(EffectPools(EffectType).EffectCount - 1)
120 Set EffectPools(EffectType).EffectList(EffectPools(EffectType).EffectCount - 1) = Nothing
126 EffectPools(EffectType).EffectCount = EffectPools(EffectType).EffectCount - 1
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
100 If Not IsArrayInitialized(EffectList.EffectList) Then
104     ReDim EffectList.EffectList(ACTIVE_EFFECT_LIST_SIZE) As IBaseEffectOverTime
    ElseIf EffectList.EffectCount >= UBound(EffectList.EffectList) Then
108     ReDim Preserve EffectList.EffectList(EffectList.EffectCount * 1.2) As IBaseEffectOverTime
    End If
116 Set EffectList.EffectList(EffectList.EffectCount) = Effect
    Call SetMask(EffectList.CallbaclMask, Effect.CallBacksMask)
120 EffectList.EffectCount = EffectList.EffectCount + 1
    Exit Sub
AddEffect_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.AddEffect", Erl)
End Sub

Public Sub RemoveEffect(ByRef EffectList As t_EffectOverTimeList, ByRef Effect As IBaseEffectOverTime)
On Error GoTo RemoveEffect_Err
    Dim i As Integer
100 For i = 0 To EffectList.EffectCount - 1
106     If EffectList.EffectList(i).UniqueId() = Effect.UniqueId() Then
            Call RemoveEffectAtPos(EffectList, i)
            Exit Sub
        End If
    Next i
    Exit Sub
RemoveEffect_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.RemoveEffect", Erl)
End Sub

Public Function FindEffectOnTarget(ByVal CasterIndex As Integer, ByRef EffectList As t_EffectOverTimeList, ByVal EffectId As Integer) As IBaseEffectOverTime
On Error GoTo FindEffectOnTarget_Err
100 Set FindEffectOnTarget = Nothing
102 Dim EffectLimit As e_EOTTargetLimit
104 EffectLimit = EffectOverTime(EffectId).Limit
106 Dim i As Integer
108 If EffectLimit = e_EOTTargetLimit.eAny Then
        Exit Function
    End If
120 For i = 0 To EffectList.EffectCount - 1
        If EffectLimit = eSingle Or EffectLimit = eSingleByCaster Then
126         If EffectList.EffectList(i).EotId = EffectId Then
130             If EffectLimit = eSingle Then
132                 Set FindEffectOnTarget = EffectList.EffectList(i)
                    Exit Function
                Else
140                 If EffectList.EffectList(i).CasterRefType = eUser Then
142                     If EffectList.EffectList(i).CasterUserId = UserList(CasterIndex).ID Then
144                         Set FindEffectOnTarget = EffectList.EffectList(i)
                            Exit Function
                        End If
150                 ElseIf EffectList.EffectList(i).CasterRefType = eNpc Then
152                     If EffectList.EffectList(i).CasterIsValid Then
154                         Set FindEffectOnTarget = EffectList.EffectList(i)
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
        End If
    Next i
    Exit Function
FindEffectOnTarget_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.FindEffectOnTarget", Erl)
End Function

Public Sub ClearEffectList(ByRef EffectList As t_EffectOverTimeList, Optional ByVal Filter As e_EffectType = e_EffectType.eAny, Optional ByVal ClearForDeath As Boolean = False)
On Error GoTo ClearEffectList_Err
    Dim i As Integer
100 Do While i < EffectList.EffectCount
102     If (Filter = e_EffectType.eAny Or Filter = EffectList.EffectList(i).EffectType) And _
           Not (ClearForDeath And EffectList.EffectList(i).KeepAfterDead()) Then
104         EffectList.EffectList(i).RemoveMe = True
            Call RemoveEffectAtPos(EffectList, i)
        Else
112         i = i + 1
        End If
    Loop
Exit Sub
ClearEffectList_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.ClearEffectList", Erl)
End Sub

Public Sub RemoveEffectAtPos(ByRef EffectList As t_EffectOverTimeList, ByVal position As Integer)
On Error GoTo RemoveEffectAtPos_Err
    Dim RegenerateMask As Boolean
    RegenerateMask = EffectList.EffectList(Position).CallBacksMask > 0
    Call EffectList.EffectList(position).OnRemove
    Dim i As Integer
    For i = Position To EffectList.EffectCount - 1
106     Set EffectList.EffectList(i) = EffectList.EffectList(i + 1)
    Next i
108 Set EffectList.EffectList(EffectList.EffectCount - 1) = Nothing
110 EffectList.EffectCount = EffectList.EffectCount - 1
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

Public Function TargetApplyDamageReduction(ByRef EffectList As t_EffectOverTimeList, ByVal Damage As Long, ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType) As Long
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

Public Function ApplyEotModifier(ByRef TargetRef As t_AnyReference, ByRef EffectStats As t_EffectOverTime)
    If IsValidRef(TargetRef) Then
        Call UpdateIncreaseModifier(TargetRef, MagicBonus, EffectStats.MagicDamageDone)
        Call UpdateIncreaseModifier(TargetRef, PhysiccalBonus, EffectStats.PhysicalDamageDone)
        Call UpdateIncreaseModifier(TargetRef, MagicReduction, EffectStats.MagicDamageReduction)
        Call UpdateIncreaseModifier(TargetRef, PhysicalReduction, EffectStats.PhysicalDamageReduction)
        Call UpdateIncreaseModifier(TargetRef, MovementSpeed, EffectStats.SpeedModifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.HitBonus, EffectStats.HitModifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.EvasionBonus, EffectStats.EvasionModifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.SelfHealingBonus, EffectStats.SelfHealingBonus)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.MagicHealingBonus, EffectStats.MagicHealingBonus)
    End If
End Function

Public Function RemoveEotModifier(ByRef TargetRef As t_AnyReference, ByRef EffectStats As t_EffectOverTime)
    If IsValidRef(TargetRef) Then
        Call UpdateIncreaseModifier(TargetRef, MagicBonus, -EffectStats.MagicDamageDone)
        Call UpdateIncreaseModifier(TargetRef, PhysiccalBonus, -EffectStats.PhysicalDamageDone)
        Call UpdateIncreaseModifier(TargetRef, MagicReduction, -EffectStats.MagicDamageReduction)
        Call UpdateIncreaseModifier(TargetRef, PhysicalReduction, -EffectStats.PhysicalDamageReduction)
        Call UpdateIncreaseModifier(TargetRef, MovementSpeed, -EffectStats.SpeedModifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.HitBonus, -EffectStats.HitModifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.EvasionBonus, -EffectStats.EvasionModifier)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.SelfHealingBonus, -EffectStats.SelfHealingBonus)
        Call UpdateIncreaseModifier(TargetRef, e_ModifierTypes.MagicHealingBonus, -EffectStats.MagicHealingBonus)
    End If
End Function
