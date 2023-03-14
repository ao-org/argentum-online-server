VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBaseEffectOverTime

Private mWorldPosition As t_WorldPos
Private mEffect As e_TrapEffect
Private mid As Long
Private mDotInfo As t_BaseDotInfo
Private mElapsedTime As Long
Private mTimeLimit As Long

Private Enum e_TrapTypes
    eInmovilize = 1
    eExplosive
    ePoison
End Enum

Const ACTIVATE_TRAOP_WAW_FILE = 210

Public Sub Setup(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal EffectId As Integer, ByVal ID As Long, _
                 ByVal map As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
    mElapsedTime = 0
    Call SetRef(mDotInfo.TargetRef, SourceIndex, SourceType)
    mDotInfo.UniqueId = ID
    mDotInfo.RemoveEffect = False
    mDotInfo.EotId = EffectId
    Call SetPosition(map, TileX, TileY)
    With EffectOverTime(EffectId)
        mTimeLimit = .TickTime
        If EffectOverTime(mDotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(SourceIndex, .ClientEffectTypeId, ID, .TickTime, .buffType)
    End With
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eTrap
End Property

Public Property Get IBaseEffectOverTime_UniqueId() As Integer
    IBaseEffectOverTime_UniqueId = mDotInfo.UniqueId
End Property

Public Property Get IBaseEffectOverTime_CasterIsValid() As Boolean
    IBaseEffectOverTime_CasterIsValid = IsValidRef(mDotInfo.TargetRef)
End Property

Public Property Get IBaseEffectOverTime_CasterArrayIndex() As Integer
    IBaseEffectOverTime_CasterArrayIndex = mDotInfo.TargetRef.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_CasterRefType() As e_ReferenceType
    IBaseEffectOverTime_CasterRefType = mDotInfo.TargetRef.RefType
End Property

Public Property Get IBaseEffectOverTime_CasterUserId() As Long
    IBaseEffectOverTime_CasterUserId = mDotInfo.TargetRef.userID
End Property

Public Property Get IBaseEffectOverTime_TargetIsValid() As Boolean
    IBaseEffectOverTime_TargetIsValid = IsValidRef(mDotInfo.TargetRef)
End Property

Public Property Get IBaseEffectOverTime_TargetArrayIndex() As Integer
    IBaseEffectOverTime_TargetArrayIndex = mDotInfo.TargetRef.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_TargetRefType() As e_ReferenceType
    IBaseEffectOverTime_TargetRefType = mDotInfo.TargetRef.RefType
End Property

Public Property Get IBaseEffectOverTime_TargetUserId() As Long
    IBaseEffectOverTime_TargetUserId = mDotInfo.TargetRef.userID
End Property

Public Property Get IBaseEffectOverTime_EffectType() As e_EffectType
    IBaseEffectOverTime_EffectType = EffectOverTime(mDotInfo.EotId).buffType
End Property

Public Sub IBaseEffectOverTime_Update(ByVal deltaTime As Long)
    mElapsedTime = mElapsedTime + deltaTime
    If mElapsedTime >= mTimeLimit Or Not IsValidRef(mDotInfo.TargetRef) Then
        mDotInfo.RemoveEffect = True
        Exit Sub
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = mDotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal value As Boolean)
    mDotInfo.RemoveEffect = value
End Property

Public Sub IBaseEffectOverTime_Reset()
End Sub

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = mDotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If EffectOverTime(mDotInfo.EotId).ClientEffectTypeId > 0 Then
        Call WriteSendSkillCdUpdate(mDotInfo.TargetRef.ArrayIndex, EffectOverTime(mDotInfo.EotId).ClientEffectTypeId, _
                                    mDotInfo.UniqueId, 0, EffectOverTime(mDotInfo.EotId).buffType)
    End If
    Call RemoveFromMap
End Sub

Public Sub IBaseEffectOverTime_TargetUseMagic(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
End Sub

Public Sub IBaseEffectOverTime_TartgetWillAtack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TartgetDidHit(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TargetFailedAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub IBaseEffectOverTime_TargetWasDamaged(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Private Sub SetPosition(ByVal map As Integer, ByVal posX As Integer, ByVal posY As Integer)
    mWorldPosition.map = map
    mWorldPosition.x = posX
    mWorldPosition.y = posY
    Set MapData(mWorldPosition.map, mWorldPosition.x, mWorldPosition.y).Trap = Me
    Call modSendData.SendToAreaByPos(map, posX, posY, PrepareTrapUpdate(1, posX, posY))
End Sub

Private Sub RemoveFromMap()
    Set MapData(mWorldPosition.map, mWorldPosition.x, mWorldPosition.y).Trap = Nothing
    Call modSendData.SendToAreaByPos(mWorldPosition.map, mWorldPosition.x, mWorldPosition.y, PrepareTrapUpdate(0, mWorldPosition.x, mWorldPosition.y))
End Sub

Public Function CanAffectTarget(ByVal TargetIndex, ByVal TargetType As e_ReferenceType) As Boolean
    If mDotInfo.TargetRef.RefType = eUser Then
        If TargetType = eUser Then
            CanAffectTarget = UserMod.CanAttackUser(mDotInfo.TargetRef.ArrayIndex, mDotInfo.TargetRef.VersionId, TargetIndex, UserList(TargetIndex).VersionId) = eCanAttack
            Exit Function
        End If
        If TargetType = eNpc Then
            CanAffectTarget = UserCanAttackNpc(mDotInfo.TargetRef.ArrayIndex, TargetIndex) = eCanAttack
            Exit Function
        End If
    Else
       CanAffectTarget = False
    End If
End Function

Public Function trigger(ByVal TargetIndex, ByVal TargetType As e_ReferenceType)
    trigger = False
    Call modSendData.SendToAreaByPos(mWorldPosition.map, mWorldPosition.x, mWorldPosition.y, PrepareMessagePlayWave(ACTIVATE_TRAOP_WAW_FILE, mWorldPosition.x, mWorldPosition.y, False))
    If EffectOverTime(mDotInfo.EotId).OnHitWav > 0 Then Call modSendData.SendToAreaByPos(mWorldPosition.map, mWorldPosition.x, mWorldPosition.y, PrepareMessagePlayWave(EffectOverTime(mDotInfo.EotId).OnHitWav, mWorldPosition.x, mWorldPosition.y, False))
    If mDotInfo.RemoveEffect Then Exit Function
    mDotInfo.RemoveEffect = True
    Select Case EffectOverTime(mDotInfo.EotId).SubType
        Case e_TrapTypes.eInmovilize
            If mDotInfo.TargetRef.RefType = eUser Then
                If TargetType = eUser Then
                    Call UserMod.Inmovilize(mDotInfo.TargetRef.ArrayIndex, TargetIndex, NET_INMO_DURATION, EffectOverTime(mDotInfo.EotId).OnHitFx)
                    If EffectOverTime(mDotInfo.EotId).OnHitFx > 0 Then Call SendData(SendTarget.ToPCAliveArea, TargetIndex, PrepareMessageCreateFX(UserList(TargetIndex).Char.charindex, EffectOverTime(mDotInfo.EotId).OnHitFx, 0, mWorldPosition.x, mWorldPosition.y))
                    trigger = True
                    Exit Function
                End If
                If TargetType = eNpc Then
                    Call NPCs.Inmovilize(mDotInfo.TargetRef.ArrayIndex, TargetIndex, NET_INMO_DURATION * 10, EffectOverTime(mDotInfo.EotId).OnHitFx)
                    If EffectOverTime(mDotInfo.EotId).OnHitFx > 0 Then Call SendData(SendTarget.ToNPCAliveArea, TargetIndex, PrepareMessageCreateFX(NpcList(TargetIndex).Char.charindex, EffectOverTime(mDotInfo.EotId).OnHitFx, 0, mWorldPosition.x, mWorldPosition.y))
                    trigger = True
                    Exit Function
                End If
            End If
        Case e_TrapTypes.eExplosive
            If mDotInfo.TargetRef.RefType = eUser Then
                If TargetType = eUser Then
                    Call UserMod.DoDamageOrHeal(TargetIndex, mDotInfo.TargetRef.ArrayIndex, eUser, RandomNumber(EffectOverTime(mDotInfo.EotId).TickPowerMin, EffectOverTime(mDotInfo.EotId).TickPowerMax), e_DamageSourceType.e_trap, mDotInfo.EotId)
                    If EffectOverTime(mDotInfo.EotId).OnHitFx > 0 Then Call SendData(SendTarget.ToPCAliveArea, TargetIndex, PrepareMessageCreateFX(UserList(TargetIndex).Char.charindex, EffectOverTime(mDotInfo.EotId).OnHitFx, 0, mWorldPosition.x, mWorldPosition.y))
                    trigger = True
                    Exit Function
                End If
                If TargetType = eNpc Then
                    Call NPCs.DoDamageOrHeal(TargetIndex, mDotInfo.TargetRef.ArrayIndex, eUser, RandomNumber(EffectOverTime(mDotInfo.EotId).TickPowerMin, EffectOverTime(mDotInfo.EotId).TickPowerMax), e_DamageSourceType.e_trap, mDotInfo.EotId)
                    If EffectOverTime(mDotInfo.EotId).OnHitFx > 0 Then Call SendData(SendTarget.ToNPCAliveArea, TargetIndex, PrepareMessageCreateFX(NpcList(TargetIndex).Char.charindex, EffectOverTime(mDotInfo.EotId).OnHitFx, 0, mWorldPosition.x, mWorldPosition.y))
                    trigger = True
                    Exit Function
                End If
            End If
        Case e_TrapTypes.ePoison
            Call EffectsOverTime.CreateEffect(mDotInfo.TargetRef.ArrayIndex, mDotInfo.TargetRef.RefType, TargetIndex, TargetType, EffectOverTime(mDotInfo.EotId).ApplyeffectID)
            If TargetType = eNpc Then
                If EffectOverTime(mDotInfo.EotId).OnHitFx > 0 Then Call SendData(SendTarget.ToNPCAliveArea, TargetIndex, PrepareMessageCreateFX(NpcList(TargetIndex).Char.charindex, EffectOverTime(mDotInfo.EotId).OnHitFx, 0, mWorldPosition.x, mWorldPosition.y))
            Else
                If EffectOverTime(mDotInfo.EotId).OnHitFx > 0 Then Call SendData(SendTarget.ToPCAliveArea, TargetIndex, PrepareMessageCreateFX(UserList(TargetIndex).Char.charindex, EffectOverTime(mDotInfo.EotId).OnHitFx, 0, mWorldPosition.x, mWorldPosition.y))
            End If
            trigger = True
            Exit Function
        End Select
End Function

Public Property Get ElapsedTime() As Long
    ElapsedTime = mElapsedTime
End Property

Public Sub Disable()
    mDotInfo.RemoveEffect = True
End Sub
