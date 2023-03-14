VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EffectProvoke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBaseEffectOverTime

Dim TickTime As Long
Dim TimeSinceLastTick As Long
Private DotInfo As t_BaseDotInfo
Private Source As t_AnyReference
Private PhysicalDamageModfier As Single
Private MagicDamageModifier As Single
Private ReductionActive As Boolean
Const TaunteffectId As Integer = 4


Public Sub Setup(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
                 ByVal EffectId As Integer, ByVal ID As Integer)
    Call SetRef(DotInfo.TargetRef, TargetIndex, TargetType)
    Call SetRef(Source, SourceIndex, SourceType)
    DotInfo.UniqueId = ID
    DotInfo.RemoveEffect = False
    DotInfo.EotId = EffectId
    TimeSinceLastTick = 0
    ReductionActive = False
    If SourceType = eUser Then
        Call UserMod.RemoveInvisibility(SourceIndex)
        Call SetMask(UserList(SourceIndex).flags.StatusMask, eTaunting)
    Else
        Call SetMask(NpcList(SourceIndex).flags.StatusMask, eTaunting)
    End If
    With EffectOverTime(EffectId)
        TickTime = .TickTime
        If TargetType = eUser Then
            If .ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(TargetIndex, .ClientEffectTypeId, ID, .TickTime, .BuffType)
        End If
        Call ApplyEffectToArea(.Area)
        UserList(SourceIndex).Char.Otra_Aura = .Aura
    End With
    Call SendData(SendTarget.ToPCAliveArea, SourceIndex, PrepareMessageAuraToChar(UserList(SourceIndex).Char.charindex, UserList(SourceIndex).Char.Otra_Aura, False, 5))
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eProvoke
End Property

Public Property Get IBaseEffectOverTime_UniqueId() As Integer
    IBaseEffectOverTime_UniqueId = DotInfo.UniqueId
End Property

Public Property Get IBaseEffectOverTime_CasterIsValid() As Boolean
    IBaseEffectOverTime_CasterIsValid = IsValidRef(Source)
End Property

Public Property Get IBaseEffectOverTime_CasterArrayIndex() As Integer
    IBaseEffectOverTime_CasterArrayIndex = Source.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_CasterRefType() As e_ReferenceType
    IBaseEffectOverTime_CasterRefType = Source.RefType
End Property

Public Property Get IBaseEffectOverTime_CasterUserId() As Long
    IBaseEffectOverTime_CasterUserId = Source.userID
End Property

Public Property Get IBaseEffectOverTime_TargetIsValid() As Boolean
    IBaseEffectOverTime_TargetIsValid = IsValidRef(DotInfo.TargetRef)
End Property

Public Property Get IBaseEffectOverTime_TargetArrayIndex() As Integer
    IBaseEffectOverTime_TargetArrayIndex = DotInfo.TargetRef.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_TargetRefType() As e_ReferenceType
    IBaseEffectOverTime_TargetRefType = DotInfo.TargetRef.RefType
End Property

Public Property Get IBaseEffectOverTime_TargetUserId() As Long
    IBaseEffectOverTime_TargetUserId = DotInfo.TargetRef.userID
End Property

Public Property Get IBaseEffectOverTime_EffectType() As e_EffectType
    IBaseEffectOverTime_EffectType = EffectOverTime(DotInfo.EotId).BuffType
End Property

Public Sub IBaseEffectOverTime_Update(ByVal deltaTime As Long)
    TimeSinceLastTick = TimeSinceLastTick + deltaTime
    If TimeSinceLastTick >= TickTime Or Not IsValidRef(DotInfo.TargetRef) Then
        DotInfo.RemoveEffect = True
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal value As Boolean)
    DotInfo.RemoveEffect = value
End Property

Public Sub IBaseEffectOverTime_Reset()
End Sub

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If Source.RefType = eUser Then
        Call UnsetMask(UserList(Source.ArrayIndex).flags.StatusMask, eTaunting)
    Else
        Call UnsetMask(NpcList(Source.ArrayIndex).flags.StatusMask, eTaunting)
    End If
    UserList(Source.ArrayIndex).Char.Otra_Aura = ""
    Call SendData(SendTarget.ToPCAliveArea, Source.ArrayIndex, PrepareMessageAuraToChar(UserList(Source.ArrayIndex).Char.charindex, UserList(Source.ArrayIndex).Char.Body_Aura, True, 5))
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


Private Sub ApplyEffectToArea(ByVal Area As Integer)
    Debug.Assert Area > 0
    Dim x, y, w, h, map As Integer
    If Source.RefType = eUser Then
        x = UserList(Source.ArrayIndex).pos.x - Area / 2
        y = UserList(Source.ArrayIndex).pos.y - Area / 2
        map = UserList(Source.ArrayIndex).pos.map
    Else
        x = NpcList(Source.ArrayIndex).pos.x - Area / 2
        y = NpcList(Source.ArrayIndex).pos.y - Area / 2
        map = NpcList(Source.ArrayIndex).pos.map
    End If
    w = x + Area
    h = y + Area
    x = max(1, x)
    y = max(1, y)
    w = Min(w, 99)
    h = Min(h, 99)
    Dim i, j As Integer
    Dim Effect As IBaseEffectOverTime
    For i = y To h
        For j = x To w
            'we can't taunt someone already taunted
            If MapData(map, j, i).UserIndex > 0 Then
                If CanAttackUser(Source.ArrayIndex, Source.VersionId, MapData(map, j, i).UserIndex, UserList(MapData(map, j, i).UserIndex).VersionId) = eCanAttack Then
                    Set Effect = FindEffectOnTarget(0, UserList(MapData(map, j, i).UserIndex).EffectOverTime, TaunteffectId)
                    If Effect Is Nothing Then Call CreateEffect(Source.ArrayIndex, Source.RefType, MapData(map, j, i).UserIndex, eUser, TaunteffectId)
                End If
            ElseIf MapData(map, j, i).npcIndex > 0 Then
                If UserCanAttackNpc(Source.ArrayIndex, MapData(map, j, i).npcIndex) = eCanAttack Then
                    Set Effect = FindEffectOnTarget(0, NpcList(MapData(map, j, i).npcIndex).EffectOverTime, TaunteffectId)
                    If Effect Is Nothing Then Call CreateEffect(Source.ArrayIndex, Source.RefType, MapData(map, j, i).npcIndex, eNpc, TaunteffectId)
                End If
            End If
        Next j
    Next i
End Sub
