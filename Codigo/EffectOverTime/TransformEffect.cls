VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TransformEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBaseEffectOverTime


Dim DotHpMin As Integer
Dim DotHpMax As Integer
Dim TargetTicks As Integer
Dim TickCount As Integer
Dim TickSumption As Integer
Dim TickTime As Long
Dim TimeSinceLastTick As Long
Dim TickEffect As Integer
Dim Modifiers As Integer
Dim EffectBonus As Single
Dim npcId As Integer
Dim StatusMask As Long

Private DotInfo As t_BaseDotInfo
Private Source As t_AnyReference

Public Sub Setup(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
                 ByVal EffectId As Integer, ByVal ID As Long)
    TimeSinceLastTick = 0
    Call SetBaseDot(DotInfo, TargetIndex, TargetType, ID, EffectId)
    Call SetRef(Source, SourceIndex, SourceType)
    Call SetEffect(EffectId)
End Sub

Public Sub SetEffect(ByVal EffectId As Integer)
    DotInfo.EotId = EffectId
    TickCount = 0
    With EffectOverTime(EffectId)
        TickTime = .TickTime
        TickEffect = .TickFX
        TargetTicks = .Ticks
        TickSumption = .TickSumption
        npcId = .npcId
        StatusMask = .ApplyStatusMask
        If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
            Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, .ClientEffectTypeId, DotInfo.UniqueId, .TickTime * TargetTicks, .buffType)
        End If
    End With
    Call TransformChacter
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eTransform
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
    IBaseEffectOverTime_EffectType = EffectOverTime(DotInfo.EotId).buffType
End Property

Public Sub IBaseEffectOverTime_Update(ByVal deltaTime As Long)
    If TickCount >= TargetTicks Or Not IsValidRef(DotInfo.TargetRef) Then
        DotInfo.RemoveEffect = True
        Exit Sub
    End If
    TimeSinceLastTick = TimeSinceLastTick + deltaTime
    If TimeSinceLastTick >= TickTime Then
        TimeSinceLastTick = 0
        TickCount = TickCount + 1
        If DotInfo.TargetRef.RefType = eUser Then
            If UserMod.ModifyMana(DotInfo.TargetRef.ArrayIndex, -TickSumption, True) Then
                DotInfo.RemoveEffect = True
            End If
        End If
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal Value As Boolean)
    DotInfo.RemoveEffect = Value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    If (NewEffectId = DotInfo.EotId) Then
        IBaseEffectOverTime_Reset = False
        DotInfo.RemoveEffect = True
        Exit Function
    End If
    TimeSinceLastTick = 0
    TickCount = 0
    IBaseEffectOverTime_Reset = True
    Call ReturnCharToNormal
    Call SetEffect(NewEffectId)
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then
        Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, _
                                    DotInfo.UniqueId, (TickTime * TargetTicks) - TimeSinceLastTick, EffectOverTime(DotInfo.EotId).buffType)
    End If
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If DotInfo.Removed Then Exit Sub
    DotInfo.Removed = True
    DotInfo.RemoveEffect = True
    Call ReturnCharToNormal
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, 0, EffectOverTime(DotInfo.EotId).buffType)
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

Public Function IBaseEffectOverTime_ApplyDamageReduction(ByVal Damage As Long, ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType) As Long
    IBaseEffectOverTime_ApplyDamageReduction = Damage
End Function

Public Sub IBaseEffectOverTime_TargetWillAttackPosition(ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
End Sub

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = eTargetChangeTerrain
End Property

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
    If DotInfo.TargetRef.RefType <> eUser Then Exit Sub
    With UserList(DotInfo.TargetRef.ArrayIndex)
        If .flags.Navegando > 0 Or .flags.Montado > 0 Or .flags.Nadando > 0 Then
            DotInfo.RemoveEffect = True
        End If
    End With
End Sub

Private Sub TransformChacter()
    Call EffectsOverTime.ApplyEotModifier(DotInfo.TargetRef, EffectOverTime(DotInfo.EotId))
    If (DotInfo.TargetRef.RefType = eUser) Then
        With UserList(DotInfo.TargetRef.ArrayIndex)
            Dim NpcIndex As String
            NpcIndex = "NPC" & npcId
136         .Char.body = val(LeerNPCs.GetValue(NpcIndex, "Body"))
138         .Char.head = val(LeerNPCs.GetValue(NpcIndex, "Head"))
140         .Char.Heading = val(LeerNPCs.GetValue(NpcIndex, "Heading"))
142         .Char.BodyIdle = val(LeerNPCs.GetValue(NpcIndex, "BodyIdle"))
143         .Char.Ataque1 = val(LeerNPCs.GetValue(NpcIndex, "Ataque1"))
156         .Char.WeaponAnim = val(LeerNPCs.GetValue(NpcIndex, "Arma"))
158         .Char.ShieldAnim = val(LeerNPCs.GetValue(NpcIndex, "Escudo"))
160         .Char.CascoAnim = val(LeerNPCs.GetValue(NpcIndex, "Casco"))
161         .Char.CartAnim = val(LeerNPCs.GetValue(NpcIndex, "Cart"))
            .Char.CastAnimation = val(LeerNPCs.GetValue(NpcIndex, "CastAnimation"))
            Dim CantidadAnimaciones As Integer
144         CantidadAnimaciones = val(LeerNPCs.GetValue(NpcIndex, "Animaciones"))
128         Call ChangeUserChar(DotInfo.TargetRef.ArrayIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)
            Call SetMask(.flags.StatusMask, StatusMask)
130         Call RefreshCharStatus(DotInfo.TargetRef.ArrayIndex)
        End With
    End If
End Sub

Private Sub ReturnCharToNormal()
    Call EffectsOverTime.RemoveEotModifier(DotInfo.TargetRef, EffectOverTime(DotInfo.EotId))
    If DotInfo.TargetRef.RefType = eUser Then
        Call UnsetMask(UserList(DotInfo.TargetRef.ArrayIndex).flags.StatusMask, StatusMask)
        Call UpdateCharWithEquipedItems(DotInfo.TargetRef.ArrayIndex)
    End If
End Sub


