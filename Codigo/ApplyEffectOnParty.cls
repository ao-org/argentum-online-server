VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApplyEffectToParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBaseEffectOverTime

Dim TargetTicks As Integer
Dim TickCount As Integer
Dim TickTime As Long
Dim TimeSinceLastTick As Long
Dim TickEffect As Integer
Dim ApplyEffectId As Integer
Dim AreaOfEffect As Integer

Private DotInfo As t_BaseDotInfo
Private Source As t_AnyReference

Public Sub Setup(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
                 ByVal EffectId As Integer, ByVal ID As Long)
    TimeSinceLastTick = 0
    TickCount = 0
    Call SetBaseDot(DotInfo, TargetIndex, TargetType, ID, EffectId)
    Call SetRef(Source, SourceIndex, SourceType)
    Call SetActiveEffect
End Sub

Private Sub SetActiveEffect()
    With EffectOverTime(DotInfo.EotId)
        TargetTicks = .Ticks
        TickTime = .TickTime
        TickEffect = .TickFX
        AreaOfEffect = .Area
        ApplyEffectId = .ApplyEffectId
        If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, .ClientEffectTypeId, DotInfo.UniqueId, .TickTime * TargetTicks, .buffType)
    End With
    Call PerformAction
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.ePartyBonus
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
        Call PerformAction
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal value As Boolean)
    DotInfo.RemoveEffect = value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    TickCount = 0
    IBaseEffectOverTime_Reset = True
    If NewEffectId <> DotInfo.EotId Then
        If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, 0, EffectOverTime(DotInfo.EotId).buffType)
        DotInfo.EotId = NewEffectId
        Call SetActiveEffect
    Else
        IBaseEffectOverTime_Reset = False
        DotInfo.RemoveEffect = True
    End If
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If DotInfo.Removed Then Exit Sub
    DotInfo.Removed = True
    DotInfo.RemoveEffect = True
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
    IBaseEffectOverTime_CallBacksMask = 0
End Property

Public Sub PerformAction()
    If DotInfo.TargetRef.RefType <> eUser Then Exit Sub
    If TickEffect > 0 Then
        If IsVisible(UserList(DotInfo.TargetRef.ArrayIndex)) Then
            Call SendData(SendTarget.ToPCAliveArea, DotInfo.TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(DotInfo.TargetRef.ArrayIndex).Char.charindex, TickEffect, 0, UserList(DotInfo.TargetRef.ArrayIndex).pos.x, UserList(DotInfo.TargetRef.ArrayIndex).pos.y))
        Else
            Call SendData(SendTarget.ToIndex, DotInfo.TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(DotInfo.TargetRef.ArrayIndex).Char.charindex, TickEffect, 0, UserList(DotInfo.TargetRef.ArrayIndex).pos.x, UserList(DotInfo.TargetRef.ArrayIndex).pos.y))
        End If
    End If
    With UserList(DotInfo.TargetRef.ArrayIndex)
        If .Grupo.EnGrupo And IsValidUserRef(.Grupo.Lider) Then
            Dim i As Integer
            With UserList(.Grupo.Lider.ArrayIndex).Grupo
                For i = 1 To .CantidadMiembros
                    If IsValidUserRef(.Miembros(i)) Then
                        If UserList(DotInfo.TargetRef.ArrayIndex).pos.Map = UserList(.Miembros(i).ArrayIndex).pos.Map Then
                            If Distance(UserList(DotInfo.TargetRef.ArrayIndex).pos.x, UserList(DotInfo.TargetRef.ArrayIndex).pos.y, UserList(.Miembros(i).ArrayIndex).pos.x, UserList(.Miembros(i).ArrayIndex).pos.y) < AreaOfEffect Then
                                If CanHelpUser(DotInfo.TargetRef.ArrayIndex, .Miembros(i).ArrayIndex) = eInteractionOk Then
                                    Call ApplyEffectTo(.Miembros(i).ArrayIndex, eUser)
                                End If
                            End If
                        End If
                    End If
                Next i
            End With
        Else
            Call ApplyEffectTo(DotInfo.TargetRef.ArrayIndex, eUser)
        End If
    End With
End Sub

Private Sub ApplyEffectTo(ByVal TargetId As Integer, ByVal TargetType As e_ReferenceType)
    Dim Effect As IBaseEffectOverTime
    If TargetType = eUser Then
        If UserList(TargetId).Stats.MinHp > 0 Then
            Set Effect = FindEffectOnTarget(DotInfo.TargetRef.ArrayIndex, UserList(TargetId).EffectOverTime, ApplyEffectId)
            If Effect Is Nothing Then
                Call CreateEffect(DotInfo.TargetRef.ArrayIndex, DotInfo.TargetRef.RefType, TargetId, TargetType, ApplyEffectId)
            ElseIf EffectOverTime(DotInfo.EotId).Override Then
                Call Effect.Reset(DotInfo.TargetRef.ArrayIndex, DotInfo.TargetRef.RefType, ApplyEffectId)
            End If
        End If
    Else
        Set Effect = FindEffectOnTarget(DotInfo.TargetRef.ArrayIndex, NpcList(TargetId).EffectOverTime, ApplyEffectId)
        If Effect Is Nothing Then
            Call CreateEffect(DotInfo.TargetRef.ArrayIndex, DotInfo.TargetRef.RefType, TargetId, TargetType, ApplyEffectId)
        ElseIf EffectOverTime(DotInfo.EotId).Override Then
            Call Effect.Reset(DotInfo.TargetRef.ArrayIndex, DotInfo.TargetRef.RefType, ApplyEffectId)
        End If
    End If
End Sub

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
End Sub
