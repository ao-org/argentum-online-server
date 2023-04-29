VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DelayedBlast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBaseEffectOverTime

Dim TickTime As Long
Dim TimeSinceLastTick As Long
Private DotInfo As t_BaseDotInfo
Private TargetX As Integer
Private TargetY As Integer
Private TargetMap As Integer
Private SourceObjIndex As Integer

Public Sub Setup(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal EffectId As Integer, ByVal ID As Integer, _
                 ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer, ByVal SourceObj As Integer)
    Call SetBaseDot(DotInfo, SourceIndex, SourceType, ID, EffectId)
    TimeSinceLastTick = 0
    TargetX = TileX
    TargetY = TileY
    TargetMap = Map
    Dim Distance As t_Vector
    Dim CurrPos As t_WorldPos
    CurrPos = GetPosition(DotInfo.TargetRef)
    SourceObjIndex = SourceObj
    With EffectOverTime(EffectId)
        TickTime = .TickTime * Matematicas.Distance(CurrPos.x, CurrPos.y, TargetX, TargetY)
    End With
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eDelayedBlast
End Property

Public Property Get IBaseEffectOverTime_UniqueId() As Integer
    IBaseEffectOverTime_UniqueId = DotInfo.UniqueId
End Property

Public Property Get IBaseEffectOverTime_CasterIsValid() As Boolean
    IBaseEffectOverTime_CasterIsValid = IsValidRef(DotInfo.TargetRef)
End Property

Public Property Get IBaseEffectOverTime_CasterArrayIndex() As Integer
    IBaseEffectOverTime_CasterArrayIndex = DotInfo.TargetRef.ArrayIndex
End Property

Public Property Get IBaseEffectOverTime_CasterRefType() As e_ReferenceType
    IBaseEffectOverTime_CasterRefType = DotInfo.TargetRef.RefType
End Property

Public Property Get IBaseEffectOverTime_CasterUserId() As Long
    IBaseEffectOverTime_CasterUserId = DotInfo.TargetRef.userID
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
    TimeSinceLastTick = TimeSinceLastTick + deltaTime
    If TimeSinceLastTick >= TickTime Or Not IsValidRef(DotInfo.TargetRef) Then
        DotInfo.RemoveEffect = True
    End If
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal Value As Boolean)
    DotInfo.RemoveEffect = Value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    IBaseEffectOverTime_Reset = True
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If DotInfo.Removed Then Exit Sub
    DotInfo.Removed = True
    DotInfo.RemoveEffect = True
    Call DoBlast
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

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
End Sub

Private Sub DoBlast()
    Dim Area As Integer
    Area = EffectOverTime(DotInfo.EotId).Area
    Debug.Assert Area > 0
    If DotInfo.TargetRef.RefType = eUser Then
        If EffectOverTime(DotInfo.EotId).OnHitFx > 0 Then Call SendData(SendTarget.ToPCAliveArea, DotInfo.TargetRef.ArrayIndex, PrepareMessageFxPiso(EffectOverTime(DotInfo.EotId).OnHitFx, TargetX, TargetY))
        If EffectOverTime(DotInfo.EotId).OnHitWav > 0 Then Call modSendData.SendToAreaByPos(TargetMap, TargetX, TargetY, PrepareMessagePlayWave(EffectOverTime(DotInfo.EotId).OnHitWav, TargetX, TargetY, False))
    Else
        If EffectOverTime(DotInfo.EotId).OnHitFx > 0 Then Call SendData(SendTarget.ToNPCAliveArea, DotInfo.TargetRef.ArrayIndex, PrepareMessageFxPiso(EffectOverTime(DotInfo.EotId).OnHitFx, TargetX, TargetY))
        If EffectOverTime(DotInfo.EotId).OnHitWav > 0 Then Call modSendData.SendToAreaByPos(TargetMap, TargetX, TargetY, PrepareMessagePlayWave(EffectOverTime(DotInfo.EotId).OnHitWav, TargetX, TargetY, False))
    End If
    Dim x, y, w, h, Map As Integer
    x = TargetX - Area \ 2
    y = TargetY - Area \ 2
    w = x + (Area - 1)
    h = y + (Area - 1)
    x = max(1, x)
    y = max(1, y)
    w = Min(w, 99)
    h = Min(h, 99)
    Dim i, j As Integer
    Dim Effect As IBaseEffectOverTime
    For i = y To h
        For j = x To w
            'we can't taunt someone already taunted
            If MapData(TargetMap, j, i).UserIndex > 0 Then
                If DotInfo.TargetRef.RefType = eUser Then
                    If UserMod.CanAttackUser(DotInfo.TargetRef.ArrayIndex, DotInfo.TargetRef.VersionId, MapData(TargetMap, j, i).UserIndex, UserList(MapData(TargetMap, j, i).UserIndex).VersionId) = eCanAttack Then
                        Call UserDoDamageToUser(DotInfo.TargetRef.ArrayIndex, MapData(TargetMap, j, i).UserIndex, GetDamageForTile(j, i, GetUserDamageWithItem(DotInfo.TargetRef.ArrayIndex, SourceObjIndex, 0)), e_phisical, SourceObjIndex)
                    End If
                Else
                    If NPCs.CanAttackUser(DotInfo.TargetRef.ArrayIndex, MapData(TargetMap, j, i).UserIndex) = eCanAttack Then
                        Call NpcDoDamageToUser(DotInfo.TargetRef.ArrayIndex, MapData(TargetMap, j, i).UserIndex, GetDamageForTile(j, i, GetNpcDamage(DotInfo.TargetRef.ArrayIndex)), e_phisical, SourceObjIndex)
                    End If
                End If
            ElseIf MapData(TargetMap, j, i).npcIndex > 0 Then
                If DotInfo.TargetRef.RefType = eUser Then
                    If UserCanAttackNpc(DotInfo.TargetRef.ArrayIndex, MapData(TargetMap, j, i).NpcIndex) = eCanAttack Then
                        Call UserDamageToNpc(DotInfo.TargetRef.ArrayIndex, MapData(TargetMap, j, i).npcIndex, GetDamageForTile(j, i, GetUserDamageWithItem(DotInfo.TargetRef.ArrayIndex, SourceObjIndex, 0)), e_phisical, SourceObjIndex)
                    End If
                Else
                    If NPCs.CanAttackNpc(DotInfo.TargetRef.ArrayIndex, MapData(TargetMap, j, i).NpcIndex) = eCanAttack Then
                        Call NpcDamageToNpc(DotInfo.TargetRef.ArrayIndex, MapData(TargetMap, j, i).npcIndex, GetDamageForTile(j, i, GetNpcDamage(DotInfo.TargetRef.ArrayIndex)))
                    End If
                End If
            End If
            'add dama to itself (only for user as ia can't avoid this)
            If DotInfo.TargetRef.RefType = eUser Then
                If MapData(TargetMap, j, i).UserIndex = DotInfo.TargetRef.ArrayIndex Then
                    Call UserDoDamageToUser(DotInfo.TargetRef.ArrayIndex, MapData(TargetMap, j, i).UserIndex, GetDamageForTile(j, i, GetUserDamageWithItem(DotInfo.TargetRef.ArrayIndex, SourceObjIndex, 0)), e_phisical, SourceObjIndex)
                End If
            End If
            
        Next j
    Next i
End Sub

Private Function GetDamageForTile(ByVal TileX As Integer, ByVal TileY As Integer, ByVal BaseDamage As Long) As Integer
    Dim Distance As Double
    Dim BlastArea As Double
    Dim Percent As Double
    BlastArea = max(EffectOverTime(DotInfo.EotId).Area / 2, 1) * 2
    Distance = Matematicas.Distance(TileX, TileY, TargetX, TargetY)
    Percent = (BlastArea - Distance) / BlastArea
    GetDamageForTile = max(0, BaseDamage * Percent)
End Function


