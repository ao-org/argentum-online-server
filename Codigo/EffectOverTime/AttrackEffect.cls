VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AttrackEffect"
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
Dim SpeedModifier As Single

Private DotInfo As t_BaseDotInfo
Private Source As t_AnyReference

Public Sub Setup(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
                 ByVal EffectId As Integer, ByVal ID As Long)
    TimeSinceLastTick = 0
    TickCount = 0
    Call SetBaseDot(DotInfo, TargetIndex, TargetType, ID, EffectId)
    Call SetRef(Source, SourceIndex, SourceType)
    With EffectOverTime(EffectId)
        TargetTicks = .Ticks
        TickTime = .TickTime
        TickEffect = .TickFX
        SpeedModifier = .SpeedModifier
        If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(TargetIndex, .ClientEffectTypeId, ID, .TickTime * TargetTicks, e_EffectType.eInformativeDebuff)
        If Source.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(TargetIndex, .ClientEffectTypeId, ID, .TickTime * TargetTicks, e_EffectType.eInformativeBuff)
    End With
    Call EffectsOverTime.ApplyEotModifier(DotInfo.TargetRef, EffectOverTime(EffectId))
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.ePullTarget
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

Public Sub IBaseEffectOverTime_TargetWillAttackPosition(ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
End Sub

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = 0
End Property

Public Sub IBaseEffectOverTime_Update(ByVal deltaTime As Long)
    If TickCount >= TargetTicks Or Not IsValidRef(DotInfo.TargetRef) Or Not IsValidRef(Source) Then
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

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal Value As Boolean)
    DotInfo.RemoveEffect = Value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    TickCount = 0
    IBaseEffectOverTime_Reset = True
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    If DotInfo.Removed Then Exit Sub
    DotInfo.Removed = True
    DotInfo.RemoveEffect = True
    Call EffectsOverTime.RemoveEotModifier(DotInfo.TargetRef, EffectOverTime(DotInfo.EotId))
    If DotInfo.TargetRef.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(DotInfo.TargetRef.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, 0, e_EffectType.eInformativeDebuff)
    If Source.RefType = eUser And EffectOverTime(DotInfo.EotId).ClientEffectTypeId > 0 Then Call WriteSendSkillCdUpdate(Source.ArrayIndex, EffectOverTime(DotInfo.EotId).ClientEffectTypeId, DotInfo.UniqueId, 0, e_EffectType.eInformativeBuff)
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

Public Function IsValidPos(ByVal map As Integer, ByVal TileX As Integer, ByVal TileY As Integer) As Boolean
    IsValidPos = False
    Dim CanGoWater As Boolean
    Dim Mounted As Boolean
    If DotInfo.TargetRef.RefType = eUser Then
        CanGoWater = UserList(DotInfo.TargetRef.ArrayIndex).flags.Navegando
        Mounted = UserList(DotInfo.TargetRef.ArrayIndex).flags.Montado
    Else
        CanGoWater = NpcList(DotInfo.TargetRef.ArrayIndex).flags.AguaValida
    End If
    IsValidPos = LegalPos(map, TileX, TileY, CanGoWater, Not CanGoWater, Mounted, False, False)
End Function

Public Sub PerformAction()
    Dim Direction As t_Vector
    Dim TargetPos As t_WorldPos
    Dim CurrPos As t_WorldPos
    Dim TargetPoint As t_WorldPos
    TargetPos = GetPosition(Source)
    CurrPos = GetPosition(DotInfo.TargetRef)
    Direction.x = TargetPos.x - CurrPos.x
    Direction.y = TargetPos.y - CurrPos.y
    Dim TranslationDiff As Double
    TranslationDiff = Abs(Direction.x) - Abs(Direction.y)
    TargetPos.map = CurrPos.map
    If Abs(TranslationDiff) < 0.3 Then 'if they are similar we are close to 45% let move in both directions
        TargetPos.x = CurrPos.x + Sgn(Direction.x)
        TargetPos.y = CurrPos.y + Sgn(Direction.y)
    ElseIf TranslationDiff > 0 Then 'x axis is bigger than
        TargetPos.x = CurrPos.x + Sgn(Direction.x)
        TargetPos.y = CurrPos.y
    Else
        TargetPos.x = CurrPos.x
        TargetPos.y = CurrPos.y + Sgn(Direction.y)
    End If
    If IsValidPos(TargetPos.map, TargetPos.x, TargetPos.y) Then
        Call TranslateTo(TargetPos)
    End If
End Sub

Private Sub TranslateTo(ByRef TargetPos As t_WorldPos)
    If DotInfo.TargetRef.RefType = eUser Then
        Call TranslateUserPos(DotInfo.TargetRef.ArrayIndex, TargetPos, 200)
    Else
        Call TranslateNpcChar(DotInfo.TargetRef.ArrayIndex, TargetPos, 200)
    End If
    Dim currentPos As t_WorldPos
    currentPos = GetPosition(DotInfo.TargetRef)
    If Not MapData(currentPos.map, currentPos.x, currentPos.y).Trap Is Nothing Then
         Call ModMap.ActivateTrap(DotInfo.TargetRef.ArrayIndex, DotInfo.TargetRef.RefType, currentPos.map, currentPos.x, currentPos.y)
    End If
End Sub

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
End Sub
