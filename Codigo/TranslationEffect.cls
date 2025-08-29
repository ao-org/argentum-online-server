VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TranslationEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Implements IBaseEffectOverTime

Dim TargetTicks As Integer
Dim TickCount As Integer
Dim TickTime As Long
Dim TimeSinceLastTick As Long

Private DotInfo As t_BaseDotInfo
Private Source As t_AnyReference
Private StartPos As t_WorldPos
Private CurrentStep As Double
Dim Direction As t_Vector

Public Sub Setup(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
    On Error Goto Setup_Err
                 ByVal EffectId As Integer, ByVal ID As Long)
    TimeSinceLastTick = 0
    TickCount = 0
    If TargetType = eNpc Then
        If IsSet(NpcList(TargetIndex).flags.EffectInmunity, e_Inmunities.eTranslation) Then
            DotInfo.Removed = True
            DotInfo.RemoveEffect = True
            Exit Sub
        End If
    End If
    Call SetBaseDot(DotInfo, TargetIndex, TargetType, ID, EffectId)
    Call SetRef(Source, SourceIndex, SourceType)
    With EffectOverTime(EffectId)
        TargetTicks = .Ticks - 1 ' we perform the first translation right away
        TickTime = .TickTime
        'to prevent user throwing creatures to other players to kill them lets add a provoke effect on hit so it target the caster
        If TargetType = eNpc And .ApplyEffectId > 0 Then
            Dim Effect As IBaseEffectOverTime
            Set Effect = FindEffectOnTarget(0, NpcList(TargetIndex).EffectOverTime, .ApplyEffectId)
            If Effect Is Nothing Then Call CreateEffect(SourceIndex, SourceType, TargetIndex, eNpc, .ApplyEffectId)
        End If
    End With
    CurrentStep = 0
    StartPos = GetPosition(DotInfo.TargetRef)
    Dim CasterPos As t_WorldPos
    CasterPos = GetPosition(Source)
    Call SetTranslationState(DotInfo.TargetRef, True)
    Direction.x = (StartPos.x - CasterPos.x)
    Direction.y = (StartPos.y - CasterPos.y)
    Direction = GetNormal(Direction)
    Call PerformAction
    Exit Sub
Setup_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.Setup", Erl)
End Sub

Public Property Get IBaseEffectOverTime_TypeId() As Integer
    IBaseEffectOverTime_TypeId = e_EffectOverTimeType.eTranslation
End Property

Public Property Get IBaseEffectOverTime_SharedTypeId() As Integer
    IBaseEffectOverTime_SharedTypeId = EffectOverTime(DotInfo.EotId).SharedTypeId
End Property

Public Property Get IBaseEffectOverTime_UniqueId() As Long
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
    On Error Goto IBaseEffectOverTime_Update_Err
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
    Exit Sub
IBaseEffectOverTime_Update_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.IBaseEffectOverTime_Update", Erl)
End Sub

Public Property Get IBaseEffectOverTime_RemoveMe() As Boolean
    IBaseEffectOverTime_RemoveMe = DotInfo.RemoveEffect
End Property

Public Property Let IBaseEffectOverTime_RemoveMe(ByVal value As Boolean)
    DotInfo.RemoveEffect = value
End Property

Public Function IBaseEffectOverTime_Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
    On Error Goto IBaseEffectOverTime_Reset_Err
    TickCount = 0
    IBaseEffectOverTime_Reset = True
    Exit Function
IBaseEffectOverTime_Reset_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.IBaseEffectOverTime_Reset", Erl)
End Function

Public Property Get IBaseEffectOverTime_EotId() As Integer
    IBaseEffectOverTime_EotId = DotInfo.EotId
End Property

Public Sub IBaseEffectOverTime_OnRemove()
    On Error Goto IBaseEffectOverTime_OnRemove_Err
    If DotInfo.Removed Then Exit Sub
    DotInfo.Removed = True
    DotInfo.RemoveEffect = True
    Call SetTranslationState(DotInfo.TargetRef, False)
    Exit Sub
IBaseEffectOverTime_OnRemove_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.IBaseEffectOverTime_OnRemove", Erl)
End Sub

Public Sub IBaseEffectOverTime_TargetUseMagic(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
    On Error Goto IBaseEffectOverTime_TargetUseMagic_Err
    Exit Sub
IBaseEffectOverTime_TargetUseMagic_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.IBaseEffectOverTime_TargetUseMagic", Erl)
End Sub

Public Sub IBaseEffectOverTime_TartgetWillAtack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TartgetWillAtack_Err
    Exit Sub
IBaseEffectOverTime_TartgetWillAtack_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.IBaseEffectOverTime_TartgetWillAtack", Erl)
End Sub

Public Sub IBaseEffectOverTime_TartgetDidHit(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TartgetDidHit_Err
    Exit Sub
IBaseEffectOverTime_TartgetDidHit_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.IBaseEffectOverTime_TartgetDidHit", Erl)
End Sub

Public Sub IBaseEffectOverTime_TargetFailedAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TargetFailedAttack_Err
    Exit Sub
IBaseEffectOverTime_TargetFailedAttack_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.IBaseEffectOverTime_TargetFailedAttack", Erl)
End Sub

Public Sub IBaseEffectOverTime_TargetWasDamaged(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    On Error Goto IBaseEffectOverTime_TargetWasDamaged_Err
    Exit Sub
IBaseEffectOverTime_TargetWasDamaged_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.IBaseEffectOverTime_TargetWasDamaged", Erl)
End Sub

Public Function IBaseEffectOverTime_ApplyDamageReduction(ByVal Damage As Long, ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType) As Long
    On Error Goto IBaseEffectOverTime_ApplyDamageReduction_Err
    IBaseEffectOverTime_ApplyDamageReduction = Damage
    Exit Function
IBaseEffectOverTime_ApplyDamageReduction_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.IBaseEffectOverTime_ApplyDamageReduction", Erl)
End Function

Public Sub IBaseEffectOverTime_TargetWillAttackPosition(ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
    On Error Goto IBaseEffectOverTime_TargetWillAttackPosition_Err
    Exit Sub
IBaseEffectOverTime_TargetWillAttackPosition_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.IBaseEffectOverTime_TargetWillAttackPosition", Erl)
End Sub

Public Function IBaseEffectOverTime_ChangeTarget(ByVal NewTargetIndex As Integer, ByVal NewTargetType As e_ReferenceType) As Boolean
    On Error Goto IBaseEffectOverTime_ChangeTarget_Err
    IBaseEffectOverTime_ChangeTarget = False
    Exit Function
IBaseEffectOverTime_ChangeTarget_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.IBaseEffectOverTime_ChangeTarget", Erl)
End Function

Public Property Get IBaseEffectOverTime_CallBacksMask() As Long
    IBaseEffectOverTime_CallBacksMask = 0
End Property

Public Property Get IBaseEffectOverTime_KeepAfterDead() As Boolean
    IBaseEffectOverTime_KeepAfterDead = False
End Property

Public Sub IBaseEffectOverTime_TargetChangeTerrain()
    On Error Goto IBaseEffectOverTime_TargetChangeTerrain_Err
    Exit Sub
IBaseEffectOverTime_TargetChangeTerrain_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.IBaseEffectOverTime_TargetChangeTerrain", Erl)
End Sub

Public Function IsValidPos(ByVal map As Integer, ByVal TileX As Integer, ByVal TileY As Integer) As Boolean
    On Error Goto IsValidPos_Err
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
    Exit Function
IsValidPos_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.IsValidPos", Erl)
End Function

Private Sub PerformAction()
    On Error Goto PerformAction_Err
    Dim currentPos As t_WorldPos
    Dim TargetPoint As t_Vector
    Dim TargetTranslation As t_Vector
    Dim TargetPos As t_WorldPos
    Dim TranslationDiff As Double
    Dim Tanslation As Integer
    CurrentStep = CurrentStep + 1
    currentPos = GetPosition(DotInfo.TargetRef)
    TargetPoint.x = Direction.x * (CurrentStep) + StartPos.x
    TargetPoint.y = Direction.y * (CurrentStep) + StartPos.y
    TargetTranslation.x = TargetPoint.x - currentPos.x
    TargetTranslation.y = TargetPoint.y - currentPos.y
    TranslationDiff = Abs(TargetTranslation.x) - Abs(TargetTranslation.y)
    TargetPos.map = currentPos.map
    If Abs(TranslationDiff) < 0.3 Then 'if they are similar we are close to 45% let move in both directions
        TargetPos.x = currentPos.x + Sgn(TargetTranslation.x)
        TargetPos.y = currentPos.y + Sgn(TargetTranslation.y)
    ElseIf TranslationDiff > 0 Then 'x axis is bigger than
        TargetPos.x = currentPos.x + Sgn(TargetTranslation.x)
        TargetPos.y = currentPos.y
    Else
        TargetPos.x = currentPos.x
        TargetPos.y = currentPos.y + Sgn(TargetTranslation.y)
    End If
    If IsValidPos(TargetPos.map, TargetPos.x, TargetPos.y) Then
        Call TranslateTo(TargetPos)
    Else
        DotInfo.RemoveEffect = True
    End If
    Exit Sub
PerformAction_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.PerformAction", Erl)
End Sub

Private Sub TranslateTo(ByRef TargetPos As t_WorldPos)
    On Error Goto TranslateTo_Err
    If DotInfo.TargetRef.RefType = eUser Then
        Call TranslateUserPos(DotInfo.TargetRef.ArrayIndex, TargetPos, TickTime)
    Else
        Call TranslateNpcChar(DotInfo.TargetRef.ArrayIndex, TargetPos, TickTime)
    End If
    Dim currentPos As t_WorldPos
    currentPos = GetPosition(DotInfo.TargetRef)
    If Not MapData(currentPos.map, currentPos.x, currentPos.y).Trap Is Nothing Then
         Call ModMap.ActivateTrap(DotInfo.TargetRef.ArrayIndex, DotInfo.TargetRef.RefType, currentPos.map, currentPos.x, currentPos.y)
    End If
    Exit Sub
TranslateTo_Err:
    Call TraceError(Err.Number, Err.Description, "TranslationEffect.TranslateTo", Erl)
End Sub




