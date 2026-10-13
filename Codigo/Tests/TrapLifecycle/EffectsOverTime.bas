Attribute VB_Name = "EffectsOverTime"
' Harness-only boundaries. Production routines are appended verbatim by the runner.
Option Explicit
Private UniqueIdCounter As Long
Private ActiveEffects As t_EffectOverTimeList
Public Const ACTIVE_EFFECT_LIST_SIZE As Integer = 10

Public Sub ResetEffects()
    Erase ActiveEffects.EffectList
    ActiveEffects.EffectCount = 0
    UniqueIdCounter = 0
End Sub
Public Function NextIdAfter(ByVal Previous As Long) As Long
    UniqueIdCounter = Previous
    NextIdAfter = GetNextId()
End Function
Private Function GetEOT(ByVal Kind As Integer) As IBaseEffectOverTime
    Set GetEOT = New clsTrap
End Function
Private Sub RecycleEffect(ByRef Effect As IBaseEffectOverTime)
    Recycled = Recycled + 1
End Sub
Public Sub Tick(ByVal Elapsed As Long)
    Dim i As Integer
    Do While i < ActiveEffects.EffectCount
        If UpdateEffect(i, Elapsed) Then i = i + 1
    Loop
End Sub
Public Sub CreateEffect(ByVal Source As Integer, ByVal SourceType As e_ReferenceType, ByVal Target As Integer, ByVal TargetType As e_ReferenceType, ByVal EffectId As Integer)
    Call Expect(EffectId = 8, "poison effect id")
    PoisonCalls = PoisonCalls + 1
End Sub
