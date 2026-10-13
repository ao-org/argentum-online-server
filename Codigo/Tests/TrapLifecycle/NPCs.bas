Attribute VB_Name = "NPCs"
' Spy effect boundaries: assert exactly one existing effect invocation per trigger.
Option Explicit
Public Sub Paralice(ByVal Source As Integer, ByVal Target As Integer, ByVal Duration As Double, ByVal Fx As Integer)
    Call Expect(Duration = 13, "unchanged configured paralysis multiplier")
    ParalysisCalls = ParalysisCalls + 1
End Sub
Public Sub DoDamageOrHeal(ByVal Target As Integer, ByVal Source As Integer, ByVal SourceType As e_ReferenceType, ByVal Damage As Long, ByVal DamageType As e_DamageSourceType, ByVal EffectId As Integer)
    Call Expect(Damage = -100, "configured damage range passed through")
    DamageCalls = DamageCalls + 1
End Sub
