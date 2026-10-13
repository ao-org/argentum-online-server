Attribute VB_Name = "UserMod"
Option Explicit
Public Function Inmovilize(ByVal Source As Integer, ByVal Target As Integer, ByVal Duration As Long, ByVal Fx As Integer) As Boolean
    Err.Raise 5, , "Player target must remain unreachable"
End Function
Public Sub DoDamageOrHeal(ByVal Target As Integer, ByVal Source As Integer, ByVal SourceType As e_ReferenceType, ByVal Damage As Long, ByVal DamageType As e_DamageSourceType, ByVal EffectId As Integer, ByVal Message1 As Integer, ByVal Message2 As Integer)
    Err.Raise 5, , "Player target must remain unreachable"
End Sub
