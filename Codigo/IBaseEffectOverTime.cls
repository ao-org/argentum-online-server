VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBaseEffectOverTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get CasterIsValid() As Boolean
End Property

Public Property Get CasterArrayIndex() As Integer
End Property

Public Property Get CasterRefType() As e_ReferenceType
End Property

Public Property Get CasterUserId() As Long
End Property

Public Property Get TargetIsValid() As Boolean
End Property

Public Property Get TargetArrayIndex() As Integer
End Property

Public Property Get TargetRefType() As e_ReferenceType
End Property

Public Property Get TargetUserId() As Long
End Property

Public Property Get UniqueId() As Integer
End Property

Public Property Get TypeId() As Integer
End Property

Public Property Get EotId() As Integer
End Property

Public Sub Update(ByVal deltaTime As Long)
End Sub

Public Property Get EffectType() As e_EffectType
End Property

Public Property Get RemoveMe() As Boolean
End Property

Public Property Let RemoveMe(ByVal value As Boolean)
End Property

Public Sub Reset()
End Sub

Public Sub OnRemove()
End Sub

Public Sub TargetUseMagic(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
End Sub

Public Sub TartgetWillAtack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub TartgetDidHit(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub TargetFailedAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

Public Sub TargetWasDamaged(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

