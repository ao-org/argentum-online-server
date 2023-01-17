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


Public Property Get NextEffect() As IBaseEffectOverTime
End Property

Public Property Let NextEffect(ByRef NextElement As IBaseEffectOverTime)
End Property

Public Sub SetTarget(ByVal index As Integer, ByVal RefType As e_ReferenceType)

End Sub

Public Property Get UniqueId() As Integer
End Property

Public Sub Update(ByVal deltaTime As Long)

End Sub

Public Function RemoveMe() As Boolean
    RemoveMe = False
End Function
