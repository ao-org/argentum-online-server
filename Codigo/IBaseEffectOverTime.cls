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

Public Property Get UniqueId() As Long
End Property

Public Property Get TypeId() As Integer
End Property

Public Property Get SharedTypeId() As Integer
End Property

Public Property Get EotId() As Integer
End Property

Public Sub Update(ByVal DeltaTime As Long)
End Sub

Public Property Get EffectType() As e_EffectType
End Property

Public Property Get RemoveMe() As Boolean
End Property

Public Property Let RemoveMe(ByVal value As Boolean)
End Property

Public Function Reset(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal NewEffectId As Integer) As Boolean
End Function

Public Sub OnRemove()
End Sub

'called when the owner of the effect cast any spell
Public Sub TargetUseMagic(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
End Sub

'called when the owner of the effect is attacking some target it should handle a ditHit or FailedAttack after this call
Public Sub TargetWillAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

'called when the owner of the effect hit something
Public Sub TargetDidHit(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

'called when the owner of the effect fails and attacks
Public Sub TargetFailedAttack(ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

'called when the owner of the effect is damaged
Public Sub TargetWasDamaged(ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
End Sub

'called to get damage reduction if requiered
Public Function ApplyDamageReduction(ByVal Damage As Long, ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType) As Long
End Function

'called when the owner of the effect attack a position (like a player attack anything infornt of him)
Public Sub TargetWillAttackPosition(ByVal Map As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
End Sub

Public Function ChangeTarget(ByVal NewTargetIndex As Integer, ByVal NewTargetType As e_ReferenceType) As Boolean
End Function

Public Property Get KeepAfterDead() As Boolean
End Property

Public Property Get CallBacksMask() As Long
End Property

Public Sub TargetChangeTerrain()
End Sub
