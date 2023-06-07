VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IInventoryInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'request if user can pickup an item, return false if user can't pick up item at pos
Public Function UserCanPickUpItem(ByVal UserIndex As Integer) As Boolean

End Function

Public Sub UserDidPickupItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)

End Sub

'request if user can drop item at position
Public Function UserCanDropItemAt(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer) As Boolean

End Function

'inform that user drop an item
Public Sub UserDropItem(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
End Sub
