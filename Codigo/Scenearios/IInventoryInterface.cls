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

'request if user can pickup an item, return false if user can't pick up item at pos
Public Function UserCanPickUpItem(ByVal UserIndex As Integer) As Boolean
    On Error Goto UserCanPickUpItem_Err

    Exit Function
UserCanPickUpItem_Err:
    Call TraceError(Err.Number, Err.Description, "IInventoryInterface.UserCanPickUpItem", Erl)
End Function

Public Sub UserDidPickupItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
    On Error Goto UserDidPickupItem_Err

    Exit Sub
UserDidPickupItem_Err:
    Call TraceError(Err.Number, Err.Description, "IInventoryInterface.UserDidPickupItem", Erl)
End Sub

'request if user can drop item at position
Public Function UserCanDropItemAt(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer) As Boolean
    On Error Goto UserCanDropItemAt_Err

    Exit Function
UserCanDropItemAt_Err:
    Call TraceError(Err.Number, Err.Description, "IInventoryInterface.UserCanDropItemAt", Erl)
End Function

'inform that user drop an item
Public Sub UserDropItem(ByVal UserIndex As Integer, ByVal Slot As Integer, ByVal Map As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
    On Error Goto UserDropItem_Err
    Exit Sub
UserDropItem_Err:
    Call TraceError(Err.Number, Err.Description, "IInventoryInterface.UserDropItem", Erl)
End Sub
