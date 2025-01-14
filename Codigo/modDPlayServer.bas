Attribute VB_Name = "modDPlayServer"
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
'
Option Explicit
Option Base 0

Public Const AppGuid = "{5726CF1F-702B-4008-98BC-BF9C95F9E288}"

Public dx As New DirectX8
Public dps As DirectPlay8Server
Public dpa As DirectPlay8Address
Public glNumPlayers As Long
Public gfStarted As Boolean

Public Sub InitDPlay()
    Set dps = dx.DirectPlayServerCreate
    Set dpa = dx.DirectPlayAddressCreate
End Sub

Public Sub Cleanup()
    'Shut down our message handler
    If Not dps Is Nothing Then dps.UnRegisterMessageHandler
    'Close down our session
    If Not dps Is Nothing Then dps.Close
    Set dps = Nothing
    Set dpa = Nothing
    Set dx = Nothing
End Sub

