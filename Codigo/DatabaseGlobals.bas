Attribute VB_Name = "DatabaseGlobals"
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

Option Base 0

Public Database_Enabled     As Boolean
Public Database_Driver      As String
Public Database_Source      As String
Public Database_Host        As String
Public Database_Name        As String
Public Database_Username    As String
Public Database_Password    As String

Public Const MAX_ASYNC     As Byte = 20
Public Current_async       As Byte

Public Connection          As ADODB.Connection
Public Connection_async(1 To MAX_ASYNC)    As ADODB.Connection

Public Builder             As cStringBuilder

Public Function post_increment(ByRef value As Integer) As Integer
     post_increment = value
     value = value + 1
End Function
