VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCapturaParticipante"
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
Dim UserRef As t_UserReference
Public LastPosMap As Integer
Public LastPosX As Byte
Public LastPosY As Byte
Public Team As Byte


Public Sub SetPlayer(ByVal UserIndex As Integer)
    On Error Goto SetPlayer_Err
    Call SetUserRef(UserRef, UserIndex)
    Exit Sub
SetPlayer_Err:
    Call TraceError(Err.Number, Err.Description, "clsCapturaParticipante.SetPlayer", Erl)
End Sub

Public Function PlayerIndex() As Integer
    On Error Goto PlayerIndex_Err
    PlayerIndex = UserRef.ArrayIndex
    Exit Function
PlayerIndex_Err:
    Call TraceError(Err.Number, Err.Description, "clsCapturaParticipante.PlayerIndex", Erl)
End Function

Public Function IsValid() As Boolean
    On Error Goto IsValid_Err
    IsValid = IsValidUserRef(UserRef)
    Exit Function
IsValid_Err:
    Call TraceError(Err.Number, Err.Description, "clsCapturaParticipante.IsValid", Erl)
End Function
