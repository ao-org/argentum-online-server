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
Public dx  As New DirectX8
Public dps As DirectPlay8Server
Public dpa As DirectPlay8Address

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

Public Sub HandleDPlayError(ByVal ErrNumber As Long, ByVal ErrDescription As String, ByVal place As String, ByVal line As String)
    Select Case Err.Number
        Case DPNERR_INVALIDPLAYER
            Call TraceError(ErrNumber, "DPNERR_INVALIDPLAYER: The player ID is not recognized as a valid player ID for this game session.", place, line)
        Case DPNERR_INVALIDPARAM
            Call TraceError(ErrNumber, "One or more of the parameters passed to the method are invalid.", place, line)
        Case DPNERR_NOTHOST:
            Call TraceError(ErrNumber, _
                    "The client attempted to connect to a nonhost computer. Additionally, this error value may be returned by a nonhost that tried to set the application description. ", _
                    place, line)
        Case DPNERR_INVALIDFLAGS
            Call TraceError(ErrNumber, "The flags passed to this method are invalid.", place, line)
        Case DPNERR_TIMEDOUT
            Call TraceError(ErrNumber, "The operation could not complete because it has timed out.", place, line)
        Case Else
            Call TraceError(ErrNumber, "Unknown error", place, line)
    End Select
    Err.Clear
End Sub
