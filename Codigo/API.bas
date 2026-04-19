Attribute VB_Name = "API"
' Argentum 20 Game Server
'
'    Copyright (C) 2026 Noland Studios LTD
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
Option Explicit

Public packetResend As New CColaArray


Public Sub SendDataAPI(ByRef data As String)
    
    On Error GoTo ErrHandler:

    If frmAPISocket.Socket.State = sckConnected Then
        Call frmAPISocket.Socket.SendData(data)
        
    Else
        'Lo agrego a la cola para enviarlo mas tarde.
        Call API.packetResend.Push(data)

    End If

    Exit Sub
    
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "API_Manager.SendDataAPI")
    
End Sub


