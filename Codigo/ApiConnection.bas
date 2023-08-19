Attribute VB_Name = "ApiConnection"
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
'**************************************************************
'Autor: Lucas Recoaro (Recox)
'Este modulo es el encargado de hacer los requests a los diferentes endpoints de la API
'La API esta escrita en javascript (Node.js/Express) y podemos desde obtener datos a hacer backup de charfiles o cuentas
'https://github.com/ao-libre/ao-api-server
'**************************************************************

Option Explicit
Private XmlHttp As Object
Private Endpoint As String
Private Parameters As String

Public Sub ApiEndpointBanUser(ByVal username As String)
    Endpoint = API_URL_SERVER & "/banUserInMysql/" & username
    Call SendGETRequest(Endpoint)
End Sub

Public Sub ApiEndpointBanAccount(ByVal account As String)
    Endpoint = API_URL_SERVER & "/banAccountInMysql/" & account
    Call SendGETRequest(Endpoint)
End Sub

' Ejemplo de como enviar un POST Request.
' Public Sub ApiEndpointSendServerDataToApiToShowOnlineUsers()
'     Endpoint = "https://api.argentumonline.org/api/v1/servers/sendUsersOnline"
'     Parameters = "serverName=" & NombreServidor & "&quantityUsers=" & LastUser & "&ip=" & IpPublicaServidor & "&port=" & Puerto

'     Call SendPOSTRequest(Endpoint, Parameters)
' End Sub

Private Sub SendPOSTRequest(ByVal Endpoint As String, ByVal Parameters As String)

On Error GoTo ErrorHandler

    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    
    XmlHttp.Open "POST", Endpoint, True
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        
    'Por alguna razon tengo que castearlo a string, sino no funciona, la verdad no tengo idea por que ya que la variable es String
    XmlHttp.Send CStr(Parameters)
    
    Set XmlHttp = Nothing
    Exit Sub
    
ErrorHandler:

    If Err.Number <> 0 Then
        Call LogError("Error POST endpoint: " & Endpoint & ". La Api parece estar offline. " & Err.Number & " - " & Err.Description)
    End If
    
End Sub

Private Sub SendGETRequest(ByVal Endpoint As String)
On Error GoTo ErrorHandler

    Set XmlHttp = CreateObject("Microsoft.XmlHttp")
    
    XmlHttp.Open "GET", Endpoint, True
    XmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XmlHttp.Send

    Set XmlHttp = Nothing
    Exit Sub

ErrorHandler:

    If Err.Number <> 0 Then
        Call LogError("Error GET endpoint: " & Endpoint & ". La Api parece estar offline. " & Err.Number & " - " & Err.Description)
    End If
    
End Sub
