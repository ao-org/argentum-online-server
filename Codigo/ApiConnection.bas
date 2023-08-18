Attribute VB_Name = "ApiConnection"
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

Public Sub ApiEndpointBanUser(username)
    Endpoint = API_URL_SERVER & "/banUserInMysql/" & username
    Call SendGETRequest(Endpoint)
End Sub

Public Sub ApiEndpointBanAccount(account)
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

ErrorHandler:

    If Err.Number <> 0 Then
        Call LogError("Error GET endpoint: " & Endpoint & ". La Api parece estar offline. " & Err.Number & " - " & Err.Description)
    End If
    
End Sub
