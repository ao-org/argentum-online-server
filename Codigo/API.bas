Attribute VB_Name = "API"
Option Explicit

Public packetResend As New CColaArray

Public Sub HandleIncomingAPIData(ByRef data As String)
    
    With frmAPISocket
    
        ' Para debuguear :P
        If .Visible Then
            .txtResponse.Text = vbNullString
            .txtResponse.Text = data
            DoEvents
        End If
        
    End With
    
    ' Parseamos el JSON que recibimo.
    Dim response As Object: Set response = mod_JSON.parse(data)
    
    ' Esta es la acción que la API nos pide ejecutar.
    Dim Command As String: Command = response.Item("header").Item("action")
    
    Select Case Command
    
        Case "user_load"
            ' TODO ?
        
        Case "recursos_reload"
            ' TODO !
        
        Case Else
            Call RegistrarError(-API_Port, "Hemos recibido un comando invalido de la API." & vbNewLine & "Comando: " & Command, "API.HandleIncomingAPIData")
            
    End Select

End Sub

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

Sub SaveUserAPI(ByVal UserIndex As Integer, Optional ByVal Logout As Boolean = False)
        
        On Error GoTo SaveUserAPI_Err:

        Dim SavePacket As New JS_Object
        Dim Header As New JS_Object
        Dim Body As New JS_Object
        Dim Main As New JS_Object
    
100     Header.Item("action") = "user_save"
102     Header.Item("expectsResponse") = False
        
104     SavePacket.Item("header") = Header
   
106     With UserList(UserIndex)

            '*************************************************************
            '   USER
            '*************************************************************
108         Body.Item("user") = API_User.Principal(UserIndex, Logout)
            
            '*************************************************************
            '   ATRIBUTOS
            '*************************************************************
110         Body.Item("attribute") = API_User.Atributos(UserIndex)
            
            '*************************************************************
            '   HECHIZOS
            '*************************************************************
112         Body.Item("spell") = API_User.Hechizo(UserIndex)
            
            '*************************************************************
            '   INVENTARIO
            '*************************************************************
114         Body.Item("inventory_item") = API_User.Inventario(UserIndex)
            
            '*************************************************************
            '   INVENTARIO DEL BANCO
            '*************************************************************
116         Body.Item("bank_item") = API_User.InventarioBanco(UserIndex)
            
            '*************************************************************
            '   SKILLS
            '*************************************************************
118         Body.Item("skillpoint") = API_User.Habilidades(UserIndex)
            
            '*************************************************************
            '   MASCOTAS
            '*************************************************************
120         Body.Item("pet") = API_User.Mascotas(UserIndex)
            
            '*************************************************************
            '   QUESTS
            '*************************************************************
122         Body.Item("quest") = API_User.Quest(UserIndex)
        
            '*************************************************************
            '   QUESTS TERMINADAS
            '*************************************************************
124         If .QuestStats.NumQuestsDone > 0 Then
126             Body.Item("quest_done") = API_User.QuestTerminadas(UserIndex)
            End If
        
            '*************************************************************
            '   ENVIAMOS A LA API
            '*************************************************************
128         SavePacket.Item("body") = Body

130         Dim UserData As String: UserData = SavePacket.ToString
         
            ' Para fines de desarrollo
132         If frmAPISocket.Visible Then frmAPISocket.txtSend = UserData
134         Debug.Print vbNewLine & UserData
        
            'Lo mandamos a la API
136         Call API.SendDataAPI(UserData)
            
        End With

        Exit Sub

SaveUserAPI_Err:
138     Call RegistrarError(Err.Number, Err.Description, "API.SaveUserAPI", Erl)
        
End Sub

