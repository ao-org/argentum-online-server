Attribute VB_Name = "wskapiAO"
'**************************************************************
' wskapiAO.bas
'
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit

''
' Modulo para manejar Winsock
'



Private totalProcessTime  As Currency
Private totalProcessCount As Long

'Si la variable esta en TRUE , al iniciar el WsApi se crea
'una ventana LABEL para recibir los mensajes. Al detenerlo,
'se destruye.
'Si es FALSE, los mensajes se envian al form frmMain (o el
'que sea).
#Const WSAPI_CREAR_LABEL = True

Private Const SD_BOTH As Long = &H2

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const WS_CHILD = &H40000000
Public Const GWL_WNDPROC = (-4)
Private Const SIZE_RCVBUF        As Long = 10240
Private Const SIZE_SNDBUF        As Long = 10240
    
Private Const MAX_ITERATIONS_HID As Long = 200

''
'Esto es para agilizar la busqueda del slot a partir de un socket dado,
'sino, la funcion BuscaSlotSock se nos come todo el uso del CPU.
'
' @param Sock sock
' @param slot slot
'
Public Type tSockCache

    Sock As Long
    slot As Long

End Type

Public WSAPISock2Usr  As New Dictionary

' ====================================================================================
' ====================================================================================

Public OldWProc       As Long
Public ActualWProc    As Long
Public hWndMsg        As Long

' ====================================================================================
' ====================================================================================

Public SockListen     As Long
Public LastSockListen As Long

' ====================================================================================
' ====================================================================================

Public Function BuscaSlotSock(ByVal S As Long) As Long

        On Error GoTo BuscaSlotSock_Err
        
100     If WSAPISock2Usr.Exists(S) Then
102         BuscaSlotSock = WSAPISock2Usr.Item(S)
        Else
104         BuscaSlotSock = -1
        End If
        Exit Function
            
HayError:
            
BuscaSlotSock_Err:
106     Call RegistrarError(Err.Number, Err.Description, "wskapiAO.BuscaSlotSock", Erl)

108     Resume Next

End Function

Public Sub AgregaSlotSock(ByVal Sock As Long, ByVal slot As Long)
        
        On Error GoTo AgregaSlotSock_Err
        
100     Debug.Print "AgregaSockSlot"

102     If WSAPISock2Usr.Count > MaxUsers Then
104         Call CloseSocket(slot)
            Exit Sub
        End If

106     Call WSAPISock2Usr.Add(Sock, slot)
        
        Exit Sub

AgregaSlotSock_Err:
108     Call RegistrarError(Err.Number, Err.Description, "wskapiAO.AgregaSlotSock", Erl)

110     Resume Next
        
End Sub

Public Sub BorraSlotSock(ByVal Sock As Long)
        
        On Error GoTo BorraSlotSock_Err
        
100     If Not WSAPISock2Usr.Exists(Sock) Then Exit Sub

        Dim cant As Long

102     cant = WSAPISock2Usr.Count

104     WSAPISock2Usr.Remove Sock

106     Debug.Print "BorraSockSlot " & cant & " -> " & WSAPISock2Usr.Count

        Exit Sub

BorraSlotSock_Err:
108     Call RegistrarError(Err.Number, Err.Description, "wskapiAO.BorraSlotSock", Erl)
        
End Sub



Public Sub LogCustom(ByVal str As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\custom.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & "(" & Timer & ") " & str
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub LogApiSock(ByVal str As String)

        On Error GoTo ErrHandler

        Dim nfile As Integer

100     nfile = FreeFile ' obtenemos un canal
102     Open App.Path & "\logs\wsapi.log" For Append Shared As #nfile
104     Print #nfile, Date & " " & Time & " " & str
106     Close #nfile

        Exit Sub

ErrHandler:

End Sub

Public Sub EventoSockAccept(ByVal UserSocketID As Long, UserIP As Long)
        
    On Error GoTo EventoSockAccept_Err

    '========================
    'USO DE LA API DE WINSOCK
    '========================
    
    Dim NewIndex  As Integer
    Dim i         As Long
    Dim tStr      As String
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   BIENVENIDO AL SERVIDOR!!!!!!!!
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim Data() As Byte
    
    'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
    NewIndex = NextOpenUser ' Nuevo indice
    
    If NewIndex <= MaxUsers Then
        
        'Make sure both outgoing and incoming data buffers are clean
        Call UserList(NewIndex).incomingData.Clean
        Call UserList(NewIndex).outgoingData.Clean
        
        UserList(NewIndex).IP = GetAscIP(UserIP)

        'Busca si esta banneada la ip
        If IP_Blacklist.Exists(UserList(NewIndex).IP) <> 0 Then
            Call WriteShowMessageBox(NewIndex, "Se te ha prohibido la entrada al servidor. Cod: #0003")
                    
            Data = UserList(NewIndex).outgoingData.ReadAll

            Call API_send(UserSocketID, Data(0), ByVal UBound(Data()) + 1, ByVal 0)

            Call frmMain.Winsock.WSA_CloseSocket(UserSocketID)
            Exit Sub

        End If
        
        If NewIndex > LastUser Then LastUser = NewIndex
        
        UserList(NewIndex).ConnID = UserSocketID
        UserList(NewIndex).ConnIDValida = True
        
        Call AgregaSlotSock(UserSocketID, NewIndex)
            
    Else
    
        Dim TempBuffer As t_DataBuffer
            TempBuffer = PrepareMessageErrorMsg("El server se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")

        Data = TempBuffer.Data
        
        Call API_send(ByVal UserSocketID, Data(0), ByVal TempBuffer.Length, ByVal 0)
        Call frmMain.Winsock.WSA_CloseSocket(UserSocketID)

    End If

    Exit Sub

EventoSockAccept_Err:
    Call RegistrarError(Err.Number, Err.Description, "wskapiAO.EventoSockAccept", Erl)

    Resume Next
        
End Sub
 
Public Sub EventoSockRead(ByVal slot As Integer, ByRef Datos() As Byte, ByVal Length As Long)
        
        On Error GoTo EventoSockRead_Err

        Dim a As Currency
        Dim f As Currency

100     Call QueryPerformanceCounter(a)

102     With UserList(slot)
 
            #If AntiExternos = 1 Then

                ' Acá aplicamos la encriptacion Xor al paquete
108             Call Security.XorData(Datos, Length - 1, UserList(slot).XorIndexIn)

            #End If

110         Call .incomingData.WriteBlock(Datos, Length)

112         If .ConnID <> -1 Then

                ' WyroX: Pongo un límite a este loop... en caso de que por algún error bloquee el server
                Dim Iterations As Long
                Dim PacketID   As Byte
                Dim LastPacket As Byte

114             Do While HandleIncomingData(slot)
116                 PacketID = UserList(slot).incomingData.PeekByte

118                 If PacketID = LastPacket Then

120                     Iterations = Iterations + 1

122                     If Iterations >= MAX_ITERATIONS_HID Then
124                         Call RegistrarError(-1, "Se supero el maximo de iteraciones de HandleIncomingData. Paquete: " & PacketID, "wskapiAO.EventoSockRead", Erl)
126                         Call CloseSocket(slot)
                            Exit Do

                        End If

                    Else
128                     Iterations = 0
130                     LastPacket = PacketID

                    End If

                Loop
                
            Else
                Exit Sub

            End If
   
        End With

132     Call QueryPerformanceCounter(f)

134     totalProcessTime = totalProcessTime + (f - a)
136     totalProcessCount = totalProcessCount + 1

        Exit Sub

EventoSockRead_Err:
138     Call RegistrarError(Err.Number, Err.Description, "wskapiAO.EventoSockRead", Erl)

140     Resume Next
        
End Sub

Public Sub EventoSockClose(ByVal slot As Integer)
        
        On Error GoTo EventoSockClose_Err

        'Es el mismo user al que está revisando el centinela??
        'Si estamos acá es porque se cerró la conexión, no es un /salir, y no queremos banearlo....
100     If Centinela.RevisandoUserIndex = slot Then Call modCentinela.CentinelaUserLogout
    
102     If UserList(slot).flags.UserLogged Then
104         Call CloseSocketSL(slot)
106         Call Cerrar_Usuario(slot)
        Else
108         Call CloseSocket(slot)

        End If

        Exit Sub

EventoSockClose_Err:
110     Call RegistrarError(Err.Number, Err.Description, "wskapiAO.EventoSockClose", Erl)

112     Resume Next
        
End Sub

Public Sub WSApiReiniciarSockets()
        
        On Error GoTo WSApiReiniciarSockets_Err

        Dim i As Long
        
        Set frmMain.Winsock = Nothing
        
        'Cierra todas las conexiones
102     For i = 1 To MaxUsers

104         If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
106             Call CloseSocket(i)

            End If
        
            'Call ResetUserSlot(i)
108     Next i
    
110     For i = 1 To MaxUsers
112         Set UserList(i).incomingData = Nothing
114         Set UserList(i).outgoingData = Nothing
116     Next i
    
        ' No 'ta el PRESERVE :p
118     ReDim UserList(1 To MaxUsers)

120     For i = 1 To MaxUsers
122         UserList(i).ConnID = -1
124         UserList(i).ConnIDValida = False
        
126         Set UserList(i).incomingData = New clsByteQueue
128         Set UserList(i).outgoingData = New clsByteQueue
130     Next i
    
132     LastUser = 1
134     NumUsers = 0
        
        Set frmMain.Winsock = New clsWinsock
        
138     Call Sleep(100)

        Exit Sub

WSApiReiniciarSockets_Err:
144     Call RegistrarError(Err.Number, Err.Description, "wskapiAO.WSApiReiniciarSockets", Erl)

146     Resume Next
        
End Sub

Public Function CondicionSocket(ByRef lpCallerId As WSABUF, ByRef lpCallerData As WSABUF, ByRef lpSQOS As FLOWSPEC, ByVal Reserved As Long, ByRef lpCalleeId As WSABUF, ByRef lpCalleeData As WSABUF, ByRef Group As Long, ByVal dwCallbackData As Long) As Long
        
        On Error GoTo CondicionSocket_Err

        Dim sa As sockaddr
    
        'Check if we were requested to force reject
100     If dwCallbackData = 1 Then
102         CondicionSocket = CF_REJECT
            Exit Function

        End If
    
        'Get the address
104     Call CopyMemory(sa, ByVal lpCallerId.LpBuffer, lpCallerId.dwBufferLen)

        
        ' Si esta en la lista de IPs prohibidas, rechazamos la conexion
        If IP_Blacklist.Exists(GetAscIP(sa.sin_addr)) Then
            Debug.Print "La IP " & GetAscIP(sa.sin_addr) & " esta baneada, CONEXION RECHAZADA"
            CondicionSocket = CF_REJECT
            Exit Function
        End If
        
110     CondicionSocket = CF_ACCEPT 'En realdiad es al pedo, porque CondicionSocket se inicializa a 0, pero así es más claro....
        
        Exit Function

CondicionSocket_Err:
112     Call RegistrarError(Err.Number, Err.Description, "wskapiAO.CondicionSocket", Erl)

114     Resume Next
        
End Function
