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

#If UsarQueSocket = 1 Then

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
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

    Private Const WS_CHILD = &H40000000
    Public Const GWL_WNDPROC = (-4)
    Private Const SIZE_RCVBUF As Long = 8192
    Private Const SIZE_SNDBUF As Long = 8192
    
    Private Const MAX_ITERATIONS_HID As Long = 100

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

    Public WSAPISock2Usr  As New Collection

    ' ====================================================================================
    ' ====================================================================================

    Public OldWProc       As Long
    Public ActualWProc    As Long
    Public hWndMsg        As Long

    ' ====================================================================================
    ' ====================================================================================

    Public SockListen     As Long
    Public LastSockListen As Long

#End If

' ====================================================================================
' ====================================================================================

Public Sub IniciaWsApi(ByVal hwndParent As Long)
        
        On Error GoTo IniciaWsApi_Err
        
        #If UsarQueSocket = 1 Then
100         Call LogApiSock("IniciaWsApi")
102         Debug.Print "IniciaWsApi"

            #If WSAPI_CREAR_LABEL Then
104             hWndMsg = CreateWindowEx(0, "STATIC", "AOMSG", WS_CHILD, 0, 0, 0, 0, hwndParent, 0, App.hInstance, ByVal 0&)
            #Else
106             hWndMsg = hwndParent
            #End If 'WSAPI_CREAR_LABEL

108         OldWProc = SetWindowLong(hWndMsg, GWL_WNDPROC, AddressOf WndProc)
110         ActualWProc = GetWindowLong(hWndMsg, GWL_WNDPROC)

            Dim Desc As String

112         Call StartWinsock(Desc)

        #End If

        
        Exit Sub

IniciaWsApi_Err:
114     Call RegistrarError(Err.Number, Err.description, "wskapiAO.IniciaWsApi", Erl)
116     Resume Next
        
End Sub

Public Sub LimpiaWsApi()
        
        On Error GoTo LimpiaWsApi_Err
        
        #If UsarQueSocket = 1 Then

100         Call LogApiSock("LimpiaWsApi")

102         If WSAStartedUp Then
104             Call EndWinsock

            End If

106         If OldWProc <> 0 Then
108             SetWindowLong hWndMsg, GWL_WNDPROC, OldWProc
110             OldWProc = 0

            End If

            #If WSAPI_CREAR_LABEL Then

112             If hWndMsg <> 0 Then
114                 DestroyWindow hWndMsg

                End If

            #End If

        #End If

        
        Exit Sub

LimpiaWsApi_Err:
116     Call RegistrarError(Err.Number, Err.description, "wskapiAO.LimpiaWsApi", Erl)
118     Resume Next
        
End Sub

Public Function BuscaSlotSock(ByVal S As Long) As Long
        #If UsarQueSocket = 1 Then

            On Error GoTo hayerror

100         BuscaSlotSock = WSAPISock2Usr.Item(CStr(S))
            Exit Function

102 hayerror:   ' The socket was already removed

104         BuscaSlotSock = -1
106         Err.Clear

        #End If

End Function

Public Sub AgregaSlotSock(ByVal Sock As Long, ByVal slot As Long)
        
        On Error GoTo AgregaSlotSock_Err
        
100     Debug.Print "AgregaSockSlot"
        #If (UsarQueSocket = 1) Then

            'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("AgregaSlotSock:: sock=" & Sock & " slot=" & Slot)

102         If WSAPISock2Usr.Count > MaxUsers Then
                'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("Imposible agregarSlotSock (wsapi2usr.count>maxusers)")
104             Call CloseSocket(slot)
                Exit Sub

            End If

106         WSAPISock2Usr.Add CStr(slot), CStr(Sock)

            'Dim Pri As Long, Ult As Long, Med As Long
            'Dim LoopC As Long
            '
            'If WSAPISockChacheCant > 0 Then
            '    Pri = 1
            '    Ult = WSAPISockChacheCant
            '    Med = Int((Pri + Ult) / 2)
            '
            '    Do While (Pri <= Ult) And (Ult > 1)
            '        If Sock < WSAPISockChache(Med).Sock Then
            '            Ult = Med - 1
            '        Else
            '            Pri = Med + 1
            '        End If
            '        Med = Int((Pri + Ult) / 2)
            '    Loop
            '
            '    Pri = IIf(Sock < WSAPISockChache(Med).Sock, Med, Med + 1)
            '    Ult = WSAPISockChacheCant
            '    For LoopC = Ult To Pri Step -1
            '        WSAPISockChache(LoopC + 1) = WSAPISockChache(LoopC)
            '    Next LoopC
            '    Med = Pri
            'Else
            '    Med = 1
            'End If
            'WSAPISockChache(Med).Slot = Slot
            'WSAPISockChache(Med).Sock = Sock
            'WSAPISockChacheCant = WSAPISockChacheCant + 1

        #End If

        
        Exit Sub

AgregaSlotSock_Err:
108     Call RegistrarError(Err.Number, Err.description, "wskapiAO.AgregaSlotSock", Erl)
110     Resume Next
        
End Sub

Public Sub BorraSlotSock(ByVal Sock As Long)
        
        On Error GoTo BorraSlotSock_Err
    
        
        #If (UsarQueSocket = 1) Then

            Dim cant As Long

100         cant = WSAPISock2Usr.Count

            

102         WSAPISock2Usr.Remove CStr(Sock)

104         Debug.Print "BorraSockSlot " & cant & " -> " & WSAPISock2Usr.Count

        #End If

        
        Exit Sub

BorraSlotSock_Err:
        Call RegistrarError(Err.Number, Err.description, "wskapiAO.BorraSlotSock", Erl)

        
End Sub

Public Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        
        On Error GoTo WndProc_Err
    
        
        #If UsarQueSocket = 1 Then

            

100         Dim ttt As Long: ttt = GetTickCount()

            Dim ret      As Long
            Dim Tmp()    As Byte
            Dim S        As Long, e As Long
            Dim n        As Integer
            Dim Dale     As Boolean
            Dim UltError As Long

102         WndProc = 0

104         Select Case msg

                Case 1025

106                 S = wParam
108                 e = WSAGetSelectEvent(lParam)
                    'Debug.Print "Msg: " & msg & " W: " & wParam & " L: " & lParam
                    'Call LogApiSock("Msg: " & msg & " W: " & wParam & " L: " & lParam)
    
110                 Select Case e

                        Case FD_ACCEPT

                            'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("FD_ACCEPT")
112                         If S = SockListen Then
                                'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("sockLIsten = " & s & ". Llamo a Eventosocketaccept")
114                             Call EventoSockAccept(S)
                            End If
        
                            '    Case FD_WRITE
                            '        N = BuscaSlotSock(s)
                            '        If N < 0 And s <> SockListen Then
                            '            'Call apiclosesocket(s)
                            '            call WSApiCloseSocket(s)
                            '            Exit Function
                            '        End If
                            '

                            '        Call IntentarEnviarDatosEncolados(N)
                            '
                            ''        Dale = UserList(N).ColaSalida.Count > 0
                            ''        Do While Dale
                            ''            Ret = WsApiEnviar(N, UserList(N).ColaSalida.Item(1), False)
                            ''            If Ret <> 0 Then
                            ''                If Ret = WSAEWOULDBLOCK Then
                            ''                    Dale = False
                            ''                Else
                            ''                    'y aca que hacemo' ?? help! i need somebody, help!
                            ''                    Dale = False
                            ''                    Debug.Print "ERROR AL ENVIAR EL DATO DESDE LA COLA " & Ret & ": " & GetWSAErrorString(Ret)
                            ''                End If
                            ''            Else
                            ''            '    Debug.Print "Dato de la cola enviado"
                            ''                UserList(N).ColaSalida.Remove 1
                            ''                Dale = (UserList(N).ColaSalida.Count > 0)
                            ''            End If
                            ''        Loop

116                     Case FD_READ
        
118                         n = BuscaSlotSock(S)

120                         If n < 0 And S <> SockListen Then
                                'Call apiclosesocket(s)
122                             Call WSApiCloseSocket(S)
                                Exit Function

                            End If
        
                            'Call WSAAsyncSelect(s, hWndMsg, ByVal 1025, ByVal (0))
        
                            '4k de buffer
124                         ReDim Preserve Tmp(SIZE_RCVBUF - 1) As Byte
        
126                         ret = recv(S, Tmp(0), SIZE_RCVBUF, 0)

                            ' Comparo por = 0 ya que esto es cuando se cierra
                            ' "gracefully". (mas abajo)
128                         If ret < 0 Then
130                             UltError = Err.LastDllError

132                             If UltError = WSAEMSGSIZE Then
134                                 Debug.Print "WSAEMSGSIZE"
136                                 ret = SIZE_RCVBUF
                                
                                Else
138                                 Debug.Print "Error en Recv: " & GetWSAErrorString(UltError)
140                                 Call LogApiSock("Error en Recv: N=" & n & " S=" & S & " Str=" & GetWSAErrorString(UltError))
                
                                    'no hay q llamar a CloseSocket() directamente,
                                    'ya q pueden abusar de algun error para
                                    'desconectarse sin los 10segs. CREEME.
                                    '    Call C l o s e Socket(N)
            
142                                 Call CloseSocketSL(n)
144                                 Call Cerrar_Usuario(n)
                                    Exit Function

                                End If

146                         ElseIf ret = 0 Then
148                             Call CloseSocketSL(n)
150                             Call Cerrar_Usuario(n)

                            End If
        
                            'Call WSAAsyncSelect(s, hWndMsg, ByVal 1025, ByVal (FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT))
        
152                         ReDim Preserve Tmp(ret - 1) As Byte
        
                            'Call LogApiSock("WndProc:FD_READ:N=" & N & ":TMP=" & Tmp)
        
154                         Call EventoSockRead(n, Tmp)
        
156                     Case FD_CLOSE
                            'Debug.Print WSAGETSELECTERROR(lParam)
158                         n = BuscaSlotSock(S)

160                         If S <> SockListen Then Call apiclosesocket(S)
        
162                         Call LogApiSock("WndProc:FD_CLOSE:N=" & n & ":Err=" & WSAGetAsyncError(lParam))
        
164                         If n > 0 Then
166                             Call BorraSlotSock(S)
168                             UserList(n).ConnID = -1
170                             UserList(n).ConnIDValida = False
172                             Call EventoSockClose(n)

                            End If
        
                    End Select

174             Case Else
176                 WndProc = CallWindowProc(OldWProc, hWnd, msg, wParam, lParam)

            End Select

        #End If

178     OutputDebugString "SocketProc: " & (GetTickCount() - ttt)

        
        Exit Function

WndProc_Err:
        Call RegistrarError(Err.Number, Err.description, "wskapiAO.WndProc", Erl)

        
End Function

'Retorna 0 cuando se envió o se metio en la cola,
'retorna <> 0 cuando no se pudo enviar o no se pudo meter en la cola
Public Function WsApiEnviar(ByVal slot As Integer, ByRef str As String) As Long
        
        On Error GoTo WsApiEnviar_Err
        
        #If UsarQueSocket = 1 Then

            'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("WsApiEnviar:: slot=" & Slot & " str=" & str & " len(str)=" & Len(str) & " encolar=" & Encolar)

            Dim ret      As String
            Dim UltError As Long
            Dim Retorno  As Long
            Dim data()   As Byte

100         ReDim Preserve data(Len(str) - 1) As Byte

102         data = StrConv(str, vbFromUnicode)

104         Retorno = 0

            'Debug.Print ">>>> " & str

106         If UserList(slot).ConnID <> -1 And UserList(slot).ConnIDValida Then
108             ret = send(ByVal UserList(slot).ConnID, data(0), ByVal UBound(data()) + 1, ByVal 0)

110             If ret < 0 Then
112                 UltError = Err.LastDllError

114                 If UltError = WSAEWOULDBLOCK Then
            
                        ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
116                     Call UserList(slot).outgoingData.WriteASCIIStringFixed(str)

                    End If

118                 Retorno = UltError

                End If

120         ElseIf UserList(slot).ConnID <> -1 And Not UserList(slot).ConnIDValida Then

122             If Not UserList(slot).Counters.Saliendo Then
124                 Retorno = -1

                End If

            End If

126         WsApiEnviar = Retorno

        #End If

        
        Exit Function

WsApiEnviar_Err:
128     Call RegistrarError(Err.Number, Err.description, "wskapiAO.WsApiEnviar", Erl)
130     Resume Next
        
End Function

Public Sub LogCustom(ByVal str As String)
        #If (UsarQueSocket = 1) Then

            On Error GoTo ErrHandler

            Dim nfile As Integer

100         nfile = FreeFile ' obtenemos un canal
102         Open App.Path & "\logs\custom.log" For Append Shared As #nfile
104         Print #nfile, Date & " " & Time & "(" & Timer & ") " & str
106         Close #nfile

            Exit Sub

ErrHandler:

        #End If

End Sub

Public Sub LogApiSock(ByVal str As String)
        #If (UsarQueSocket = 1) Then

            On Error GoTo ErrHandler

            Dim nfile As Integer

100         nfile = FreeFile ' obtenemos un canal
102         Open App.Path & "\logs\wsapi.log" For Append Shared As #nfile
104         Print #nfile, Date & " " & Time & " " & str
106         Close #nfile

            Exit Sub

ErrHandler:

        #End If

End Sub

Public Sub EventoSockAccept(ByVal SockID As Long)
        
        On Error GoTo EventoSockAccept_Err
        
        #If UsarQueSocket = 1 Then
            '==========================================================
            'USO DE LA API DE WINSOCK
            '========================
    
            Dim NewIndex  As Integer
            Dim ret       As Long
            Dim Tam       As Long, sa As sockaddr
            Dim NuevoSock As Long
            Dim i         As Long
            Dim tStr      As String
    
100         Tam = sockaddr_size
    
            '=============================================
            'SockID es en este caso es el socket de escucha,
            'a diferencia de socketwrench que es el nuevo
            'socket de la nueva conn
    
            'Modificado por Maraxus
            'Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 0)
102         ret = accept(SockID, sa, Tam)

104         If ret = INVALID_SOCKET Then
106             i = Err.LastDllError
108             Call LogCriticEvent("Error en Accept() API " & i & ": " & GetWSAErrorString(i))
                Exit Sub

            End If
    
110         If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
112             Call WSApiCloseSocket(NuevoSock)
                Exit Sub

            End If

            'If Ret = INVALID_SOCKET Then
            '    If Err.LastDllError = 11002 Then
            '        ' We couldn't decide if to accept or reject the connection
            '        'Force reject so we can get it out of the queue
            '        LogCustom ("Pre WSAAccept CallbackData=1")
            '        Ret = WSAAccept(SockID, sa, Tam, AddressOf CondicionSocket, 1)
            '        LogCustom ("WSAccept Callbackdata 1, devuelve " & Ret)
            '        Call LogCriticEvent("Error en WSAAccept() API 11002: No se pudo decidir si aceptar o rechazar la conexión.")
            '    Else
            '        i = Err.LastDllError
            '        LogCustom ("Error en WSAAccept() API " & i & ": " & GetWSAErrorString(i))
            '        Call LogCriticEvent("Error en WSAAccept() API " & i & ": " & GetWSAErrorString(i))
            '        Exit Sub
            '    End If
            'End If

            'saco nagle
    
114         NuevoSock = ret

            'Call setsockopt(wsock.SocketHandle, 6, 1, True, 4) 'old: If setsockopt(NuevoSock, SOL_SOCKET, TCP_NODELAY, True, 1) <> 0 Then
            'algoritmo de nagle vb6
            ' If setsockopt(NuevoSock, 6, 1, True, 4) <> 0 Then
116        ' If setsockopt(NuevoSock, 6, TCP_NODELAY, True, 4) <> 0 Then
118        '     i = Err.LastDllError
120        '     Call LogCriticEvent("Error al setear el delay " & i & ": " & GetWSAErrorString(i))

            'End If

            'saco nagle
    
            'Nuevo sin nagle
            'NuevoSock = Ret
             If setsockopt(NuevoSock, SOL_SOCKET, SO_LINGER, 0, 4) <> 0 Then
               i = Err.LastDllError
              Call LogCriticEvent("Error al setear lingers." & i & ": " & GetWSAErrorString(i))
             End If
            'Nuevo sin nagle
    
            'Seteamos el tamaño del buffer de entrada
122         If setsockopt(NuevoSock, SOL_SOCKET, SO_RCVBUFFER, SIZE_RCVBUF, 4) <> 0 Then
124             i = Err.LastDllError
126             Call LogCriticEvent("Error al setear el tamaño del buffer de entrada " & i & ": " & GetWSAErrorString(i))

            End If

            'Seteamos el tamaño del buffer de salida
128         If setsockopt(NuevoSock, SOL_SOCKET, SO_SNDBUFFER, SIZE_SNDBUF, 4) <> 0 Then
130             i = Err.LastDllError
132             Call LogCriticEvent("Error al setear el tamaño del buffer de salida " & i & ": " & GetWSAErrorString(i))

            End If

            'If SecurityIp.IPSecuritySuperaLimiteConexiones(sa.sin_addr) Then
            'tStr = "Limite de conexiones para su IP alcanzado."
            'Call send(ByVal NuevoSock, ByVal tStr, ByVal Len(tStr), ByVal 0)
            'Call WSApiCloseSocket(NuevoSock)
            'Exit Sub
            'End If
    
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '   BIENVENIDO AL SERVIDOR!!!!!!!!
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
            'Mariano: Baje la busqueda de slot abajo de CondicionSocket y limite x ip
134         NewIndex = NextOpenUser ' Nuevo indice
    
136         If NewIndex <= MaxUsers Then
        
                'Make sure both outgoing and incoming data buffers are clean
138             Call UserList(NewIndex).incomingData.ReadASCIIStringFixed(UserList(NewIndex).incomingData.Length)
140             Call UserList(NewIndex).outgoingData.ReadASCIIStringFixed(UserList(NewIndex).outgoingData.Length)
        
142             UserList(NewIndex).ip = GetAscIP(sa.sin_addr)

                'Busca si esta banneada la ip
144             For i = 1 To BanIps.Count

146                 If BanIps.Item(i) = UserList(NewIndex).ip Then
                        'Call apiclosesocket(NuevoSock)
148                     Call WriteErrorMsg(NewIndex, "Su IP se encuentra bloqueada en este servidor.")
                    
                        'Call SecurityIp.IpRestarConexion(sa.sin_addr)
150                     Call WSApiCloseSocket(NuevoSock)
                        Exit Sub

                    End If

152             Next i
        
154             If NewIndex > LastUser Then LastUser = NewIndex
        
156             UserList(NewIndex).ConnID = NuevoSock
158             UserList(NewIndex).ConnIDValida = True
        
160             Call AgregaSlotSock(NuevoSock, NewIndex)
            Else

                Dim str    As String
                Dim data() As Byte
        
162             str = Protocol.PrepareMessageErrorMsg("El server se encuentra lleno en este momento. Disculpe las molestias ocasionadas.")
        
164             ReDim Preserve data(Len(str) - 1) As Byte
        
166             data = StrConv(str, vbFromUnicode)
        
168             Call send(ByVal NuevoSock, data(0), ByVal UBound(data()) + 1, ByVal 0)
170             Call WSApiCloseSocket(NuevoSock)

            End If
    
        #End If

        
        Exit Sub

EventoSockAccept_Err:
172     Call RegistrarError(Err.Number, Err.description, "wskapiAO.EventoSockAccept", Erl)
174     Resume Next
        
End Sub
 
Public Sub EventoSockRead(ByVal slot As Integer, ByRef Datos() As Byte)
        
        On Error GoTo EventoSockRead_Err
        
        #If UsarQueSocket = 1 Then

            Dim a As Currency
            Dim f As Currency

100         QueryPerformanceCounter a

102         With UserList(slot)
 
                #If AntiExternos Then

104                 If UserList(slot).flags.UserLogged Then
106                     Security.NAC_D_Byte Datos, UserList(slot).Redundance
108                     Elsete
110                     Security.NAC_D_Byte Datos, 13 'DEFAULT

                    End If

                #End If
 
112             Call .incomingData.WriteBlock(Datos)

114             If .ConnID <> -1 Then

                    ' WyroX: Pongo un límite a este loop... en caso de que por algún error bloquee el server
                    Dim Iterations As Long

116                 Do While HandleIncomingData(slot)
                        Iterations = Iterations + 1
                        If Iterations >= MAX_ITERATIONS_HID Then
                            Call RegistrarError(-1, "Se supero el maximo de iteraciones de HandleIncomingData. Paquete: " & UserList(slot).incomingData.PeekByte, "wskapiAO.EventoSockRead")
                            Call CloseSocket(slot)
                            Exit Do
                        End If
                    Loop
                Else
                    Exit Sub

                End If
   
            End With

118         QueryPerformanceCounter f

120         totalProcessTime = totalProcessTime + (f - a)
122         totalProcessCount = totalProcessCount + 1

        #End If

        
        Exit Sub

EventoSockRead_Err:
124     Call RegistrarError(Err.Number, Err.description, "wskapiAO.EventoSockRead", Erl)
126     Resume Next
        
End Sub

Public Sub EventoSockClose(ByVal slot As Integer)
        
        On Error GoTo EventoSockClose_Err
        
        #If UsarQueSocket = 1 Then

            'Es el mismo user al que está revisando el centinela??
            'Si estamos acá es porque se cerró la conexión, no es un /salir, y no queremos banearlo....
100         If Centinela.RevisandoUserIndex = slot Then Call modCentinela.CentinelaUserLogout
    
102         If UserList(slot).flags.UserLogged Then
104             Call CloseSocketSL(slot)
106             Call Cerrar_Usuario(slot)
            Else
108             Call CloseSocket(slot)

            End If

        #End If

        
        Exit Sub

EventoSockClose_Err:
110     Call RegistrarError(Err.Number, Err.description, "wskapiAO.EventoSockClose", Erl)
112     Resume Next
        
End Sub

Public Sub WSApiReiniciarSockets()
        
        On Error GoTo WSApiReiniciarSockets_Err
        
        #If UsarQueSocket = 1 Then

            Dim i As Long

            'Cierra el socket de escucha
100         If SockListen >= 0 Then Call apiclosesocket(SockListen)
    
            'Cierra todas las conexiones
102         For i = 1 To MaxUsers

104             If UserList(i).ConnID <> -1 And UserList(i).ConnIDValida Then
106                 Call CloseSocket(i)

                End If
        
                'Call ResetUserSlot(i)
108         Next i
    
110         For i = 1 To MaxUsers
112             Set UserList(i).incomingData = Nothing
114             Set UserList(i).outgoingData = Nothing
116         Next i
    
            ' No 'ta el PRESERVE :p
118         ReDim UserList(1 To MaxUsers)

120         For i = 1 To MaxUsers
122             UserList(i).ConnID = -1
124             UserList(i).ConnIDValida = False
        
126             Set UserList(i).incomingData = New clsByteQueue
128             Set UserList(i).outgoingData = New clsByteQueue
130         Next i
    
132         LastUser = 1
134         NumUsers = 0
    
136         Call LimpiaWsApi
138         Call Sleep(100)
140         Call IniciaWsApi(frmMain.hWnd)
142         SockListen = ListenForConnect(Puerto, hWndMsg, "")

        #End If

        
        Exit Sub

WSApiReiniciarSockets_Err:
144     Call RegistrarError(Err.Number, Err.description, "wskapiAO.WSApiReiniciarSockets", Erl)
146     Resume Next
        
End Sub

Public Sub WSApiCloseSocket(ByVal Socket As Long)
        
        On Error GoTo WSApiCloseSocket_Err
        
        #If UsarQueSocket = 1 Then
100         Call WSAAsyncSelect(Socket, hWndMsg, ByVal 1025, ByVal (FD_CLOSE))
102         Call ShutDown(Socket, SD_BOTH)
        #End If

        
        Exit Sub

WSApiCloseSocket_Err:
104     Call RegistrarError(Err.Number, Err.description, "wskapiAO.WSApiCloseSocket", Erl)
106     Resume Next
        
End Sub

Public Function CondicionSocket(ByRef lpCallerId As WSABUF, ByRef lpCallerData As WSABUF, ByRef lpSQOS As FLOWSPEC, ByVal Reserved As Long, ByRef lpCalleeId As WSABUF, ByRef lpCalleeData As WSABUF, ByRef Group As Long, ByVal dwCallbackData As Long) As Long
        
        On Error GoTo CondicionSocket_Err
        
        #If UsarQueSocket = 1 Then

            Dim sa As sockaddr
    
            'Check if we were requested to force reject

100         If dwCallbackData = 1 Then
102             CondicionSocket = CF_REJECT
                Exit Function

            End If
    
            'Get the address
104         Call CopyMemory(sa, ByVal lpCallerId.LpBuffer, lpCallerId.dwBufferLen)
    
106         If Not SecurityIp.IpSecurityAceptarNuevaConexion(sa.sin_addr) Then
108             CondicionSocket = CF_REJECT
                Exit Function

            End If

110         CondicionSocket = CF_ACCEPT 'En realdiad es al pedo, porque CondicionSocket se inicializa a 0, pero así es más claro....
        #End If

        
        Exit Function

CondicionSocket_Err:
112     Call RegistrarError(Err.Number, Err.description, "wskapiAO.CondicionSocket", Erl)
114     Resume Next
        
End Function
