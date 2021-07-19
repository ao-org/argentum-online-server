VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FD_SETSIZE = 64

Private Type fd_set
    fd_count As Integer
    fd_array(FD_SETSIZE) As Integer
End Type

Private Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Private Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Const hostent_size = 16

Private Type servent
    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long
End Type

Private Const servent_size = 14

Private Type protoent
    p_name As Long
    p_aliases As Long
    p_proto As Integer
End Type

Private Const protoent_size = 10

Private Const IPPROTO_TCP = 6
Private Const IPPROTO_UDP = 17

Private Const INADDR_NONE = &HFFFFFFFF
Private Const INADDR_ANY = &H0

Private Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Private Const sockaddr_size = 16

Private saZero As sockaddr

'Agregado por Maraxus
Private Type WSABUF
    dwBufferLen As Long
    LpBuffer    As Long
End Type

'Agregado por Maraxus
 Private Type FLOWSPEC
    TokenRate           As Long     'In Bytes/sec
    TokenBucketSize     As Long     'In Bytes
    PeakBandwidth       As Long     'In Bytes/sec
    Latency             As Long     'In microseconds
    DelayVariation      As Long     'In microseconds
    ServiceType         As Integer  'Guaranteed, Predictive,
    
    'Best Effort, etc.
    MaxSduSize          As Long     'In Bytes
    MinimumPolicedSize  As Long     'In Bytes
End Type

'Agregado por Maraxus
Private Const WSA_FLAG_OVERLAPPED = &H1

'Agregados por Maraxus
Private Const CF_ACCEPT = &H0
Private Const CF_REJECT = &H1

'Agregado por Maraxus
Private Const SD_RECEIVE As Long = &H0&
Private Const SD_SEND    As Long = &H1&
Private Const SD_BOTH    As Long = &H2&

Private Const INVALID_SOCKET = -1
Private Const SOCKET_ERROR = -1

Private Const SOCK_STREAM = 1
Private Const SOCK_DGRAM = 2

Private Const MAXGETHOSTSTRUCT = 1024

Private Const AF_INET = 2
Private Const PF_INET = 2

Private Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type

' Windows Sockets definitions of regular Microsoft C error constants
Const WSAEINTR = 10004
Const WSAEBADF = 10009
Const WSAEACCES = 10013
Const WSAEFAULT = 10014
Const WSAEINVAL = 10022
Const WSAEMFILE = 10024

' Windows Sockets definitions of regular Berkeley error constants
Const WSAEWOULDBLOCK = 10035
Const WSAEINPROGRESS = 10036
Const WSAEALREADY = 10037
Const WSAENOTSOCK = 10038
Const WSAEDESTADDRREQ = 10039
Const WSAEMSGSIZE = 10040
Const WSAEPROTOTYPE = 10041
Const WSAENOPROTOOPT = 10042
Const WSAEPROTONOSUPPORT = 10043
Const WSAESOCKTNOSUPPORT = 10044
Const WSAEOPNOTSUPP = 10045
Const WSAEPFNOSUPPORT = 10046
Const WSAEAFNOSUPPORT = 10047
Const WSAEADDRINUSE = 10048
Const WSAEADDRNOTAVAIL = 10049
Const WSAENETDOWN = 10050
Const WSAENETUNREACH = 10051
Const WSAENETRESET = 10052
Const WSAECONNABORTED = 10053
Const WSAECONNRESET = 10054
Const WSAENOBUFS = 10055
Const WSAEISCONN = 10056
Const WSAENOTCONN = 10057
Const WSAESHUTDOWN = 10058
Const WSAETOOMANYREFS = 10059
Const WSAETIMEDOUT = 10060
Const WSAECONNREFUSED = 10061
Const WSAELOOP = 10062
Const WSAENAMETOOLONG = 10063
Const WSAEHOSTDOWN = 10064
Const WSAEHOSTUNREACH = 10065
Const WSAENOTEMPTY = 10066
Const WSAEPROCLIM = 10067
Const WSAEUSERS = 10068
Const WSAEDQUOT = 10069
Const WSAESTALE = 10070
Const WSAEREMOTE = 10071

' Extended Windows Sockets error constant definitions
Const WSASYSNOTREADY = 10091
Const WSAVERNOTSUPPORTED = 10092
Const WSANOTINITIALISED = 10093
Const WSAHOST_NOT_FOUND = 11001
Const WSATRY_AGAIN = 11002
Const WSANO_RECOVERY = 11003
Const WSANO_DATA = 11004
Const WSANO_ADDRESS = 11004

'---ioctl Constants
Private Const FIONREAD = &H8004667F
Private Const FIONBIO = &H8004667E
Private Const FIOASYNC = &H8004667D

'---Windows System Functions
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub API_MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Private Declare Function API_lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

'---async notification constants
Private Const TCP_NODELAY = &H1                 ' Agregado por Maraxus
Private Const SOL_SOCKET = &HFFFF&
Private Const SO_LINGER = &H80&
Private Const SO_RCVBUFFER = &H1002&             ' Agregado por Maraxus
Private Const SO_SNDBUFFER = &H1001&              ' Agregado por Maraxus
Private Const SO_CONDITIONAL_ACCEPT = &H3002&    ' Agregado por Maraxus
Private Const FD_READ = &H1&
Private Const FD_WRITE = &H2&
Private Const FD_OOB = &H4&
Private Const FD_ACCEPT = &H8&
Private Const FD_CONNECT = &H10&
Private Const FD_CLOSE = &H20&

' Agregado por Maraxus
Private Const SOMAXCONN As Long = &H7FFFFFFF
Private SockReadBuffer As String

'---SOCKET FUNCTIONS
Private Declare Function API_accept Lib "ws2_32.DLL" Alias "WSAAccept" (ByVal SocketID As Long, pSockAddr As sockaddr, AddrLen As Long, ByVal lpfnCondition As Long, ByVal dwCallbackData As Long) As Long
Private Declare Function API_bind Lib "ws2_32.DLL" Alias "bind" (ByVal SocketID As Long, addr As sockaddr, ByVal namelen As Long) As Long
Private Declare Function API_closesocket Lib "ws2_32.DLL" Alias "closesocket" (ByVal SocketID As Long) As Long
Private Declare Function API_recv Lib "ws2_32.DLL" Alias "recv" (ByVal SocketID As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function API_listen Lib "ws2_32.DLL" Alias "listen" (ByVal SocketID As Long, ByVal backlog As Long) As Long
Private Declare Function API_connect Lib "ws2_32.DLL" Alias "connect" (ByVal SocketID As Long, addr As sockaddr, ByVal namelen As Long) As Long
Private Declare Function API_send Lib "ws2_32.DLL" Alias "send" (ByVal SocketID As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function API_setsockopt Lib "ws2_32.DLL" Alias "setsockopt" (ByVal SocketID As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function API_getsockopt Lib "ws2_32.DLL" (ByVal SocketID As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function API_ShutDown Lib "ws2_32.DLL" Alias "shutdown" (ByVal SocketID As Long, ByVal how As Long) As Long
Private Declare Function API_Socket Lib "ws2_32.DLL" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function API_inetaddr Lib "ws2_32.DLL" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function API_inet_ntoa Lib "ws2_32.DLL" Alias "inet_ntoa" (ByVal inn As Long) As Long
Private Declare Function API_htons Lib "ws2_32.DLL" Alias "htons" (ByVal hostshort As Long) As Integer
Private Declare Function API_ntohs Lib "ws2_32.DLL" Alias "ntohs" (ByVal netshort As Long) As Integer
Private Declare Function API_ntohl Lib "ws2_32.DLL" Alias "ntohl" (ByVal netlong As Long) As Long
Private Declare Function API_htonl Lib "ws2_32.DLL" Alias "ntohl" (ByVal hostlong As Long) As Long

'---DATABASE FUNCTIONS
Private Declare Function API_gethostbyname Lib "ws2_32.DLL" Alias "gethostbyname" (ByVal host_name As String) As Long

'---WINDOWS EXTENSIONS
Private Declare Function API_WSAAsyncSelect Lib "ws2_32.DLL" Alias "WSAAsyncSelect" (ByVal SocketID As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function API_WSAGetLastError Lib "ws2_32.DLL" Alias "WSAGetLastError" () As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

' Fake Window constants
Private Const WS_CHILD = &H40000000
Private Const GWL_WNDPROC = (-4)

' I/O Buffers
Private Const SIZE_RCVBUF        As Long = 10240    ' Recieved bytes
Private Const SIZE_SNDBUF        As Long = 10240    ' Sent bytes

Public Event Accept(ByVal SocketID As Long, ByVal UserIP As Long)
Public Event BeforeRead(ByVal Slot As Integer, ByRef BytesToRead As Long)
Public Event Read(ByVal Slot As Integer, ByRef Datos() As Byte, ByVal Length As Long)
Public Event Disconnect(ByVal SocketID As Integer, ByVal Slot As Integer)

' Class Properties
Private m_hWnd As Long
Private m_ListenerSocketID As Long

' Async messages handler
Private WithEvents Thread As clsSubclass
Attribute Thread.VB_VarHelpID = -1

Private Sub Thread_WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, Ret As Long, DefCall As Boolean)
    
    Dim ttt      As Long: ttt = GetTickCount()

    Dim api_return_value As Long
    Dim Tmp()           As Byte
    Dim SocketID        As Long
    Dim EventID         As Long
    Dim Slot            As Integer

    Select Case Msg

        Case 1025

            SocketID = wParam
            EventID = WSA_GetSelectEvent(lParam)
    
            Select Case EventID

                Case FD_ACCEPT

                    If SocketID = m_ListenerSocketID Then
                        Dim NuevoSocketID As Long
                        Dim SocketAddress As sockaddr
                        
                        Call Internal_SocketAccept(NuevoSocketID, SocketAddress)
                        RaiseEvent Accept(NuevoSocketID, SocketAddress.sin_addr)
                    End If
                    
                Case FD_READ
        
                    Slot = BuscaSlotSock(SocketID)

                    If Slot < 0 And SocketID <> m_ListenerSocketID Then
                        Call WSA_CloseSocket(SocketID)
                        Exit Sub

                    End If
        
                    ' WyroX: Leo hasta llenar el buffer, ni un byte más!!
                    Dim BytesToRead As Long
                    RaiseEvent BeforeRead(Slot, BytesToRead)
                    
                    If BytesToRead > 0 Then
                        ReDim Tmp(BytesToRead - 1) As Byte
            
                        api_return_value = API_recv(SocketID, Tmp(0), BytesToRead, 0)
    
                        RaiseEvent Read(Slot, Tmp, api_return_value)
                    End If
        
                Case FD_CLOSE
                    Slot = BuscaSlotSock(SocketID)

                    If SocketID <> m_ListenerSocketID Then _
                        Call API_closesocket(SocketID)
        
                    Call LogApiSock("WndProc:FD_CLOSE: Slot = " & Slot & " : Err = " & WSA_GetAsyncError(lParam))
        
                    RaiseEvent Disconnect(SocketID, Slot)
            End Select

        Case Else
            ' Jopi: !OJO! Esto esta puesto aca a proposito. NO LO SAQUES O SE ROMPE TODO
            DefCall = True

    End Select
    
End Sub

Public Property Get ListenerSocket() As Long
    ListenerSocket = m_ListenerSocketID
End Property

Public Function SendData(ByVal Slot As Integer, ByRef data() As Byte) As Long
        
    On Error GoTo WsApiEnviar_Err

    Dim Ret      As String
    Dim UltError As Long
    Dim Retorno  As Long

    Retorno = 0

    If UserList(Slot).ConnID <> -1 And UserList(Slot).ConnIDValida Then
        Ret = API_send(ByVal UserList(Slot).ConnID, data(0), ByVal UBound(data()) + 1, ByVal 0)

        If Ret < 0 Then
            UltError = Err.LastDllError
            
            ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
            If UltError = WSAEWOULDBLOCK Then
                Call UserList(Slot).outgoingData.WriteBlock(data)
            End If

            Retorno = UltError
        End If

    ElseIf UserList(Slot).ConnID <> -1 And Not UserList(Slot).ConnIDValida Then

        If Not UserList(Slot).Counters.Saliendo Then
            Retorno = -1
        End If

    End If

    SendData = Retorno

    Exit Function

WsApiEnviar_Err:
    Call TraceError(Err.Number, Err.Description, "wskapiAO.WsApiEnviar", Erl)
    Resume Next
        
End Function

Public Sub WSA_CloseSocket(ByVal SocketID As Long)
    
    ' Cierra la conexión con un cliente
    
    Call API_WSAAsyncSelect(SocketID, m_hWnd, ByVal 1025, ByVal (FD_CLOSE))
    Call API_ShutDown(SocketID, SD_BOTH)
             
End Sub

Public Sub ResetListenerSocket()
    
    Call API_closesocket(m_ListenerSocketID)
    
    ' Crear socket que escuche las peticiones de conexión entrantes
    m_ListenerSocketID = ListenForConnect(Puerto, m_hWnd, "")
    
End Sub

Private Sub Class_Initialize()
    
    Debug.Print "Iniciando Winsock API"
    
    m_hWnd = CreateWindowEx(0, "STATIC", "WINSOCK_HANDLER", WS_CHILD, 0, 0, 0, 0, frmMain.hWnd, 0, App.hInstance, ByVal 0&)
    
    Set Thread = New clsSubclass
    Call Thread.Hook(m_hWnd)
    
    ' ------------------------------------------------
    '   Winsock API startup
    ' ------------------------------------------------
    Call modWinsock.InitializeWinsockAPI
    
    ' Crear socket que escuche las peticiones de conexión entrantes
    m_ListenerSocketID = ListenForConnect(Puerto, m_hWnd, "")
    
End Sub

Private Sub Class_Terminate()
    
    ' ------------------------------------------------
    '   Winsock API de-initialization
    ' ------------------------------------------------
    Call modWinsock.DestroyWinsockAPI
    
    ' ------------------------------------------------
    '   END - Winsock API de-initialization
    ' ------------------------------------------------
    
    Set Thread = Nothing
    
    Call DestroyWindow(m_hWnd)
    
End Sub

Private Function Internal_SocketAccept(ByRef NuevoSocketID As Long, ByRef SocketAddress As sockaddr) As Long
    
    ' Se encarga de hacer los llamados a la Winsock API para aceptar la conexión entrante
    ' Devuelve el ID del Socket donde se encientra la conexion con el usuario.
    
    Dim i                   As Long
    Dim BufferSize          As Long: BufferSize = sockaddr_size
    Dim api_return_value    As Long
    
    'Modificado por Maraxus
    api_return_value = API_accept(m_ListenerSocketID, SocketAddress, BufferSize, AddressOf CondicionSocket, 0)
                
    If api_return_value = INVALID_SOCKET Then
        i = Err.LastDllError
        Call LogCriticEvent("Error en Accept() API " & i & ": " & WSA_GetErrorString(i))
        Exit Function
                
    End If
                
    NuevoSocketID = api_return_value
                
    'Call setsockopt(wsock.SocketHandle, 6, 1, True, 4)
    'old: If setsockopt(NuevoSocketID, SOL_SOCKET, TCP_NODELAY, True, 1) <> 0 Then
                    
    'algoritmo de nagle vb6
    'If setsockopt(NuevoSocketID, 6, 1, True, 4) <> 0 Then
                             
   ' If API_setsockopt(NuevoSocketID, 6, TCP_NODELAY, True, 4) <> 0 Then
   '     i = Err.LastDllError
   '     Call LogCriticEvent("Error al setear el delay " & i & ": " & WSA_GetErrorString(i))
   '
   ' End If
  
    'Nuevo sin nagle
    If API_setsockopt(NuevoSocketID, SOL_SOCKET, SO_LINGER, 0, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear lingers." & i & ": " & WSA_GetErrorString(i))
                
    End If
                    
    'Seteamos el tamaño del buffer de entrada
    If API_setsockopt(NuevoSocketID, SOL_SOCKET, SO_RCVBUFFER, SIZE_RCVBUF, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tamaño del buffer de entrada " & i & ": " & WSA_GetErrorString(i))
                
    End If
                
    'Seteamos el tamaño del buffer de salida
    If API_setsockopt(NuevoSocketID, SOL_SOCKET, SO_SNDBUFFER, SIZE_SNDBUF, 4) <> 0 Then
        i = Err.LastDllError
        Call LogCriticEvent("Error al setear el tamaño del buffer de salida " & i & ": " & WSA_GetErrorString(i))
                
    End If
    
    Internal_SocketAccept = NuevoSocketID
    
End Function

Private Function WSA_GetSelectEvent(ByVal lParam As Long) As Integer
    
    ' Determina el tipo de mensaje a procesar en WndProc
    
    If (lParam And &HFFFF&) > &H7FFF Then
        WSA_GetSelectEvent = (lParam And &HFFFF&) - &H10000
        
    Else
        WSA_GetSelectEvent = (lParam And &HFFFF&)
        
    End If

End Function

Private Function WSA_GetAsyncError(ByVal lParam As Long) As Integer

    WSA_GetAsyncError = (lParam And &HFFFF0000) \ &H10000

End Function

Private Function ConnectSock(ByVal Host As String, ByVal Port As Long, retIpPort As String, ByVal HWndToMsg As Long, ByVal Async%) As Long
        
    On Error GoTo 0
    
    Dim SocketID As Long, SelectOps As Long
    Dim sockin As sockaddr

    SockReadBuffer = vbNullString

    sockin = saZero
    sockin.sin_family = AF_INET
    sockin.sin_port = API_htons(Port)

    If sockin.sin_port = INVALID_SOCKET Then
        ConnectSock = INVALID_SOCKET
        Exit Function

    End If

    sockin.sin_addr = GetHostByNameAlias(Host)

    If sockin.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        Exit Function

    End If

    retIpPort = GetAscIP(sockin.sin_addr) & ":" & API_ntohs(sockin.sin_port)

    SocketID = API_Socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    
    If SocketID < 0 Then
        ConnectSock = INVALID_SOCKET
        Exit Function

    End If

    If SetSockLinger(SocketID, 1, 0) = SOCKET_ERROR Then
        
        If SocketID > 0 Then _
            Call API_closesocket(SocketID)

        ConnectSock = INVALID_SOCKET
        Exit Function

    End If

    If Not Async Then
    
        If Not API_connect(SocketID, sockin, sockaddr_size) = 0 Then
            
            If SocketID > 0 Then _
                Call API_closesocket(SocketID)
                
            ConnectSock = INVALID_SOCKET
            Exit Function

        End If

        If HWndToMsg <> 0 Then
        
            SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE

            If API_WSAAsyncSelect(SocketID, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
            
                If SocketID > 0 Then _
                    Call API_closesocket(SocketID)

                ConnectSock = INVALID_SOCKET
                Exit Function

            End If

        End If

    Else
    
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE

        If API_WSAAsyncSelect(SocketID, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
        
            If SocketID > 0 Then _
                Call API_closesocket(SocketID)

            ConnectSock = INVALID_SOCKET
            Exit Function

        End If

        If API_connect(SocketID, sockin, sockaddr_size) <> -1 Then
            
            If SocketID > 0 Then _
                Call API_closesocket(SocketID)

            ConnectSock = INVALID_SOCKET
            Exit Function

        End If

    End If

    ConnectSock = SocketID
        
End Function

Private Function ListenForConnect(ByVal Port As Long, ByVal HWndToMsg As Long, ByVal Enlazar As String) As Long

    On Error GoTo 0
    
    Dim SocketID As Long
    Dim SelectOps As Long
    Dim sockin As sockaddr

    sockin = saZero     'zero out the structure
    sockin.sin_family = AF_INET
    sockin.sin_port = API_htons(Port)

    If sockin.sin_port = INVALID_SOCKET Then
        ListenForConnect = INVALID_SOCKET
        Exit Function

    End If

    If LenB(Enlazar) = 0 Then
        sockin.sin_addr = API_htonl(INADDR_ANY)
        
    Else
        sockin.sin_addr = API_inetaddr(Enlazar)

    End If

    If sockin.sin_addr = INADDR_NONE Then
        ListenForConnect = INVALID_SOCKET
        Exit Function

    End If

    SocketID = API_Socket(PF_INET, SOCK_STREAM, 0)

    If SocketID < 0 Then
        ListenForConnect = INVALID_SOCKET
        Exit Function

    End If
    
    'Agregado por Maraxus
    'If API_SetSockOpt(SocketID, SOL_SOCKET, SO_CONDITIONAL_ACCEPT, True, 2) Then
    '    LogApiSock ("Error seteando conditional accept")
    '    Debug.Print "Error seteando conditional accept"
    'Else
    '    LogApiSock ("Conditional accept seteado")
    '    Debug.Print "Conditional accept seteado ^^"
    'End If
    
    If API_bind(SocketID, sockin, sockaddr_size) Then
    
        If SocketID > 0 Then _
            Call API_closesocket(SocketID)

        ListenForConnect = INVALID_SOCKET
        Exit Function

    End If

    SelectOps = FD_READ Or FD_CLOSE Or FD_ACCEPT

    If API_WSAAsyncSelect(SocketID, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
    
        If SocketID > 0 Then _
            Call API_closesocket(SocketID)

        ListenForConnect = SOCKET_ERROR
        Exit Function

    End If
    
    If API_listen(SocketID, SOMAXCONN) Then
    
        If SocketID > 0 Then _
            Call API_closesocket(SocketID)

        ListenForConnect = INVALID_SOCKET
        Exit Function

    End If

    ListenForConnect = SocketID

End Function

Private Function SetSockLinger(ByVal SockNum As Long, ByVal OnOff As Integer, ByVal LingerTime As Integer) As Long
        
    On Error GoTo 0
    
    Dim Linger As LingerType

    Linger.l_onoff = OnOff
    Linger.l_linger = LingerTime

    If API_setsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
        Debug.Print "Error setting linger info: " & API_WSAGetLastError()
        SetSockLinger = SOCKET_ERROR
    
    Else

        If API_getsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
            Debug.Print "Error getting linger info: " & API_WSAGetLastError()
            SetSockLinger = SOCKET_ERROR
            
        Else
            Debug.Print "Linger is on if nonzero: "; Linger.l_onoff
            Debug.Print "Linger time if linger is on: "; Linger.l_linger

        End If

    End If

End Function

'returns IP as long, in network byte order
Private Function GetHostByNameAlias(ByVal HostName As String) As Long
        
    On Error GoTo 0

    'Return IP address as a long, in network byte order
    Dim phe As Long
    Dim heDestHost As HostEnt
    Dim addrList As Long
    Dim retIP As Long

    retIP = API_inetaddr(HostName)

    If retIP = INADDR_NONE Then
        phe = API_gethostbyname(HostName)

        If phe <> 0 Then
        
            Call API_MemCopy(heDestHost, ByVal phe, hostent_size)
            Call API_MemCopy(addrList, ByVal heDestHost.h_addr_list, 4)
            Call API_MemCopy(retIP, ByVal addrList, heDestHost.h_length)
            
        Else
        
            retIP = INADDR_NONE

        End If

    End If

    GetHostByNameAlias = retIP

End Function

Public Function WSA_GetErrorString(ByVal ErrorNumber As Long) As String
        
    On Error GoTo 0

    Select Case ErrorNumber

        Case 10004
            WSA_GetErrorString = "Interrupted system call."

        Case 10009
            WSA_GetErrorString = "Bad file number."

        Case 10013
            WSA_GetErrorString = "Permission Denied."

        Case 10014
            WSA_GetErrorString = "Bad Address."

        Case 10022
            WSA_GetErrorString = "Invalid Argument."

        Case 10024
            WSA_GetErrorString = "Too many open files."

        Case 10035
            WSA_GetErrorString = "Operation would block."

        Case 10036
            WSA_GetErrorString = "Operation now in progress."

        Case 10037
            WSA_GetErrorString = "Operation already in progress."

        Case 10038
            WSA_GetErrorString = "Socket operation on nonsocket."

        Case 10039
            WSA_GetErrorString = "Destination address required."

        Case 10040
            WSA_GetErrorString = "Message too long."

        Case 10041
            WSA_GetErrorString = "Protocol wrong type for socket."

        Case 10042
            WSA_GetErrorString = "Protocol not available."

        Case 10043
            WSA_GetErrorString = "Protocol not supported."

        Case 10044
            WSA_GetErrorString = "Socket type not supported."

        Case 10045
            WSA_GetErrorString = "Operation not supported on socket."

        Case 10046
            WSA_GetErrorString = "Protocol family not supported."

        Case 10047
            WSA_GetErrorString = "Address family not supported by protocol family."

        Case 10048
            WSA_GetErrorString = "Address already in use."

        Case 10049
            WSA_GetErrorString = "Can't assign requested address."

        Case 10050
            WSA_GetErrorString = "Network is down."

        Case 10051
            WSA_GetErrorString = "Network is unreachable."

        Case 10052
            WSA_GetErrorString = "Network dropped connection."

        Case 10053
            WSA_GetErrorString = "Software caused connection abort."

        Case 10054
            WSA_GetErrorString = "Connection reset by peer."

        Case 10055
            WSA_GetErrorString = "No buffer space available."

        Case 10056
            WSA_GetErrorString = "Socket is already connected."

        Case 10057
            WSA_GetErrorString = "Socket is not connected."

        Case 10058
            WSA_GetErrorString = "Can't send after socket shutdown."

        Case 10059
            WSA_GetErrorString = "Too many references: can't splice."

        Case 10060
            WSA_GetErrorString = "Connection timed out."

        Case 10061
            WSA_GetErrorString = "Connection refused."

        Case 10062
            WSA_GetErrorString = "Too many levels of symbolic links."

        Case 10063
            WSA_GetErrorString = "File name too long."

        Case 10064
            WSA_GetErrorString = "Host is down."

        Case 10065
            WSA_GetErrorString = "No route to host."

        Case 10066
            WSA_GetErrorString = "Directory not empty."

        Case 10067
            WSA_GetErrorString = "Too many processes."

        Case 10068
            WSA_GetErrorString = "Too many users."

        Case 10069
            WSA_GetErrorString = "Disk quota exceeded."

        Case 10070
            WSA_GetErrorString = "Stale NFS file handle."

        Case 10071
            WSA_GetErrorString = "Too many levels of remote in path."

        Case 10091
            WSA_GetErrorString = "Network subsystem is unusable."

        Case 10092
            WSA_GetErrorString = "Winsock DLL cannot support this application."

        Case 10093
            WSA_GetErrorString = "Winsock not initialized."

        Case 10101
            WSA_GetErrorString = "Disconnect."

        Case 11001
            WSA_GetErrorString = "Host not found."

        Case 11002
            WSA_GetErrorString = "Nonauthoritative host not found."

        Case 11003
            WSA_GetErrorString = "Nonrecoverable error."

        Case 11004
            WSA_GetErrorString = "Valid name, no data record of requested type."

        Case Else:

    End Select

End Function
