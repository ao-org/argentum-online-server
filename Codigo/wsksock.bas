Attribute VB_Name = "WSKSOCK"
'**************************************************************
' WSKSOCK.bas
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


    'date stamp: sept 1, 1996 (for version control, please don't remove)

    'Visual Basic 4.0 Winsock "Header"
    '   Alot of the information contained inside this file was originally
    '   obtained from ALT.WINSOCK.PROGRAMMING and most of it has since been
    '   modified in some way.
    '
    'Disclaimer: This file is public domain, updated periodically by
    '   Topaz, SigSegV@mail.utexas.edu, Use it at your own risk.
    '   Neither myself(Topaz) or anyone related to alt.programming.winsock
    '   may be held liable for its use, or misuse.
    '
    'Declare check Aug 27, 1996. (Topaz, SigSegV@mail.utexas.edu)
    '   All 16 bit declarations appear correct, even the odd ones that
    '   pass longs inplace of in_addr and char buffers. 32 bit functions
    '   also appear correct. Some are declared to return integers instead of
    '   longs (breaking MS's rules.) however after testing these functions I
    '   have come to the conclusion that they do not work properly when declared
    
    '   following MS's rules.
    '
    'NOTES:
    '   (1) I have never used WS_SELECT (select), therefore I must warn that I do
    '       not know if fd_set and timeval are properly defined.
    '   (2) Alot of the functions are declared with "buf as any", when calling these
    '       functions you may either pass strings, byte arrays or UDT's. For 32bit I
    '       I recommend Byte arrays and the use of memcopy to copy the data back out
    '   (3) The async functions (wsaAsync*) require the use of a message hook or
    '       message window control to capture messages sent by the winsock stack. This
    '       is not to be confused with a CallBack control, The only function that uses
    '       callbacks is WSASetBlockingHook()
    '   (4) Alot of "helper" functions are provided in the file for various things
    '       before attempting to figure out how to call a function, look and see if
    '       there is already a helper function for it.
    '   (5) Data types (hostent etc) have kept there 16bit definitions, even under 32bit
    '       windows due to the problem of them not working when redfined following the
    '       suggested rules.
    Option Explicit

    Public Const FD_SETSIZE = 64

    Type fd_set
        fd_count As Integer
        fd_array(FD_SETSIZE) As Integer
    End Type

    Type timeval
        tv_sec As Long
        tv_usec As Long
    End Type

    Type HostEnt
        h_name As Long
        h_aliases As Long
        h_addrtype As Integer
        h_length As Integer
        h_addr_list As Long
    End Type

    Public Const hostent_size = 16

    Type servent
        s_name As Long
        s_aliases As Long
        s_port As Integer
        s_proto As Long
    End Type

    Public Const servent_size = 14

    Type protoent
        p_name As Long
        p_aliases As Long
        p_proto As Integer
    End Type

    Public Const protoent_size = 10
    
    Public Const IPPROTO_TCP = 6
    Public Const IPPROTO_UDP = 17
    
    Public Const INADDR_NONE = &HFFFFFFFF
    Public Const INADDR_ANY = &H0

    Type sockaddr
        sin_family As Integer
        sin_port As Integer
        sin_addr As Long
        sin_zero As String * 8
    End Type

    Public Const sockaddr_size = 16

    Public saZero As sockaddr

    Public Const WSA_DESCRIPTIONLEN = 256
    Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
    
    Public Const WSA_SYS_STATUS_LEN = 128
    Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

    Type WSADataType
        wVersion As Integer
        wHighVersion As Integer
        szDescription As String * WSA_DescriptionSize
        szSystemStatus As String * WSA_SysStatusSize
        iMaxSockets As Integer
        iMaxUdpDg As Integer
        lpVendorInfo As Long
    End Type

    'Agregado por Maraxus
    Type WSABUF
        dwBufferLen As Long
        LpBuffer    As Long
    End Type

    'Agregado por Maraxus
    Type FLOWSPEC
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
    Public Const WSA_FLAG_OVERLAPPED = &H1

    'Agregados por Maraxus
    Public Const CF_ACCEPT = &H0
    Public Const CF_REJECT = &H1

    'Agregado por Maraxus
    Public Const SD_RECEIVE As Long = &H0&
    Public Const SD_SEND    As Long = &H1&
    Public Const SD_BOTH    As Long = &H2&

    Public Const INVALID_SOCKET = -1
    Public Const SOCKET_ERROR = -1

    Public Const SOCK_STREAM = 1
    Public Const SOCK_DGRAM = 2

    Public Const MAXGETHOSTSTRUCT = 1024

    Public Const AF_INET = 2
    Public Const PF_INET = 2

    Type LingerType
        l_onoff As Integer
        l_linger As Integer
    End Type

    ' Windows Sockets definitions of regular Microsoft C error constants
    Global Const WSAEINTR = 10004
    Global Const WSAEBADF = 10009
    Global Const WSAEACCES = 10013
    Global Const WSAEFAULT = 10014
    Global Const WSAEINVAL = 10022
    Global Const WSAEMFILE = 10024

    ' Windows Sockets definitions of regular Berkeley error constants
    Global Const WSAEWOULDBLOCK = 10035
    Global Const WSAEINPROGRESS = 10036
    Global Const WSAEALREADY = 10037
    Global Const WSAENOTSOCK = 10038
    Global Const WSAEDESTADDRREQ = 10039
    Global Const WSAEMSGSIZE = 10040
    Global Const WSAEPROTOTYPE = 10041
    Global Const WSAENOPROTOOPT = 10042
    Global Const WSAEPROTONOSUPPORT = 10043
    Global Const WSAESOCKTNOSUPPORT = 10044
    Global Const WSAEOPNOTSUPP = 10045
    Global Const WSAEPFNOSUPPORT = 10046
    Global Const WSAEAFNOSUPPORT = 10047
    Global Const WSAEADDRINUSE = 10048
    Global Const WSAEADDRNOTAVAIL = 10049
    Global Const WSAENETDOWN = 10050
    Global Const WSAENETUNREACH = 10051
    Global Const WSAENETRESET = 10052
    Global Const WSAECONNABORTED = 10053
    Global Const WSAECONNRESET = 10054
    Global Const WSAENOBUFS = 10055
    Global Const WSAEISCONN = 10056
    Global Const WSAENOTCONN = 10057
    Global Const WSAESHUTDOWN = 10058
    Global Const WSAETOOMANYREFS = 10059
    Global Const WSAETIMEDOUT = 10060
    Global Const WSAECONNREFUSED = 10061
    Global Const WSAELOOP = 10062
    Global Const WSAENAMETOOLONG = 10063
    Global Const WSAEHOSTDOWN = 10064
    Global Const WSAEHOSTUNREACH = 10065
    Global Const WSAENOTEMPTY = 10066
    Global Const WSAEPROCLIM = 10067
    Global Const WSAEUSERS = 10068
    Global Const WSAEDQUOT = 10069
    Global Const WSAESTALE = 10070
    Global Const WSAEREMOTE = 10071

    ' Extended Windows Sockets error constant definitions
    Global Const WSASYSNOTREADY = 10091
    Global Const WSAVERNOTSUPPORTED = 10092
    Global Const WSANOTINITIALISED = 10093
    Global Const WSAHOST_NOT_FOUND = 11001
    Global Const WSATRY_AGAIN = 11002
    Global Const WSANO_RECOVERY = 11003
    Global Const WSANO_DATA = 11004
    Global Const WSANO_ADDRESS = 11004

    '---ioctl Constants
    Public Const FIONREAD = &H8004667F
    Public Const FIONBIO = &H8004667E
    Public Const FIOASYNC = &H8004667D

    '---Windows System Functions
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
    Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

    '---async notification constants
    Public Const TCP_NODELAY = &H1                 ' Agregado por Maraxus
    Public Const SOL_SOCKET = &HFFFF&
    Public Const SO_LINGER = &H80&
    Public Const SO_RCVBUFFER = &H1002&             ' Agregado por Maraxus
    Public Const SO_SNDBUFFER = &H1001&              ' Agregado por Maraxus
    Public Const SO_CONDITIONAL_ACCEPT = &H3002&    ' Agregado por Maraxus
    Public Const FD_READ = &H1&
    Public Const FD_WRITE = &H2&
    Public Const FD_OOB = &H4&
    Public Const FD_ACCEPT = &H8&
    Public Const FD_CONNECT = &H10&
    Public Const FD_CLOSE = &H20&

    '---SOCKET FUNCTIONS
    Public Declare Function accept Lib "wsock32.dll" (ByVal S As Long, addr As sockaddr, AddrLen As Long) As Long
    Public Declare Function bind Lib "wsock32.dll" (ByVal S As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Public Declare Function apiclosesocket Lib "wsock32.dll" Alias "closesocket" (ByVal S As Long) As Long
    Public Declare Function connect Lib "wsock32.dll" (ByVal S As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Public Declare Function ioctlsocket Lib "wsock32.dll" (ByVal S As Long, ByVal cmd As Long, argp As Long) As Long
    Public Declare Function getpeername Lib "wsock32.dll" (ByVal S As Long, sName As sockaddr, namelen As Long) As Long
    Public Declare Function getsockname Lib "wsock32.dll" (ByVal S As Long, sName As sockaddr, namelen As Long) As Long
    Public Declare Function getsockopt Lib "wsock32.dll" (ByVal S As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
    Public Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long
    Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
    Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
    Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
    Public Declare Function listen Lib "wsock32.dll" (ByVal S As Long, ByVal backlog As Long) As Long
    Public Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long
    Public Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
    Public Declare Function recv Lib "wsock32.dll" (ByVal S As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Public Declare Function recvfrom Lib "wsock32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, from As sockaddr, fromlen As Long) As Long
    Public Declare Function ws_select Lib "wsock32.dll" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Long
    Public Declare Function send Lib "wsock32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    Public Declare Function sendto Lib "wsock32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, to_addr As sockaddr, ByVal tolen As Long) As Long
    Public Declare Function setsockopt Lib "wsock32.dll" (ByVal S As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
    Public Declare Function ShutDown Lib "wsock32.dll" Alias "shutdown" (ByVal S As Long, ByVal how As Long) As Long
    Public Declare Function Socket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long

    '---DATABASE FUNCTIONS
    Public Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
    Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
    Public Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
    Public Declare Function getservbyport Lib "wsock32.dll" (ByVal Port As Long, ByVal proto As String) As Long
    Public Declare Function getservbyname Lib "wsock32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
    Public Declare Function getprotobynumber Lib "wsock32.dll" (ByVal proto As Long) As Long
    Public Declare Function getprotobyname Lib "wsock32.dll" (ByVal proto_name As String) As Long

    '---WINDOWS EXTENSIONS
    Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
    Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
    Public Declare Sub WSASetLastError Lib "wsock32.dll" (ByVal iError As Long)
    Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
    Public Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
    Public Declare Function WSAUnhookBlockingHook Lib "wsock32.dll" () As Long
    Public Declare Function WSASetBlockingHook Lib "wsock32.dll" (ByVal lpBlockFunc As Long) As Long
    Public Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
    Public Declare Function WSAAsyncGetServByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetServByPort Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal proto_name As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByNumber Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Number As Long, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal host_name As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByAddr Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, addr As Long, ByVal addr_len As Long, ByVal addr_type As Long, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSACancelAsyncRequest Lib "wsock32.dll" (ByVal hAsyncTaskHandle As Long) As Long
    Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal S As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
    Public Declare Function WSARecvEx Lib "wsock32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
    
    'Agregado por Maraxus
    Declare Function WSAAccept Lib "ws2_32.DLL" (ByVal S As Long, pSockAddr As sockaddr, AddrLen As Long, ByVal lpfnCondition As Long, ByVal dwCallbackData As Long) As Long

    ' Agregado por Maraxus
    Public Const SOMAXCONN As Long = &H7FFFFFFF
    
    'SOME STUFF I ADDED
    Public MySocket%
    Public SockReadBuffer$
    Public Const WSA_NoName = "Unknown"
    Public WSAStartedUp As Boolean     'Flag to keep track of whether winsock WSAStartup wascalled

Public Function WSAGetAsyncBufLen(ByVal lParam As Long) As Long
        
        On Error GoTo 0
    
        

100     If (lParam And &HFFFF&) > &H7FFF Then
102         WSAGetAsyncBufLen = (lParam And &HFFFF&) - &H10000
        Else
104         WSAGetAsyncBufLen = lParam And &HFFFF&
        End If

        
        Exit Function

WSAGetAsyncBufLen_Err:
106     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.WSAGetAsyncBufLen", Erl)

        
End Function

Public Function WSAGetSelectEvent(ByVal lParam As Long) As Integer
        
        On Error GoTo 0
    
        
        
100     If (lParam And &HFFFF&) > &H7FFF Then
102         WSAGetSelectEvent = (lParam And &HFFFF&) - &H10000
        Else
104         WSAGetSelectEvent = lParam And &HFFFF&
        End If
        
        
        Exit Function

WSAGetSelectEvent_Err:
106     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.WSAGetSelectEvent", Erl)

        
End Function

Public Function WSAGetAsyncError(ByVal lParam As Long) As Integer
        
        On Error GoTo 0
    
        

100     WSAGetAsyncError = (lParam And &HFFFF0000) \ &H10000
        
        
        Exit Function

WSAGetAsyncError_Err:
102     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.WSAGetAsyncError", Erl)

        
End Function

Public Function AddrToIP(ByVal AddrOrIP$) As String
        
        On Error GoTo 0
    
        
    
        Dim T() As String
        Dim Tmp As String

100     Tmp = GetAscIP(GetHostByNameAlias(AddrOrIP$))
102     T = Split(Tmp, ".")
104     AddrToIP = T(3) & "." & T(2) & "." & T(1) & "." & T(0)
        
        
        Exit Function

AddrToIP_Err:
106     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.AddrToIP", Erl)

        
End Function

'this function should work on 16 and 32 bit systems
Function ConnectSock(ByVal Host$, ByVal Port&, retIpPort$, ByVal HWndToMsg&, ByVal Async%) As Long
        
        On Error GoTo 0
    
        

        Dim S&, SelectOps&, dummy&
        Dim sockin As sockaddr

100     SockReadBuffer$ = vbNullString
102     sockin = saZero
104     sockin.sin_family = AF_INET
106     sockin.sin_port = htons(Port)

108     If sockin.sin_port = INVALID_SOCKET Then
110         ConnectSock = INVALID_SOCKET
            Exit Function

        End If

112     sockin.sin_addr = GetHostByNameAlias(Host$)

114     If sockin.sin_addr = INADDR_NONE Then
116         ConnectSock = INVALID_SOCKET
            Exit Function

        End If

118     retIpPort$ = GetAscIP$(sockin.sin_addr) & ":" & ntohs(sockin.sin_port)

120     S = Socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)

122     If S < 0 Then
124         ConnectSock = INVALID_SOCKET
            Exit Function

        End If

126     If SetSockLinger(S, 1, 0) = SOCKET_ERROR Then
128         If S > 0 Then
130             dummy = apiclosesocket(S)

            End If

132         ConnectSock = INVALID_SOCKET
            Exit Function

        End If

134     If Not Async Then
136         If Not connect(S, sockin, sockaddr_size) = 0 Then
138             If S > 0 Then
140                 dummy = apiclosesocket(S)

                End If

142             ConnectSock = INVALID_SOCKET
                Exit Function

            End If

144         If HWndToMsg <> 0 Then
146             SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE

148             If WSAAsyncSelect(S, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
150                 If S > 0 Then
152                     dummy = apiclosesocket(S)

                    End If

154                 ConnectSock = INVALID_SOCKET
                    Exit Function

                End If

            End If

        Else
    
156         SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE

158         If WSAAsyncSelect(S, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
160             If S > 0 Then
162                 dummy = apiclosesocket(S)

                End If

164             ConnectSock = INVALID_SOCKET
                Exit Function

            End If

166         If connect(S, sockin, sockaddr_size) <> -1 Then
168             If S > 0 Then
170                 dummy = apiclosesocket(S)

                End If

172             ConnectSock = INVALID_SOCKET
                Exit Function

            End If

        End If

174     ConnectSock = S
    
        
        Exit Function

ConnectSock_Err:
176     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.ConnectSock", Erl)

        
End Function

Public Function SetSockLinger(ByVal SockNum&, ByVal OnOff%, ByVal LingerTime%) As Long
        
        On Error GoTo 0
    
        
    
        Dim Linger As LingerType

100     Linger.l_onoff = OnOff
102     Linger.l_linger = LingerTime

104     If setsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
106         Debug.Print "Error setting linger info: " & WSAGetLastError()
108         SetSockLinger = SOCKET_ERROR
    
        Else

110         If getsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
112             Debug.Print "Error getting linger info: " & WSAGetLastError()
114             SetSockLinger = SOCKET_ERROR
            Else
116             Debug.Print "Linger is on if nonzero: "; Linger.l_onoff
118             Debug.Print "Linger time if linger is on: "; Linger.l_linger

            End If

        End If

        
        Exit Function

SetSockLinger_Err:
120     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.SetSockLinger", Erl)

        
End Function

Sub EndWinsock()
        
        On Error GoTo 0
    
        
        
        Dim ret&

100     If WSAIsBlocking() Then
102         ret = WSACancelBlockingCall()
        End If

104     ret = WSACleanup()
    
106     WSAStartedUp = False
           
        
        Exit Sub

EndWinsock_Err:
108     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.EndWinsock", Erl)

        
End Sub

Public Function GetAscIP(ByVal inn As Long) As String
        
        On Error GoTo 0
    
        

        Dim nStr&
        Dim lpStr&

        Dim retString$

100     retString = String(32, 0)
102     lpStr = inet_ntoa(inn)

104     If lpStr Then
    
106         nStr = lstrlen(lpStr)

108         If nStr > 32 Then nStr = 32
        
110         Call MemCopy(ByVal retString, ByVal lpStr, nStr)
        
112         retString = Left$(retString, nStr)
114         GetAscIP = retString
        
        Else
    
116         GetAscIP = "255.255.255.255"

        End If

        
        Exit Function

GetAscIP_Err:
118     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.GetAscIP", Erl)

        
End Function

Public Function GetHostByAddress(ByVal addr As Long) As String
        
        On Error GoTo 0
    
        

        Dim phe&
        Dim heDestHost As HostEnt
        Dim HostName$

100     phe = gethostbyaddr(addr, 4, PF_INET)

102     If phe Then
        
104         Call MemCopy(heDestHost, ByVal phe, hostent_size)
        
106         HostName = String(256, 0)
        
108         Call MemCopy(ByVal HostName, ByVal heDestHost.h_name, 256)
        
110         GetHostByAddress = Left$(HostName, InStr(HostName, Chr$(0)) - 1)
    
        Else
    
112         GetHostByAddress = WSA_NoName

        End If

        
        Exit Function

GetHostByAddress_Err:
114     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.GetHostByAddress", Erl)

        
End Function

'returns IP as long, in network byte order
Public Function GetHostByNameAlias(ByVal HostName$) As Long
        
        On Error GoTo 0
    
        

        'Return IP address as a long, in network byte order
        Dim phe&
        Dim heDestHost As HostEnt
        Dim addrList&
        Dim retIP&

100     retIP = inet_addr(HostName$)

102     If retIP = INADDR_NONE Then
104         phe = gethostbyname(HostName$)

106         If phe <> 0 Then
108             Call MemCopy(heDestHost, ByVal phe, hostent_size)
110             Call MemCopy(addrList, ByVal heDestHost.h_addr_list, 4)
112             Call MemCopy(retIP, ByVal addrList, heDestHost.h_length)
            
            Else
114             retIP = INADDR_NONE

            End If

        End If

116     GetHostByNameAlias = retIP
        
        
        Exit Function

GetHostByNameAlias_Err:
118     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.GetHostByNameAlias", Erl)

        
End Function

'returns your local machines name
Public Function GetLocalHostName() As String
        
        On Error GoTo 0
    
        

        Dim sName$

100     sName = String(256, 0)

102     If gethostname(sName, 256) Then
104         sName = WSA_NoName
    
        Else

106         If InStr(sName, Chr$(0)) Then
108             sName = Left$(sName, InStr(sName, Chr$(0)) - 1)
            End If

        End If

110     GetLocalHostName = sName

        
        Exit Function

GetLocalHostName_Err:
112     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.GetLocalHostName", Erl)

        
End Function

Public Function GetPeerAddress(ByVal S&) As String
        
        On Error GoTo 0
    
        

        Dim AddrLen&
        Dim sa As sockaddr

100     AddrLen = sockaddr_size

102     If getpeername(S, sa, AddrLen) Then
104         GetPeerAddress = vbNullString
    
        Else
106         GetPeerAddress = SockAddressToString(sa)

        End If
        
        
        Exit Function

GetPeerAddress_Err:
108     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.GetPeerAddress", Erl)

        
End Function

Public Function GetPortFromString(ByVal PortStr$) As Long
        
        On Error GoTo 0
    
        

        'sometimes users provide ports outside the range of a VB
        'integer, so this function returns an integer for a string
        'just to keep an error from happening, it converts the
        'number to a negative if needed
100     If val(PortStr$) > 32767 Then
102         GetPortFromString = CInt(val(PortStr$) - &H10000)
    
        Else
104         GetPortFromString = val(PortStr$)

        End If

106     If Err Then GetPortFromString = 0
        
        
        Exit Function

GetPortFromString_Err:
108     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.GetPortFromString", Erl)

        
End Function

Function GetProtocolByName(ByVal Protocol$) As Long
        
        On Error GoTo 0
    
        

        Dim tmpShort&
        Dim ppe&
        Dim peDestProt As protoent

100     ppe = getprotobyname(Protocol)

102     If ppe Then
        
104         Call MemCopy(peDestProt, ByVal ppe, protoent_size)
        
106         GetProtocolByName = peDestProt.p_proto
    
        Else
    
108         tmpShort = val(Protocol)

110         If tmpShort Then
112             GetProtocolByName = htons(tmpShort)
        
            Else
114             GetProtocolByName = SOCKET_ERROR

            End If

        End If

        
        Exit Function

GetProtocolByName_Err:
116     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.GetProtocolByName", Erl)

        
End Function

Function GetServiceByName(ByVal service$, ByVal Protocol$) As Long
        
        On Error GoTo 0
    
        

        Dim Serv&
        Dim pse&
        Dim seDestServ As servent

100     pse = getservbyname(service, Protocol)

102     If pse Then
        
104         Call MemCopy(seDestServ, ByVal pse, servent_size)
        
106         GetServiceByName = seDestServ.s_port
    
        Else
        
108         Serv = val(service)

110         If Serv Then
112             GetServiceByName = htons(Serv)
            Else
114             GetServiceByName = INVALID_SOCKET

            End If

        End If

        
        Exit Function

GetServiceByName_Err:
116     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.GetServiceByName", Erl)

        
End Function

Function GetSockAddress(ByVal S&) As String
        
        On Error GoTo 0
    
        

        Dim AddrLen&
        Dim ret&
        Dim sa As sockaddr
        Dim szRet$

100     szRet = String(32, 0)
102     AddrLen = sockaddr_size

104     If getsockname(S, sa, AddrLen) Then
106         GetSockAddress = vbNullString
    
        Else
108         GetSockAddress = SockAddressToString(sa)

        End If

        
        Exit Function

GetSockAddress_Err:
110     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.GetSockAddress", Erl)

        
End Function

'this function should work on 16 and 32 bit systems
Function GetWSAErrorString(ByVal errnum&) As String
        
        On Error GoTo 0
    
        

        

100     Select Case errnum

            Case 10004
102             GetWSAErrorString = "Interrupted system call."

104         Case 10009
106             GetWSAErrorString = "Bad file number."

108         Case 10013
110             GetWSAErrorString = "Permission Denied."

112         Case 10014
114             GetWSAErrorString = "Bad Address."

116         Case 10022
118             GetWSAErrorString = "Invalid Argument."

120         Case 10024
122             GetWSAErrorString = "Too many open files."

124         Case 10035
126             GetWSAErrorString = "Operation would block."

128         Case 10036
130             GetWSAErrorString = "Operation now in progress."

132         Case 10037
134             GetWSAErrorString = "Operation already in progress."

136         Case 10038
138             GetWSAErrorString = "Socket operation on nonsocket."

140         Case 10039
142             GetWSAErrorString = "Destination address required."

144         Case 10040
146             GetWSAErrorString = "Message too long."

148         Case 10041
150             GetWSAErrorString = "Protocol wrong type for socket."

152         Case 10042
154             GetWSAErrorString = "Protocol not available."

156         Case 10043
158             GetWSAErrorString = "Protocol not supported."

160         Case 10044
162             GetWSAErrorString = "Socket type not supported."

164         Case 10045
166             GetWSAErrorString = "Operation not supported on socket."

168         Case 10046
170             GetWSAErrorString = "Protocol family not supported."

172         Case 10047
174             GetWSAErrorString = "Address family not supported by protocol family."

176         Case 10048
178             GetWSAErrorString = "Address already in use."

180         Case 10049
182             GetWSAErrorString = "Can't assign requested address."

184         Case 10050
186             GetWSAErrorString = "Network is down."

188         Case 10051
190             GetWSAErrorString = "Network is unreachable."

192         Case 10052
194             GetWSAErrorString = "Network dropped connection."

196         Case 10053
198             GetWSAErrorString = "Software caused connection abort."

200         Case 10054
202             GetWSAErrorString = "Connection reset by peer."

204         Case 10055
206             GetWSAErrorString = "No buffer space available."

208         Case 10056
210             GetWSAErrorString = "Socket is already connected."

212         Case 10057
214             GetWSAErrorString = "Socket is not connected."

216         Case 10058
218             GetWSAErrorString = "Can't send after socket shutdown."

220         Case 10059
222             GetWSAErrorString = "Too many references: can't splice."

224         Case 10060
226             GetWSAErrorString = "Connection timed out."

228         Case 10061
230             GetWSAErrorString = "Connection refused."

232         Case 10062
234             GetWSAErrorString = "Too many levels of symbolic links."

236         Case 10063
238             GetWSAErrorString = "File name too long."

240         Case 10064
242             GetWSAErrorString = "Host is down."

244         Case 10065
246             GetWSAErrorString = "No route to host."

248         Case 10066
250             GetWSAErrorString = "Directory not empty."

252         Case 10067
254             GetWSAErrorString = "Too many processes."

256         Case 10068
258             GetWSAErrorString = "Too many users."

260         Case 10069
262             GetWSAErrorString = "Disk quota exceeded."

264         Case 10070
266             GetWSAErrorString = "Stale NFS file handle."

268         Case 10071
270             GetWSAErrorString = "Too many levels of remote in path."

272         Case 10091
274             GetWSAErrorString = "Network subsystem is unusable."

276         Case 10092
278             GetWSAErrorString = "Winsock DLL cannot support this application."

280         Case 10093
282             GetWSAErrorString = "Winsock not initialized."

284         Case 10101
286             GetWSAErrorString = "Disconnect."

288         Case 11001
290             GetWSAErrorString = "Host not found."

292         Case 11002
294             GetWSAErrorString = "Nonauthoritative host not found."

296         Case 11003
298             GetWSAErrorString = "Nonrecoverable error."

300         Case 11004
302             GetWSAErrorString = "Valid name, no data record of requested type."

            Case Else:

        End Select

        
        Exit Function

GetWSAErrorString_Err:
304     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.GetWSAErrorString", Erl)

        
End Function

'this function DOES work on 16 and 32 bit systems
Function IpToAddr(ByVal AddrOrIP$) As String
        
        On Error GoTo 0
    
        

        

100     IpToAddr = GetHostByAddress(GetHostByNameAlias(AddrOrIP$))

102     If Err Then IpToAddr = WSA_NoName

        
        Exit Function

IpToAddr_Err:
104     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.IpToAddr", Erl)

        
End Function

'this function DOES work on 16 and 32 bit systems
Function IrcGetAscIp(ByVal IPL$) As String

        'this function is IRC specific, it expects a long ip stored in Network byte order, in a string
        'the kind that would be parsed out of a DCC command string
        On Error GoTo IrcGetAscIPError:

        Dim lpStr&
        Dim nStr&
        Dim retString$
        Dim inn&

100     If val(IPL) > 2147483647 Then
102         inn = val(IPL) - 4294967296#
    
        Else
104         inn = val(IPL)

        End If

106     inn = ntohl(inn)
108     retString = String(32, 0)
110     lpStr = inet_ntoa(inn)

112     If lpStr = 0 Then
114         IrcGetAscIp = "0.0.0.0"
            Exit Function
        End If

116     nStr = lstrlen(lpStr)
118     If nStr > 32 Then nStr = 32
    
120     Call MemCopy(ByVal retString, ByVal lpStr, nStr)
    
122     retString = Left$(retString, nStr)
    
124     IrcGetAscIp = retString
    
        Exit Function
    
IrcGetAscIPError:
126     IrcGetAscIp = "0.0.0.0"
        Exit Function
128     Resume

End Function

Public Function GetLongIp(ByVal IPS As String) As Long
        
        On Error GoTo 0
    
        

100     GetLongIp = inet_addr(IPS)

        
        Exit Function

GetLongIp_Err:
102     Call RegistrarError(Err.Number, Err.description, "WSKSOCK.GetLongIp", Erl)

        
End Function

'this function DOES work on 16 and 32 bit systems
Function IrcGetLongIp(ByVal AscIp$) As String

        'this function converts an ascii ip string into a long ip in network byte order
        'and stick it in a string suitable for use in a DCC command.
        On Error GoTo IrcGetLongIpError:

        Dim inn&

100     inn = inet_addr(AscIp)
102     inn = htonl(inn)

104     If inn < 0 Then
106         IrcGetLongIp = CVar(inn + 4294967296#)
            Exit Function
        
        Else
108         IrcGetLongIp = CVar(inn)
            Exit Function

        End If

        Exit Function
    
IrcGetLongIpError:
110     IrcGetLongIp = "0"
        Exit Function
112     Resume

End Function

'this function should work on 16 and 32 bit systems

Public Function ListenForConnect(ByVal Port&, ByVal HWndToMsg&, ByVal Enlazar As String) As Long
        On Error GoTo 0
    
        Dim S&, dummy&
        Dim SelectOps&
        Dim sockin As sockaddr

100     sockin = saZero     'zero out the structure
102     sockin.sin_family = AF_INET
104     sockin.sin_port = htons(Port)

106     If sockin.sin_port = INVALID_SOCKET Then
108         ListenForConnect = INVALID_SOCKET
            Exit Function

        End If

110     If LenB(Enlazar) = 0 Then
112         sockin.sin_addr = htonl(INADDR_ANY)
        Else
114         sockin.sin_addr = inet_addr(Enlazar)

        End If

116     If sockin.sin_addr = INADDR_NONE Then
118         ListenForConnect = INVALID_SOCKET
            Exit Function

        End If

120     S = Socket(PF_INET, SOCK_STREAM, 0)

122     If S < 0 Then
124         ListenForConnect = INVALID_SOCKET
            Exit Function

        End If
    
        'Agregado por Maraxus
        'If setsockopt(s, SOL_SOCKET, SO_CONDITIONAL_ACCEPT, True, 2) Then
        '    LogApiSock ("Error seteando conditional accept")
        '    Debug.Print "Error seteando conditional accept"
        'Else
        '    LogApiSock ("Conditional accept seteado")
        '    Debug.Print "Conditional accept seteado ^^"
        'End If
    
126     If bind(S, sockin, sockaddr_size) Then
128         If S > 0 Then
130             dummy = apiclosesocket(S)

            End If

132         ListenForConnect = INVALID_SOCKET
            Exit Function

        End If

        'SelectOps = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
134     SelectOps = FD_READ Or FD_CLOSE Or FD_ACCEPT

136     If WSAAsyncSelect(S, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
138         If S > 0 Then
140             dummy = apiclosesocket(S)

            End If

142         ListenForConnect = SOCKET_ERROR
            Exit Function

        End If
    
        'If listen(s, 5) Then
144     If listen(S, SOMAXCONN) Then
146         If S > 0 Then
148             dummy = apiclosesocket(S)

            End If

150         ListenForConnect = INVALID_SOCKET
            Exit Function

        End If

152     ListenForConnect = S

End Function

Public Function kSendData(ByVal S&, vMessage As Variant) As Long
        
        On Error GoTo 0
    
        

        Dim TheMsg() As Byte, sTemp$

100     TheMsg = vbNullString

102     Select Case VarType(vMessage)

            Case 8209   'byte array
104             sTemp = vMessage
106             TheMsg = sTemp

108         Case 8      'string, if we recieve a string, its assumed we are linemode
110             sTemp = StrConv(vMessage, vbFromUnicode)

112         Case Else
114             sTemp = CStr(vMessage)
116             sTemp = StrConv(vMessage, vbFromUnicode)

        End Select

118     TheMsg = sTemp

120     If UBound(TheMsg) > -1 Then
122         kSendData = send(S, TheMsg(0), UBound(TheMsg) + 1, 0)
        End If

        
End Function

Public Function SockAddressToString(sa As sockaddr) As String
        
        On Error GoTo 0

100     SockAddressToString = GetAscIP(sa.sin_addr) & ":" & ntohs(sa.sin_port)

        
End Function

Public Function StartWinsock(sDescription As String) As Boolean
        
        On Error GoTo 0

        Dim StartupData As WSADataType

100     If Not WSAStartedUp Then

            'If Not WSAStartup(&H101, StartupData) Then
102         If Not WSAStartup(&H202, StartupData) Then  'Use sockets v2.2 instead of 1.1 (Maraxus)
104             WSAStartedUp = True
                '            Debug.Print "wVersion="; StartupData.wVersion, "wHighVersion="; StartupData.wHighVersion
                '            Debug.Print "If wVersion == 257 then everything is kewl"
                '            Debug.Print "szDescription="; StartupData.szDescription
                '            Debug.Print "szSystemStatus="; StartupData.szSystemStatus
                '            Debug.Print "iMaxSockets="; StartupData.iMaxSockets, "iMaxUdpDg="; StartupData.iMaxUdpDg
106             sDescription = StartupData.szDescription
            Else
108             WSAStartedUp = False

            End If

        End If

110     StartWinsock = WSAStartedUp
        
End Function

Public Function WSAMakeSelectReply(TheEvent%, TheError%) As Long
        
        On Error GoTo 0

100     WSAMakeSelectReply = (TheError * &H10000) + (TheEvent And &HFFFF&)
    
End Function

