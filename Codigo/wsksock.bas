Attribute VB_Name = "modWinsock"
Option Explicit

Public Declare Function API_closesocket Lib "ws2_32.DLL" Alias "closesocket" (ByVal SocketID As Long) As Long
Public Declare Function API_send Lib "ws2_32.DLL" Alias "send" (ByVal SocketID As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function API_inet_ntoa Lib "ws2_32.DLL" Alias "inet_ntoa" (ByVal inn As Long) As Long
Public Declare Function API_lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Declare Sub API_MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)

'---WINDOWS EXTENSIONS
Private Declare Function API_WSAStartup Lib "ws2_32.DLL" Alias "WSAStartup" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Private Declare Function API_WSACleanup Lib "ws2_32.DLL" Alias "WSACleanup" () As Long
Private Declare Function API_WSACancelBlockingCall Lib "ws2_32.DLL" Alias "WSACancelBlockingCall" () As Long
Private Declare Function API_WSAIsBlocking Lib "ws2_32.DLL" Alias "WSAIsBlocking" () As Long

'Agregado por Maraxus
Public Type WSABUF
    dwBufferLen As Long
    LpBuffer    As Long
End Type

'Agregado por Maraxus
Public Type FLOWSPEC
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
Public Const CF_ACCEPT = &H0
Public Const CF_REJECT = &H1

Public Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Public Const sockaddr_size = 16


' --------------------------------------------------------
' Used for WinsockAPI startup process
' --------------------------------------------------------
Private Const WSA_DESCRIPTIONLEN = 256
Private Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1

Private Const WSA_SYS_STATUS_LEN = 128
Private Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Private Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
' --------------------------------------------------------
' Used for WinsockAPI startup process
' --------------------------------------------------------

' Error codes used outside the class
Global Const WSAEWOULDBLOCK = 10035
Global Const WSAEMSGSIZE = 10040

' Flag to keep track of whether winsock WSAStartup was called
Public WSAStartedUp As Boolean

Public Sub InitializeWinsockAPI()
    
    ' Prevent multiple instances of the API
    If WSAStartedUp Then Exit Sub
    
    Dim StartupData As WSADataType

    If Not WSAStartedUp Then
        
        'Using sockets v2.2
        If Not API_WSAStartup(&H202, StartupData) Then
            WSAStartedUp = True
            
        Else
            WSAStartedUp = False

        End If

    End If
    
End Sub

Public Sub DestroyWinsockAPI()
    
    If Not WSAStartedUp Then Exit Sub
    
    Dim Ret As Long

    If API_WSAIsBlocking() Then
        Ret = API_WSACancelBlockingCall()
    End If
    
    Ret = API_WSACleanup()
    
    WSAStartedUp = False
    
End Sub

Public Function GetAscIP(ByVal inn As Long) As String
        
    On Error GoTo 0

    Dim nStr As Long
    Dim retString As String: retString = String(32, 0)
    Dim lpStr As Long: lpStr = API_inet_ntoa(inn)

    If lpStr Then
    
        nStr = API_lstrlen(lpStr)
        If nStr > 32 Then nStr = 32
        
        Call API_MemCopy(ByVal retString, ByVal lpStr, nStr)
        
        retString = Left$(retString, nStr)
        GetAscIP = retString
        
    Else
    
        GetAscIP = "255.255.255.255"

    End If
    
End Function

