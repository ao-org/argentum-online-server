Attribute VB_Name = "AntiCheat"
Option Explicit

Public Enum e_ACInitResult
    eOk = 0
    eFailedPlatform
    eFAiledConnectAC
End Enum

Private Type SINGLESTRINGPARAM
    Ptr As Long
    Len As Long
End Type

Public Type t_AntiCheatCallbacks
    SendToClient As Long
    LogMessage As Long
    RegisterRemoteUserId As Long
    ActionRequired As Long
End Type

Public Enum e_ActionRequiredType
    eEOS_ACCCA_Invalid = 0
    eEOS_ACCCA_RemovePlayer = 1
End Enum

Public Enum e_ActionRequiredReason
    eEOS_ACCCAR_Invalid = 0
    eEOS_ACCCAR_InternalError = 1
    eEOS_ACCCAR_InvalidMessage = 2
    eEOS_ACCCAR_AuthenticationFailed = 3
    eEOS_ACCCAR_NullClient = 4
    eEOS_ACCCAR_HeartbeatTimeout = 5
    eEOS_ACCCAR_ClientViolation = 6
    eEOS_ACCCAR_BackendViolation = 7
    eEOS_ACCCAR_TemporaryCooldown = 8
    eEOS_ACCCAR_TemporaryBanned = 9
    eEOS_ACCCAR_PermanentBanned = 10
End Enum

Public Enum EOS_ELogLevel
    EOS_LOG_Off = 0
    EOS_LOG_Fatal = 100
    EOS_LOG_Error = 200
    EOS_LOG_Warning = 300
    EOS_LOG_Info = 400
    EOS_LOG_Verbose = 500
    EOS_LOG_VeryVerbose = 600
End Enum

Private Declare Function InitializeAC Lib "AOACServer.dll" (ByRef Callbacks As t_AntiCheatCallbacks) As Long
Private Declare Sub UnloadAC Lib "AOACServer.dll" ()
Private Declare Sub Update Lib "AOACServer.dll" ()
Private Declare Sub AddPendingRegister Lib "AOACServer.dll" (ByRef UserReference As t_UserReference)
Private Declare Function QueryAndRemoveOldPendingRegistey Lib "AOACServer.dll" (ByRef UserReference As t_UserReference, ByVal ElapsedThreshold As Long) As Long
Private Declare Sub UnRegisterClient Lib "AOACServer.dll" (ByVal UserIndex As Integer)
Private Declare Sub HandleRemoteMessage Lib "AOACServer.dll" (ByRef UserReference As t_UserReference, ByRef data As Byte, ByVal DataSize As Integer)
Dim EnableAnticheat As Boolean

Private Function GetStringFromPtr(ByVal Ptr As Long, ByVal Size As Long) As String
    Dim Buffer() As Byte
    ReDim Buffer(0 To (Size - 1)) As Byte
    CopyMemory Buffer(0), ByVal Ptr, Size
    GetStringFromPtr = StrConv(Buffer, vbUnicode)
End Function

Private Function FARPROC(pfn As Long) As Long
    FARPROC = pfn
End Function

Public Sub InitializeAntiCheat()
    On Error GoTo InitializeAC_Err
    EnableAnticheat = IsFeatureEnabled("anti-cheat")
    If EnableAnticheat Then
        Dim InitResult As e_ACInitResult
        Dim Callbacks  As t_AntiCheatCallbacks
        Callbacks.SendToClient = FARPROC(AddressOf SendToClientCB)
        Callbacks.LogMessage = FARPROC(AddressOf LogMessageCB)
        Callbacks.RegisterRemoteUserId = FARPROC(AddressOf RegisterRemoteUserIdCb)
        Callbacks.ActionRequired = FARPROC(AddressOf ClientActionRequired)
        InitResult = InitializeAC(Callbacks)
        If InitResult <> eOk Then
            Call MsgBox("El juego se inicio sin activar el anti cheat, debe activarlo para poder conectarse a los servidores.")
        End If
    End If
    Exit Sub
InitializeAC_Err:
    Call TraceError(Err.Number, Err.Description, "AOAC.InitializeAC", Erl)
End Sub

Public Sub OnNewPlayerConnect(ByVal UserIndex As Integer)
    If EnableAnticheat Then
        Dim UserRef As t_UserReference
        Call SetUserRef(UserRef, UserIndex)
        Call AddPendingRegister(UserRef)
        Call WriteAntiCheatStartSeassion(UserIndex)
    End If
End Sub

Public Sub KickUnregisteredPlayers()
    If EnableAnticheat Then
        Dim UserRef As t_UserReference
        Dim Result  As Long
        Result = QueryAndRemoveOldPendingRegistey(UserRef, 30000)
        If Result > 0 And IsValidUserRef(UserRef) Then
            'Call modNetwork.Kick(UserList(UserRef.ArrayIndex).ConnectionDetails.ConnID, "Anticheat detection timeout")
        End If
    End If
End Sub

Public Sub AntiCheatUpdate()
    If EnableAnticheat Then
        Call Update
        #If DIRECT_PLAY = 0 Then
            'Needs debugging to understand why this does not work with DPLAY, suspect the problem is in Read/Write/SafeArray
            Call KickUnregisteredPlayers
        #End If
    End If
End Sub

Public Sub UnloadAntiCheat()
    If EnableAnticheat Then
        Call UnloadAC
    End If
End Sub

Public Sub SendToClientCB(ByRef TargetUser As t_UserReference, ByVal data As Long, ByVal DataSize As Long)
    If EnableAnticheat Then
        If IsValidUserRef(TargetUser) Then
            Call WriteAntiCheatMessage(TargetUser.ArrayIndex, data, DataSize)
        End If
    End If
End Sub

Public Sub HandleAntiCheatServerMessage(ByVal UserIndex As Integer, ByRef data() As Byte)
    If EnableAnticheat Then
        Dim UserRef As t_UserReference
        Call SetUserRef(UserRef, UserIndex)
        Call HandleRemoteMessage(UserRef, data(0), UBound(data))
    End If
End Sub

Public Sub LogMessageCB(ByRef Message As SINGLESTRINGPARAM, ByVal LogLevel As Long)
    Dim MessageStr As String
    If Message.Len > 0 Then
        MessageStr = GetStringFromPtr(Message.Ptr, Message.Len)
    End If
    If LogLevel Then
    End If
    If LogLevel < EOS_LOG_Warning Then
        Call LogThis(0, "Anticheat: " & MessageStr, vbLogEventTypeError)
    End If
End Sub

Public Sub RegisterRemoteUserIdCb(ByRef UserRef As t_UserReference, ByRef Id As SINGLESTRINGPARAM)
    Dim IdStr As String
    If Id.Len > 0 Then
        IdStr = GetStringFromPtr(Id.Ptr, Id.Len)
    End If
    If IsValidUserRef(UserRef) Then
        Call SaveEpicLogin(IdStr, UserRef.ArrayIndex)
    End If
End Sub

Public Sub ClientActionRequired(ByRef UserRef As t_UserReference, ByVal Action As Long, ByRef ReasonString As SINGLESTRINGPARAM)
    Dim ReasonStr As String
    If ReasonString.Len > 0 Then
        ReasonStr = GetStringFromPtr(ReasonString.Ptr, ReasonString.Len)
    End If
    If Action = eEOS_ACCCA_RemovePlayer And IsValidUserRef(UserRef) Then
        'Call modNetwork.Kick(UserList(UserRef.ArrayIndex).ConnectionDetails.ConnID, ReasonStr)
    End If
End Sub

Public Sub OnPlayerDisconnect(ByVal UserIndex As Integer)
    If EnableAnticheat Then
        Call UnRegisterClient(UserIndex)
    End If
End Sub
