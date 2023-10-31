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

Private Declare Function InitializeAC Lib "AOACServer.dll" (ByRef Callbacks As t_AntiCheatCallbacks) As Long
Private Declare Sub UnloadAC Lib "AOACServer.dll" ()
Private Declare Sub Update Lib "AOACServer.dll" ()
Private Declare Sub AddPendingRegister Lib "AOACServer.dll" (ByRef UserReference As t_UserReference)
Private Declare Function QueryAndRemoveOldPendingRegistey Lib "AOACServer.dll" (ByRef UserReference As t_UserReference, ByVal ElapsedThreshold As Long) As Long
Private Declare Sub UnRegisterClient Lib "AOACServer.dll" (ByVal UserIndex As Integer)
Private Declare Sub HandleRemoteMessage Lib "AOACServer.dll" (ByRef UserReference As t_UserReference, ByRef Data As Byte, ByVal DataSize As Integer)

Private Function GetStringFromPtr(ByVal Ptr As Long, ByVal size As Long) As String
    Dim Buffer() As Byte
    ReDim Buffer(0 To (size - 1)) As Byte
    CopyMemory Buffer(0), ByVal Ptr, size
    GetStringFromPtr = StrConv(Buffer, vbUnicode)
End Function

Private Function FARPROC(pfn As Long) As Long
  FARPROC = pfn
End Function

Public Sub InitializeAntiCheat()
On Error GoTo InitializeAC_Err
    Dim InitResult As e_ACInitResult
    Dim Callbacks As t_AntiCheatCallbacks
    Callbacks.SendToClient = FARPROC(AddressOf SendToClientCB)
    Callbacks.LogMessage = FARPROC(AddressOf LogMessageCB)
    Callbacks.RegisterRemoteUserId = FARPROC(AddressOf RegisterRemoteUserIdCb)
    Callbacks.ActionRequired = FARPROC(AddressOf ClientActionRequired)
    InitResult = InitializeAC(Callbacks)
    If InitResult <> eOk Then
        Call MsgBox("El juego se inicio sin activar el anti cheat, debe activarlo para poder conectarse a los servidores.")
    End If
    Exit Sub
InitializeAC_Err:
    Call TraceError(Err.Number, Err.Description, "AOAC.InitializeAC", Erl)
End Sub

Public Sub OnNewPlayerConnect(ByVal UserIndex As Integer)
    Dim UserRef As t_UserReference
    Call SetUserRef(UserRef, UserIndex)
    Call AddPendingRegister(UserRef)
    Call WriteAntiCheatStartSeassion(UserIndex)
End Sub

Public Sub KickUnregisteredPlayers()
    Dim UserRef As t_UserReference
    Dim Result As Long
    Result = QueryAndRemoveOldPendingRegistey(UserRef, 10000)
    If Result > 0 And IsValidUserRef(UserRef) Then
        Call modNetwork.Kick(UserList(UserRef.ArrayIndex).ConnectionDetails.ConnID, "Anticheat detection timeout")
    End If
End Sub

Public Sub AntiCheatUpdate()
    Call Update
    Call KickUnregisteredPlayers
End Sub

Public Sub UnloadAntiCheat()
    Call UnloadAC
End Sub
Public Sub SendToClientCB(ByRef TargetUser As t_UserReference, ByVal Data As Long, ByVal DataSize As Long)
    If IsValidUserRef(TargetUser) Then
        Call WriteAntiCheatMessage(TargetUser.ArrayIndex, Data, DataSize)
    End If
End Sub

Public Sub HandleAntiCheatServerMessage(ByVal UserIndex As Integer, ByRef Data() As Byte)
    Dim UserRef As t_UserReference
    Call SetUserRef(UserRef, UserIndex)
    Call HandleRemoteMessage(UserRef, Data(0), UBound(Data))
End Sub

Public Sub LogMessageCB(ByRef Message As SINGLESTRINGPARAM)
    Dim MessageStr As String
    If Message.Len > 0 Then
        MessageStr = GetStringFromPtr(Message.Ptr, Message.Len)
    End If
    Debug.Print MessageStr
End Sub

Public Sub RegisterRemoteUserIdCb(ByRef UserRef As t_UserReference, ByRef Id As SINGLESTRINGPARAM)
    Dim IdStr As String
    If Id.Len > 0 Then
        IdStr = GetStringFromPtr(Id.Ptr, Id.Len)
    End If
End Sub

Public Sub ClientActionRequired(ByRef UserRef As t_UserReference, ByVal Action As Long, ByVal ReasonCode As Long, ByRef ReasonString As SINGLESTRINGPARAM)
    Dim ReasonStr As String
    If ReasonString.Len > 0 Then
        ReasonStr = GetStringFromPtr(ReasonString.Ptr, ReasonString.Len)
    End If
    If Action = eEOS_ACCCA_RemovePlayer And IsValidUserRef(UserRef) Then
        Call modNetwork.Kick(UserList(UserRef.ArrayIndex).ConnectionDetails.ConnID, ReasonStr)
    End If
End Sub

Public Sub OnPlayerDisconnect(ByVal UserIndex As Integer)
    Call UnRegisterClient(UserIndex)
End Sub
