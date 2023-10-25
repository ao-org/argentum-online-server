Attribute VB_Name = "AntiCheat"
Option Explicit


Public Enum e_ACInitResult
    eOk = 0
    eFailedPlatform
    eFAiledConnectAC
End Enum

Public Type t_AntiCheatCallbacks
    SendToClient As Long
End Type

Public Declare Function InitializeAC Lib "AOACServer.dll" (ByRef Callbacks As t_AntiCheatCallbacks) As Long
Public Declare Sub UnloadAC Lib "AOACServer.dll" ()
Public Declare Sub Update Lib "AOACServer.dll" ()

Private Function FARPROC(pfn As Long) As Long
  FARPROC = pfn
End Function

Public Sub InitializeAntiCheat()
On Error GoTo InitializeAC_Err
    Dim InitResult As e_ACInitResult
    Dim Callbacks As t_AntiCheatCallbacks
    Callbacks.SendToClient = FARPROC(AddressOf SendToClientCB)
    InitResult = InitializeAC()
    If InitResult <> eOk Then
        Call MsgBox("El juego se inicio sin activar el anti cheat, debe activarlo para poder conectarse a los servidores.")
    End If
    Exit Sub
InitializeAC_Err:
    Call TraceError(Err.Number, Err.Description, "AOAC.InitializeAC", Erl)
End Sub

Public Sub SendToClientCB(ByVal TargetUser As Long, ByRef Data As Byte, ByVal DataSize As Long)

End Sub
