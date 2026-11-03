VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
  Persistable = 0
  DataBindingBehavior = 0
  DataSourceBehavior = 0
  MTSTransactionMode = 0
END
Attribute VB_Name = "clsHuntAsyncFence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit
Private WithEvents mConnection As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1
Private mCommands As Collection
Private mStarted As Long
Private mCompleted As Long
Public Sub Attach(ByVal source As ADODB.Connection)
    Set mCommands = New Collection
    Set mConnection = source
End Sub
Public Property Get Started() As Long
    Started = mStarted
End Property
Public Property Get Completed() As Long
    Completed = mCompleted
End Property
Public Property Get Pending() As Long
    Dim i As Long, command As ADODB.Command
    For i = mCommands.count To 1 Step -1
        Set command = mCommands(i)
        If (command.State And (adStateExecuting Or adStateFetching Or adStateConnecting)) = 0 Then Call mCommands.Remove(i)
    Next i
    Pending = mCommands.count
End Property
Private Sub mConnection_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Dim command As ADODB.Command
    mStarted = mStarted + 1
    For Each command In mCommands
        If command Is pCommand Then Exit Sub
    Next command
    Call mCommands.Add(pCommand)
End Sub
Private Sub mConnection_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    mCompleted = mCompleted + 1
    ' Delivery can lag provider completion. Pending uses command.State, so no
    ' DoEvents/reentrant game processing is needed. Never destroy pRecordset.
End Sub
