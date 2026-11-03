VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
  Persistable = 0
  DataBindingBehavior = 0
  DataSourceBehavior = 0
  MTSTransactionMode = 0
END
Attribute VB_Name = "clsPartyQuestAsyncFence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Server. Copyright (C) 2026 Noland Studios LTD.
' Licensed under the GNU Affero General Public License, version 3 or later.
Option Explicit
Private WithEvents mConnection As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1
Private mPending As Long
Private mStarted As Long
Private mCompleted As Long
Public Sub Attach(ByVal source As ADODB.Connection)
    Set mConnection = source
End Sub
Public Property Get Started() As Long
    Started = mStarted
End Property
Public Property Get Completed() As Long
    Completed = mCompleted
End Property
Public Property Get Pending() As Long
    Pending = mPending
End Property
Private Sub mConnection_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    mPending = mPending + 1: mStarted = mStarted + 1
End Sub
Private Sub mConnection_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    If mPending > 0 Then mPending = mPending - 1
    mCompleted = mCompleted + 1
    ' Observe completion only. Never destroy pRecordset in an ADO callback.
End Sub
