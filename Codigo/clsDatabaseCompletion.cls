VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsDatabaseCompletion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mConnection As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1
Private mCoordinator As Boolean

Public Sub Attach(ByVal connection As ADODB.Connection, ByVal coordinator As Boolean)
    Set mConnection = connection
    mCoordinator = coordinator
End Sub

Private Sub mConnection_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    ' Do not destroy the ADO recordset or start another command inside this event.
    ' The main loop consumes the completed operation on its next poll.
    If mCoordinator Then
        Call modWorldAdmission.CoordinatorCompleted(RecordsAffected, pError, pRecordset)
    Else
        Call modWorldAdmission.SaveCompleted(pCommand, pError)
    End If
End Sub
