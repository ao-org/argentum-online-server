VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
  Persistable = 0
  DataBindingBehavior = 0
  DataSourceBehavior = 0
  MTSTransactionMode = 0
END
Attribute VB_Name = "clsCharacterSaveWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents db As ADODB.Connection
Attribute db.VB_VarHelpID = -1
Private mJob As clsCharacterSave
Private activeCommand As ADODB.Command
Private stage As Long
Private waiting As Boolean
Private completed As Boolean
Private callbackError As String
Private quarantined As Boolean
Private submitting As Boolean
#If UNIT_TEST = 1 Then
Public TestSynchronous As Boolean
Public TestDispatchError As Boolean
Public TestCallbackDuringSubmit As Boolean
Friend Property Get TestCommand() As ADODB.Command
    Set TestCommand = activeCommand
End Property
Friend Sub TestLateCompletion(ByVal command As ADODB.Command)
    Dim status As ADODB.EventStatusEnum
    status = adStatusOK
    Call db_ExecuteComplete(1, Nothing, status, command, Nothing, db)
End Sub
#End If

Friend Sub OpenConnection(ByVal connectionString As String)
    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    Call db.Open(connectionString)
    Call db.Execute("PRAGMA foreign_keys=ON")
End Sub
Friend Property Get Idle() As Boolean
    Idle = (mJob Is Nothing) And Not quarantined
End Property
Friend Sub Start(ByVal job As clsCharacterSave)
    If Not Idle Then Call Err.Raise(5, , "Character save connection already owned")
    Set mJob = job
    stage = job.NextStatement
    waiting = False
    completed = False
End Sub

' Per-statement autocommit avoids holding a SQLite write lock across main-loop callbacks.
' The event only records completion. Main-loop Tick advances one statement at a time.
Friend Sub Tick()
    On Error GoTo DispatchFailed
    If mJob Is Nothing Or quarantined Or submitting Then Exit Sub
    If waiting Then
        If Not completed Then Exit Sub
        waiting = False
        completed = False
        Set activeCommand.ActiveConnection = Nothing
        Set activeCommand = Nothing
        If LenB(callbackError) Then
            Call mJob.MarkFailed("statement=" & stage & " " & callbackError, True)
            Set mJob = Nothing
            Exit Sub
        Else
            Call mJob.StatementCompleted(stage)
            stage = stage + 1
        End If
    End If
    If stage > mJob.Count Then
        Call mJob.Committed()
        Set mJob = Nothing
        Exit Sub
    End If
    Set activeCommand = mJob.Statement(stage)
    Set activeCommand.ActiveConnection = db
    callbackError = vbNullString
    completed = False
    waiting = True
    submitting = True
    #If UNIT_TEST = 1 Then
        If TestDispatchError Then Call Err.Raise(5, , "Injected dispatch exception")
        If TestSynchronous Then
            Call activeCommand.Execute()
        Else
            Call activeCommand.Execute(, , adAsyncExecute)
        End If
    #Else
        Call activeCommand.Execute(, , adAsyncExecute)
    #End If
    submitting = False
    Exit Sub
DispatchFailed:
    submitting = False
    ' An Execute exception does not prove no work reached the provider. Never reuse it.
    Call mJob.MarkFailed("dispatch stage=" & stage & " error=" & Err.Number & " " & Err.Description, False)
    quarantined = True
End Sub

Private Sub db_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    If Not waiting Or completed Or quarantined Then Exit Sub
    If Not (pConnection Is db) Then Exit Sub
    If Not (pCommand Is activeCommand) Then Exit Sub
    #If UNIT_TEST = 1 Then
        If submitting Then TestCallbackDuringSubmit = True
    #End If
    If Not pError Is Nothing Then callbackError = CStr(pError.Number) & " " & pError.Description
    If adStatus = adStatusErrorsOccurred And LenB(callbackError) = 0 Then callbackError = "ADO execution failed without error details"
    If LenB(callbackError) = 0 And Not mJob.ValidAffectedRows(stage, RecordsAffected) Then callbackError = "Character UPDATE did not affect exactly one row"
    completed = True
End Sub
