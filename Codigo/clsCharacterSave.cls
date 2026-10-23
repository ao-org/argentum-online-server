VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
  Persistable = 0
  DataBindingBehavior = 0
  DataSourceBehavior = 0
  MTSTransactionMode = 0
END
Attribute VB_Name = "clsCharacterSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCommands As New Collection
Private mSnapshot As t_User
Private mCharacterId As Long
Private mSession As Long
Private mUserIndex As Integer
Private mFinal As Boolean
Private mSealed As Boolean
Private mState As e_CharacterSaveResult
Private mFailure As String
Private mRetryable As Boolean
Private mCompletedStatements As Long
Private mCreation As Boolean
Private mSingleRowStatement As Long

Friend Sub Initialize(ByVal userIndex As Integer, ByRef User As t_User, ByVal finalSave As Boolean)
    mCharacterId = User.Id
    mSession = User.SaveSession
    mUserIndex = userIndex
    mFinal = finalSave
End Sub

' Parameters are copied into private commands now, never read from a reused UserList slot.
Friend Sub Add(ByVal sql As String, ParamArray values() As Variant)
    If mSealed Then Call Err.Raise(5, , "Save snapshot is sealed")
    Dim cmd As New ADODB.Command, value As Variant, inner As Variant
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    For Each value In values
        If IsArray(value) Then
            For Each inner In value
                Call cmd.Parameters.Append(CreateParameter(inner, adParamInput))
            Next inner
        Else
            Call cmd.Parameters.Append(CreateParameter(value, adParamInput))
        End If
    Next value
    Call mCommands.Add(cmd)
End Sub

Friend Sub AddCards(ByRef User As t_User)
    If mSealed Then Call Err.Raise(5, , "Save snapshot is sealed")
    Dim cmd As New ADODB.Command, data(0 To 1023) As Byte, i As Long
    For i = 1 To UBound(User.AccountCollectibleCardQuantities)
        data(i - 1) = User.AccountCollectibleCardQuantities(i)
    Next i
    cmd.CommandType = adCmdText
    cmd.CommandText = UPSERT_NEW_COLLECTIBLE_CARDS
    Call cmd.Parameters.Append(cmd.CreateParameter(, adInteger, adParamInput, , User.AccountID))
    Call cmd.Parameters.Append(cmd.CreateParameter(, adVarBinary, adParamInput, 1024, data))
    Call mCommands.Add(cmd)
End Sub

Friend Sub Seal(ByRef User As t_User)
    mSnapshot = User
    mSealed = True
End Sub

Friend Sub MarkCreation()
    mCreation = True
End Sub

Friend Sub RequireSingleRow()
    mSingleRowStatement = mCommands.Count
End Sub

Friend Function ValidAffectedRows(ByVal index As Long, ByVal affected As Long) As Boolean
    ValidAffectedRows = (index <> mSingleRowStatement Or affected = 1)
End Function

Friend Sub FinishCreation()
    If Not mCreation Or mState <> CharacterSaveSucceeded Then Exit Sub
    If mUserIndex < 1 Or mUserIndex > MaxUsers Then Exit Sub
    With UserList(mUserIndex)
        If .Id <> mCharacterId Or .SaveSession <> mSession Then Exit Sub
        If Not .ConnectionDetails.ConnIDValida Or .flags.UserLogged Then Exit Sub
        #If PYMMO = 1 Then
            Call ConnectUser(mUserIndex, .name, True)
        #Else
            Call ConnectUser(mUserIndex, .name, False)
        #End If
    End With
End Sub

Friend Sub CaptureDeferredReset()
    If mFinal Or mUserIndex < 1 Or mUserIndex > MaxUsers Then Exit Sub
    With UserList(mUserIndex)
        If .Id <> mCharacterId Or .SaveSession <> mSession Or Not .flags.UserLogged Then Exit Sub
        If .SaveResetEpoch <> .SavedResetEpoch Then Call SaveCharacterDB(mUserIndex)
    End With
End Sub

Friend Property Get Statement(ByVal index As Long) As ADODB.Command
    ' A resumed attempt gets a fresh identity; a late event from the failed attempt cannot match it.
    Dim source As ADODB.Command, cmd As New ADODB.Command, parameter As ADODB.Parameter
    Set source = mCommands(index)
    cmd.CommandType = source.CommandType
    cmd.CommandText = source.CommandText
    For Each parameter In source.Parameters
        Call cmd.Parameters.Append(cmd.CreateParameter(parameter.name, parameter.Type, parameter.Direction, parameter.Size, parameter.value))
    Next parameter
    Set Statement = cmd
End Property
Friend Property Get Count() As Long
    Count = mCommands.Count
End Property
Friend Property Get CharacterId() As Long
    CharacterId = mCharacterId
End Property
Friend Property Get SessionId() As Long
    SessionId = mSession
End Property
Friend Property Get IsFinal() As Boolean
    IsFinal = mFinal
End Property
Friend Property Get Result() As e_CharacterSaveResult
    Result = mState
End Property
Friend Property Get Failure() As String
    Failure = mFailure
End Property
Friend Property Get Retryable() As Boolean
    Retryable = mRetryable
End Property
Friend Property Get NextStatement() As Long
    NextStatement = mCompletedStatements + 1
End Property
Friend Sub StatementCompleted(ByVal index As Long)
    If index <> mCompletedStatements + 1 Then Call Err.Raise(5, , "Unexpected save completion order")
    mCompletedStatements = index
End Sub
Friend Sub MarkFailed(ByVal reason As String, ByVal completedFailure As Boolean)
    mState = CharacterSaveFailed
    mFailure = reason
    mRetryable = completedFailure
    Call LogDatabaseError("Character save FAILED char=" & mCharacterId & " session=" & mSession & " final=" & mFinal & " completed=" & mCompletedStatements & " callback_failure=" & completedFailure & " " & reason)
    If mCompletedStatements < mCommands.Count Then Call LogDatabaseError("Failed save SQL: " & Left$(mCommands(mCompletedStatements + 1).CommandText, 512))
End Sub
Friend Sub Retry()
    If Not mRetryable Then Call Err.Raise(5, , "Save outcome is uncertain; retry refused")
    mState = CharacterSavePending
    mRetryable = False
    mFailure = vbNullString
End Sub
Friend Sub MarkRunning()
    mState = CharacterSaveRunning
End Sub
Friend Sub Committed()
    If mState = CharacterSaveSucceeded Then Exit Sub
    On Error GoTo AcknowledgeFailed
    If mCompletedStatements <> mCommands.Count Then
        Call MarkFailed("Incomplete save cannot succeed", False)
        Exit Sub
    End If
    mState = CharacterSaveSucceeded
    If mUserIndex < 1 Or mUserIndex > MaxUsers Then Exit Sub
    With UserList(mUserIndex)
        If .Id <> mCharacterId Or .SaveSession <> mSession Then Exit Sub
    End With
    Call AcknowledgeCharacterSave(UserList(mUserIndex), mSnapshot)
    Exit Sub
AcknowledgeFailed:
    Call LogDatabaseError("Save committed; dirty acknowledgement failed char=" & mCharacterId & " " & Err.Description)
End Sub
