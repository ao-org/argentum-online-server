Attribute VB_Name = "modCharacterSaves"
Option Explicit

Public Enum e_CharacterSaveResult
    CharacterSavePending = 0
    CharacterSaveRunning = 1
    CharacterSaveSucceeded = 2
    CharacterSaveFailed = 3
End Enum

Private Const SAVE_WORKERS As Long = 2
Private sessions As Scripting.Dictionary
Private sessionCharacters As Scripting.Dictionary
Private queues As Scripting.Dictionary
Private retired As Scripting.Dictionary
Private finalized As Scripting.Dictionary
Private workers(1 To SAVE_WORKERS) As clsCharacterSaveWorker
Private sequence As Long
Private pumping As Boolean
Private stopping As Boolean
Private ready As Boolean
Private notifySaved As Boolean

Public Sub InitializeCharacterSaves(ByVal connectionString As String)
    Dim i As Long
    ready = False
    Set sessions = New Scripting.Dictionary
    Set sessionCharacters = New Scripting.Dictionary
    Set queues = New Scripting.Dictionary
    Set retired = New Scripting.Dictionary
    Set finalized = New Scripting.Dictionary
    stopping = False
    notifySaved = False
    For i = 1 To SAVE_WORKERS
        Set workers(i) = New clsCharacterSaveWorker
        Call workers(i).OpenConnection(connectionString)
    Next i
    ready = True
End Sub

Public Function CharacterSaveCapacityAvailable() As Boolean
    If Not ready Or sessions Is Nothing Or stopping Then Exit Function
    CharacterSaveCapacityAvailable = (sessions.Count < MaxUsers And sequence < &H7FFFFFFF)
End Function

' Reserve space at login, including one final snapshot even when an autosave is pending.
Public Function ReserveCharacterSaveSession(ByRef User As t_User, ByRef reason As String) As Boolean
    Dim key As String
    key = CStr(User.Id)
    reason = "El personaje tiene un guardado pendiente o fallido. Intenta de nuevo; si persiste, contacta al administrador."
    If Not ready Or sessions Is Nothing Or User.Id <= 0 Then Exit Function
    If User.SaveSession <> 0 Then
        If Not sessionCharacters.Exists(CStr(User.SaveSession)) Then Exit Function
        If sessionCharacters(CStr(User.SaveSession)) <> key Then Exit Function
    End If
    If sessions.Exists(key) Then
        ReserveCharacterSaveSession = (User.SaveSession <> 0 And User.SaveSession = sessions(key))
        Exit Function
    End If
    If stopping Then Exit Function
    If sessions.Count >= MaxUsers Or sequence = &H7FFFFFFF Then
        reason = "No hay capacidad de guardado disponible. Contacta al administrador."
        Exit Function
    End If
    sequence = sequence + 1
    User.SaveSession = sequence
    Call sessions.Add(key, sequence)
    Call sessionCharacters.Add(CStr(sequence), key)
    Call queues.Add(key, New Collection)
    ReserveCharacterSaveSession = True
End Function

Public Function CharacterHasUnfinishedSave(ByVal characterId As Long) As Boolean
    If queues Is Nothing Then Exit Function
    If Not queues.Exists(CStr(characterId)) Then Exit Function
    Dim jobs As Collection, job As clsCharacterSave
    Set jobs = queues(CStr(characterId))
    For Each job In jobs
        If job.Result <> CharacterSaveSucceeded Then CharacterHasUnfinishedSave = True: Exit Function
    Next job
End Function

Public Sub ReleaseCharacterSaveSession(ByRef User As t_User)
    If sessions Is Nothing Then Exit Sub
    Dim key As String
    ' A rejected replacement login may already have overwritten User.Id.
    If Not sessionCharacters.Exists(CStr(User.SaveSession)) Then Exit Sub
    key = sessionCharacters(CStr(User.SaveSession))
    If Not sessions.Exists(key) Then Exit Sub
    If sessions(key) <> User.SaveSession Then Exit Sub
    retired(key) = True
    Call RemoveRetiredSaveSession(key)
End Sub

Private Sub RemoveRetiredSaveSession(ByVal key As String)
    If Not queues.Exists(key) Then Exit Sub
    Dim jobs As Collection
    Set jobs = queues(key)
    If jobs.Count <> 0 Or Not retired.Exists(key) Then Exit Sub
    Call queues.Remove(key)
    Call sessionCharacters.Remove(CStr(sessions(key)))
    Call sessions.Remove(key)
    Call retired.Remove(key)
    If finalized.Exists(key) Then Call finalized.Remove(key)
End Sub

Public Function CanCaptureCharacterSave(ByRef User As t_User, ByVal finalSave As Boolean) As Boolean
    Dim reason As String, jobs As Collection, job As clsCharacterSave
    If Not ReserveCharacterSaveSession(User, reason) Then
        Call LogDatabaseError("Character save reservation failed char=" & User.Id & " " & reason)
        Exit Function
    End If
    Set jobs = queues(CStr(User.Id))
    If finalized.Exists(CStr(User.Id)) Then Exit Function
    For Each job In jobs
        If job.IsFinal Then Exit Function
    Next job
    ' Autosaves defer without touching dirty flags; logout has its reserved second slot.
    CanCaptureCharacterSave = finalSave Or jobs.Count = 0
End Function

Public Sub EnqueueCharacterSave(ByVal job As clsCharacterSave)
    Dim key As String, jobs As Collection
    key = CStr(job.CharacterId)
    If Not sessions.Exists(key) Then Call Err.Raise(5, , "Missing save reservation")
    If sessions(key) <> job.SessionId Then Call Err.Raise(5, , "Stale save session")
    If finalized.Exists(key) Then Call Err.Raise(5, , "Session already has a final save")
    Set jobs = queues(key)
    If jobs.Count >= 2 Then Call Err.Raise(5, , "Reserved final save slot exhausted")
    Call jobs.Add(job)
    If job.IsFinal Then finalized(key) = True
End Sub

Public Sub TickCharacterSaves()
    On Error GoTo Failed
    If queues Is Nothing Or pumping Then Exit Sub
    pumping = True
    Dim i As Long, key As Variant, jobs As Collection, job As clsCharacterSave
    For i = 1 To SAVE_WORKERS
        Call workers(i).Tick()
    Next i
    For Each key In queues.Keys
        Set jobs = queues(key)
        If jobs.Count Then
            Set job = jobs(1)
            If job.Result = CharacterSaveSucceeded Then
                Call jobs.Remove(1)
                Call job.FinishCreation()
                Call job.CaptureDeferredReset()
            End If
        End If
        If jobs.Count Then
            Set job = jobs(1)
            If job.Result = CharacterSavePending Then
                For i = 1 To SAVE_WORKERS
                    If workers(i).Idle Then
                        Call job.MarkRunning()
                        Call workers(i).Start(job)
                        Call workers(i).Tick()
                        Exit For
                    End If
                Next i
            End If
        End If
        Call RemoveRetiredSaveSession(CStr(key))
    Next key
    If notifySaved Then
        If Not CharacterSavesUnresolved() Then
            notifySaved = False
            If Not stopping Then Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(MSG_SERVIDOR_PERSONAJES_GRABADOS, vbNullString, e_TextChannel.TEXTCHANNEL_SERVER_STAFF, e_FontTypeNames.FONTTYPE_SERVER))
        End If
    End If
    pumping = False
    Exit Sub
Failed:
    pumping = False
    Call LogDatabaseError("Character save scheduler error=" & Err.Number & " " & Err.Description)
End Sub

Public Sub RequestCharacterSaveNotice()
    notifySaved = True
End Sub

Public Function CharacterSavesUnresolved() As Boolean
    If queues Is Nothing Then Exit Function
    Dim key As Variant, jobs As Collection
    For Each key In queues.Keys
        Set jobs = queues(key)
        If jobs.Count Then CharacterSavesUnresolved = True: Exit Function
    Next key
End Function

Public Sub ReportUnresolvedCharacterSaves()
    If queues Is Nothing Then Exit Sub
    Dim key As Variant, jobs As Collection, job As clsCharacterSave
    For Each key In queues.Keys
        Set jobs = queues(key)
        For Each job In jobs
            Call LogDatabaseError("Unresolved character save char=" & job.CharacterId & " session=" & job.SessionId & " final=" & job.IsFinal & " state=" & job.Result & " next_statement=" & job.NextStatement & " statements=" & job.Count & " " & job.Failure)
        Next job
    Next key
End Sub

' Operator-reviewed recovery only: resume at the failed statement, never replay completed writes.
' The operator must first resolve the logged SQL failure. Uncertain dispatch is not resumable.
Public Function RetryFailedCharacterSave(ByVal characterId As Long) As Boolean
    If queues Is Nothing Then Exit Function
    If Not queues.Exists(CStr(characterId)) Then Exit Function
    Dim jobs As Collection, job As clsCharacterSave
    Set jobs = queues(CStr(characterId))
    If jobs.Count = 0 Then Exit Function
    Set job = jobs(1)
    If job.Result <> CharacterSaveFailed Or Not job.Retryable Then Exit Function
    Call job.Retry()
    Call LogDatabaseError("Operator resumed character save char=" & characterId & " session=" & job.SessionId & " statement=" & job.NextStatement)
    RetryFailedCharacterSave = True
End Function

' Only invoked by the local operator's Guardar Usuarios button, never by a timer.
Public Sub OfferCharacterSaveRecovery()
    On Error GoTo Failed
    If queues Is Nothing Then Exit Sub
    Dim key As Variant, jobs As Collection, job As clsCharacterSave, found As Boolean
    For Each key In queues.Keys
        Set jobs = queues(key)
        If jobs.Count Then
            Set job = jobs(1)
            If job.Result = CharacterSaveFailed Then found = True
        End If
    Next key
    Call ReportUnresolvedCharacterSaves()
    If Not found Then Exit Sub
    Dim answer As String, id As Long
    answer = InputBox("Hay guardados fallidos. Revisa el log de base de datos y corrige la causa antes de continuar." & vbCrLf & _
        "Para reanudar SOLO la sentencia fallida y las siguientes, ingresa el ID del personaje. Cancelar conserva los datos pendientes.", "Recuperacion de guardado")
    If Len(answer) = 0 Then Exit Sub
    If Not IsNumeric(answer) Then Exit Sub
    If CDbl(answer) <= 0 Or CDbl(answer) > 2147483647# Or Fix(CDbl(answer)) <> CDbl(answer) Then Exit Sub
    id = CLng(answer)
    If MsgBox("Confirma que revisaste el error y el estado parcial del personaje " & id & ". No se repetiran las sentencias ya completadas.", vbYesNo Or vbDefaultButton2 Or vbExclamation, "Reanudar guardado") <> vbYes Then Exit Sub
    If Not RetryFailedCharacterSave(id) Then Call MsgBox("No se puede reanudar: el resultado es incierto o no hay un fallo confirmado. No reinicies el servidor; revisa el log.", vbExclamation)
    Exit Sub
Failed:
    Call LogDatabaseError("Save recovery UI error=" & Err.Number & " " & Err.Description)
End Sub

Public Sub StopCharacterSaveAdmission()
    stopping = True
End Sub

Public Property Get CharacterSaveShutdownRequested() As Boolean
    CharacterSaveShutdownRequested = stopping
End Property

#If UNIT_TEST = 1 Then
Public Function CharacterSaveTestJob(ByVal characterId As Long, Optional ByVal tail As Boolean = False) As clsCharacterSave
    Dim jobs As Collection
    Set jobs = queues(CStr(characterId))
    If tail Then
        Set CharacterSaveTestJob = jobs(jobs.Count)
    Else
        Set CharacterSaveTestJob = jobs(1)
    End If
End Function
#End If
