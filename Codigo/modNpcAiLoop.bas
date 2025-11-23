Attribute VB_Name = "modNpcAiLoop"
Option Explicit

Public IntervaloNPCAI As Long
Public Const DEFAULT_NPC_AI_INTERVAL_MS As Long = 100
Private Const NPC_AI_PERFORMANCE_LIMIT_MS As Long = 600

Private lastNpcAiUpdate As Long

Public Sub InitializeNpcAiInterval(ByVal configuredIntervalMs As Long)
    IntervaloNPCAI = NormalizeNpcAiInterval(configuredIntervalMs)
    lastNpcAiUpdate = 0
End Sub

Public Sub MaybeUpdateNpcAI(ByVal currentTime As Long)
    If Not ShouldRunNpcAi(currentTime) Then Exit Sub
    Call RunNpcAiUpdate
End Sub

Private Function ShouldRunNpcAi(ByVal currentTime As Long) As Boolean
    If lastNpcAiUpdate = 0 Then
        lastNpcAiUpdate = currentTime
        Exit Function
    End If

    If TicksElapsed(lastNpcAiUpdate, currentTime) < IntervaloNPCAI Then Exit Function

    lastNpcAiUpdate = currentTime
    ShouldRunNpcAi = True
End Function

Private Function NormalizeNpcAiInterval(ByVal intervalMs As Long) As Long
    If intervalMs <= 0 Then
        NormalizeNpcAiInterval = DEFAULT_NPC_AI_INTERVAL_MS
    Else
        NormalizeNpcAiInterval = intervalMs
    End If
End Function

Private Sub RunNpcAiUpdate()
    On Error GoTo ErrorHandler
    Dim NpcIndex         As Long
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)

    If Not haciendoBK Then
        For NpcIndex = 1 To LastNPC
            With NpcList(NpcIndex)
                If .pos.Map > 0 Then
                    If MapInfo(.pos.Map).NumUsers > 0 Or MapInfo(.pos.Map).ForceUpdate Then
                        If .flags.NPCActive Then
                            If .npcType = DummyTarget Then
                                Call NpcDummyUpdate(NpcIndex)
                            Else
                                If .flags.Paralizado > 0 Then Call EfectoParalisisNpc(NpcIndex)
                                If .flags.Inmovilizado > 0 Then Call EfectoInmovilizadoNpc(NpcIndex)
                                If IntervaloPermiteMoverse(NpcIndex) Then Call NpcAI(NpcIndex)
                            End If
                        End If
                    End If
                End If
            End With
        Next NpcIndex
    End If

    Call PerformTimeLimitCheck(PerformanceTimer, "NpcAI_Update", NPC_AI_PERFORMANCE_LIMIT_MS)
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description & vbNewLine & "NPC: " & NpcList(NpcIndex).name & " en la posicion: " & NpcList(NpcIndex).pos.Map & "-" & NpcList(NpcIndex).pos.x _
            & "-" & NpcList(NpcIndex).pos.y, "modNpcAiLoop.MaybeUpdateNpcAI", Erl)
    Call MuereNpc(NpcIndex, 0)
End Sub
