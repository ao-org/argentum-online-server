Attribute VB_Name = "AutomatedActions"
Option Explicit


Private m_LastAutomatedActionAttempt As Long

Public Sub RunAutomatedActions()
    On Error GoTo RunAutomatedActions_Err
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    If m_LastAutomatedActionAttempt = 0 Then
        m_LastAutomatedActionAttempt = nowRaw
        Exit Sub
    End If
    If TicksElapsed(m_LastAutomatedActionAttempt, nowRaw) < IntervalAutomatedAction Then Exit Sub
    m_LastAutomatedActionAttempt = nowRaw
    Dim UserIndex        As Integer
    Dim PerformanceTimer As Long
    Call PerformanceTestStart(PerformanceTimer)
    For UserIndex = 1 To LastUser
        With UserList(UserIndex)
            If .flags.UserLogged Then
                If .AutomatedAction.IsActive Then
                    Debug.Print "User doing automated action " & UserIndex
                    Select Case .AutomatedAction.skill
                        Case e_Skill.Talar
                            Call ChopWood(.AutomatedAction, UserIndex)
                        Case Else
                            Debug.Assert False
                    End Select
                End If
            End If
        End With
    Next
    Call PerformTimeLimitCheck(PerformanceTimer, "AutomatedActions.RunAutomatedActions", 100)
    Exit Sub
RunAutomatedActions_Err:
    Call TraceError(Err.Number, Err.Description, "AutomatedActions.RunAutomatedActions", Erl)
End Sub

Public Sub StartAutomatedAction(ByVal x As Byte, ByVal y As Byte, ByVal skill As e_Skill, ByVal UserIndex As Integer)
    With UserList(UserIndex)
        .AutomatedAction.x = x
        .AutomatedAction.y = y
        .AutomatedAction.skill = skill
        .AutomatedAction.StartingTime = GetTickCountRaw()
        .AutomatedAction.IsActive = True
    End With
End Sub
