Attribute VB_Name = "ModAutomatedActions"
Option Explicit
Private m_LastAutomatedActionAttempt As Long
Public Const MIN_STA_REQUIRED        As Integer = 5
Public Const CLOSE_DISTANCE_EXTRACTION = 1

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
                    Select Case .AutomatedAction.skill
                        Case e_Skill.Talar
                            Call ChopWood(UserIndex)
                        Case e_Skill.Mineria
                            Call MineMinerals(UserIndex)
                        Case e_Skill.Pescar
                            Call PerformFishing(UserIndex)
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

Public Function DecreaseUserStamina(ByVal UserIndex As Integer, ByVal StaminaRequirement As Integer) As Boolean
    With UserList(UserIndex)
        If .Stats.MinSta > StaminaRequirement Then
            Call QuitarSta(UserIndex, StaminaRequirement)
            DecreaseUserStamina = True
        Else
            'Msg2129=¡No tengo energía!
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareLocalizedChatOverHead(2129, UserList(UserIndex).Char.charindex, vbWhite))
            'Msg93=Estás muy cansado
            Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
            Call ResetUserAutomatedActions(UserIndex)
            DecreaseUserStamina = False
        End If
    End With
End Function

Public Function CheckResourceDistance(ByVal UserIndex As Integer, ByVal MaxExtractionDistance, ByVal TargetX As Integer, ByVal TargetY As Integer) As Boolean
    With UserList(UserIndex)
        Dim playerPosition As t_WorldPos
        playerPosition.Map = .pos.Map
        playerPosition.x = .pos.x
        playerPosition.y = .pos.y
        Dim resource As t_WorldPos
        resource.x = TargetX
        resource.y = TargetY
        resource.Map = .pos.Map
        If Distancia(playerPosition, resource) > MaxExtractionDistance Then
            Exit Function
        End If
        CheckResourceDistance = True
    End With
End Function

Public Function ValidResourceAtPos(ByVal UserIndex As Integer, ByVal WorkingToolIndex As e_OBJType, ByVal TargetX As Integer, ByVal TargetY As Integer)
    With UserList(UserIndex)
        If MapData(.pos.Map, TargetX, TargetY).ObjInfo.ObjIndex = 0 Then
            Exit Function
        End If
        If ObjData(MapData(.pos.Map, TargetX, TargetY).ObjInfo.ObjIndex).OBJType <> WorkingToolIndex Then
            Exit Function
        End If
        ValidResourceAtPos = True
    End With
End Function

Public Function CanUserExtractResource(ByVal UserIndex As Integer, ByVal ResourceType As e_OBJType, ByVal TargetX As Integer, ByVal TargetY As Integer) As Boolean
    If Not ValidResourceAtPos(UserIndex, ResourceType, TargetX, TargetY) Then
        Exit Function
    End If
    If Not CheckResourceDistance(UserIndex, CLOSE_DISTANCE_EXTRACTION, TargetX, TargetY) Then
        Call WriteLocaleMsg(UserIndex, 324, e_FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
    CanUserExtractResource = True
End Function
