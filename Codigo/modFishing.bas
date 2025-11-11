Attribute VB_Name = "modFishing"
Option Explicit

Private FishingLevelBonus() As Double
Private FishingBonusesInitialized As Boolean

Public Sub InitializeFishingBonuses()
    If FishingBonusesInitialized Then Exit Sub

    ReDim FishingLevelBonus(1 To 47) As Double

    FishingLevelBonus(1) = 0#
    FishingLevelBonus(2) = 0.009
    FishingLevelBonus(3) = 0.015
    FishingLevelBonus(4) = 0.019
    FishingLevelBonus(5) = 0.025
    FishingLevelBonus(6) = 0.03
    FishingLevelBonus(7) = 0.035
    FishingLevelBonus(8) = 0.04
    FishingLevelBonus(9) = 0.045
    FishingLevelBonus(10) = 0.05
    FishingLevelBonus(11) = 0.06
    FishingLevelBonus(12) = 0.07
    FishingLevelBonus(13) = 0.08
    FishingLevelBonus(14) = 0.09
    FishingLevelBonus(15) = 0.1
    FishingLevelBonus(16) = 0.11
    FishingLevelBonus(17) = 0.13
    FishingLevelBonus(18) = 0.14
    FishingLevelBonus(19) = 0.16
    FishingLevelBonus(20) = 0.18
    FishingLevelBonus(21) = 0.2
    FishingLevelBonus(22) = 0.22
    FishingLevelBonus(23) = 0.24
    FishingLevelBonus(24) = 0.27
    FishingLevelBonus(25) = 0.3
    FishingLevelBonus(26) = 0.32
    FishingLevelBonus(27) = 0.35
    FishingLevelBonus(28) = 0.37
    FishingLevelBonus(29) = 0.4
    FishingLevelBonus(30) = 0.43
    FishingLevelBonus(31) = 0.47
    FishingLevelBonus(32) = 0.51
    FishingLevelBonus(33) = 0.55
    FishingLevelBonus(34) = 0.58
    FishingLevelBonus(35) = 0.62
    FishingLevelBonus(36) = 0.7
    FishingLevelBonus(37) = 0.77
    FishingLevelBonus(38) = 0.84
    FishingLevelBonus(39) = 0.92
    FishingLevelBonus(40) = 1#
    FishingLevelBonus(41) = 1.1
    FishingLevelBonus(42) = 1.15
    FishingLevelBonus(43) = 1.3
    FishingLevelBonus(44) = 1.5
    FishingLevelBonus(45) = 1.8
    FishingLevelBonus(46) = 2#
    FishingLevelBonus(47) = 2.5

    FishingBonusesInitialized = True
End Sub

Public Sub PerformFishing(ByVal UserIndex As Integer, Optional ByVal UsingFishingNet As Boolean = False)
    On Error GoTo ErrHandler

    If Not IsValidUserIndex(UserIndex) Then
        Call TraceError(1001, "Invalid user index in PerformFishing: " & UserIndex, "modFishing.PerformFishing", Erl)
        Exit Sub
    End If

    If Not FishingBonusesInitialized Then
        Call TraceError(1002, "Fishing bonuses were not initialized before use", "modFishing.PerformFishing", Erl)
        Call InitializeFishingBonuses
    End If

    Dim staminaCost As Integer
    Dim fishingRodBonus As Double
    Dim levelBonus As Double
    Dim totalBonus As Double
    Dim reward As Double
    Dim fishingChance As Integer
    Dim caughtFish As Boolean
    Dim npcIndex As Integer
    Dim workingToolIndex As Integer
    Dim fishingLevel As Long
    Dim currentMap As Integer
    Dim isSpecialFish As Boolean
    Dim stopWorking As Boolean
    Dim objValue As Integer
    Dim specialRoll As Long
    Dim fishingPoolId As Integer
    Dim targetX As Integer
    Dim targetY As Integer
    Dim i As Long

    With UserList(UserIndex)
        staminaCost = IIf(UsingFishingNet, 12, RandomNumber(2, 3))
        If .flags.Privilegios And (e_PlayerType.Consejero) Then Exit Sub

        If .Stats.MinSta > staminaCost Then
            Call QuitarSta(UserIndex, staminaCost)
        Else
            Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub
        End If

        currentMap = .pos.Map
        If Not IsValidMapIndex(currentMap) Then
            Call TraceError(1003, "Invalid map index in PerformFishing: " & currentMap, "modFishing.PerformFishing", Erl)
            Exit Sub
        End If

        If MapInfo(currentMap).Seguro = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageArmaMov(.Char.charindex, 0))
        Else
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageArmaMov(.Char.charindex, 0))
        End If

        workingToolIndex = .invent.EquippedWorkingToolObjIndex
        If Not IsValidObjectIndex(workingToolIndex) Then
            Call TraceError(1004, "Invalid fishing tool index: " & workingToolIndex, "modFishing.PerformFishing", Erl)
            Exit Sub
        End If

        fishingLevel = ClampFishingLevel(.Stats.ELV)
        levelBonus = 1 + FishingLevelBonus(fishingLevel)
        fishingRodBonus = PoderCanas(ObjData(workingToolIndex).Power) / 10
        totalBonus = fishingRodBonus * levelBonus * SvrConfig.GetValue("RecoleccionMult")

        If MapInfo(currentMap).Seguro <> 0 Then
            totalBonus = totalBonus * PorcentajePescaSegura / 100
        End If

        reward = (IntervaloTrabajarExtraer / 3600000#) * 8000# * totalBonus * 1.2 * (1 + (RandomNumber(0, 20) - 10) / 100)

        fishingChance = GetFishingChance(.Stats.UserSkills(e_Skill.Pescar))
        caughtFish = RandomNumber(1, 100) <= fishingChance

        stopWorking = False
        If caughtFish Then
            If IsFeatureEnabled("gain_exp_while_working") Then
                Call GiveExpWhileWorking(UserIndex, workingToolIndex, e_JobsTypes.Fisherman)
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)
            End If

            Dim fishingCatch As t_Obj
            fishingCatch.ObjIndex = ObtenerPezRandom(ObjData(workingToolIndex).Power)

            If Not IsValidObjectIndex(fishingCatch.ObjIndex) Then
                Call TraceError(1005, "Invalid fish object index: " & fishingCatch.ObjIndex, "modFishing.PerformFishing", Erl)
                Exit Sub
            End If

            objValue = max(ObjData(fishingCatch.ObjIndex).Valor / 3, 1)

            Dim isSpecialFishCandidate As Boolean
            isSpecialFishCandidate = (fishingCatch.ObjIndex = SvrConfig.GetValue("FISHING_SPECIALFISH1_ID")) Or _
                                     (fishingCatch.ObjIndex = SvrConfig.GetValue("FISHING_SPECIALFISH2_ID"))

            If isSpecialFishCandidate And currentMap <> SvrConfig.GetValue("FISHING_MAP_SPECIAL_FISH1_ID") Then
                fishingCatch.ObjIndex = SvrConfig.GetValue("FISHING_SPECIALFISH1_REMPLAZO_ID")
                If MapInfo(currentMap).Seguro = 0 Then
                    npcIndex = SpawnNpc(SvrConfig.GetValue("NPC_WATCHMAN_ID"), .pos, True, False)
                End If
                Call WriteMacroTrabajoToggle(UserIndex, False)
            End If

            fishingCatch.amount = Round(reward / objValue)
            If fishingCatch.amount <= 0 Then
                fishingCatch.amount = 1
            End If

            targetX = .Trabajo.Target_X
            targetY = .Trabajo.Target_Y

            If MapInfo(currentMap).Seguro = 0 Then
                fishingPoolId = SvrConfig.GetValue("FISHING_POOL_ID")
                If fishingPoolId > 0 And IsValidMapPosition(currentMap, targetX, targetY) Then
                    If fishingPoolId = MapData(currentMap, targetX, targetY).ObjInfo.ObjIndex Then
                        If fishingCatch.amount > MapData(currentMap, targetX, targetY).ObjInfo.amount Then
                            fishingCatch.amount = MapData(currentMap, targetX, targetY).ObjInfo.amount
                            Call CreateFishingPool(currentMap)
                            Call EraseObj(MapData(currentMap, targetX, targetY).ObjInfo.amount, currentMap, targetX, targetY)
                            Call WriteLocaleMsg(UserIndex, 649, e_FontTypeNames.FONTTYPE_INFO)
                            stopWorking = True
                        End If
                        MapData(currentMap, targetX, targetY).ObjInfo.amount = MapData(currentMap, targetX, targetY).ObjInfo.amount - fishingCatch.amount
                    End If
                End If
            End If

            isSpecialFish = False
            If HasSpecialFishDefinitions() Then
                For i = 1 To UBound(PecesEspeciales)
                    If PecesEspeciales(i).ObjIndex = fishingCatch.ObjIndex Then
                        isSpecialFish = True
                        Exit For
                    End If
                Next i
            End If

            If Not isSpecialFish Then
                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, ObjData(fishingCatch.ObjIndex).GrhIndex))
            Else
                .flags.PescandoEspecial = True
                Call WriteMacroTrabajoToggle(UserIndex, False)
                .Stats.NumObj_PezEspecial = fishingCatch.ObjIndex
                Call WritePelearConPezEspecial(UserIndex)
                Exit Sub
            End If

            If fishingCatch.ObjIndex = 0 Then Exit Sub

            If Not MeterItemEnInventario(UserIndex, fishingCatch) Then
                stopWorking = True
            End If

            Call WriteTextCharDrop(UserIndex, "+" & fishingCatch.amount, .Char.charindex, vbWhite)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .pos.x, .pos.y))

            If HasSpecialFishingRewards() Then
                For i = 1 To UBound(EspecialesPesca)
                    specialRoll = RandomNumber(1, IIf(UsingFishingNet, EspecialesPesca(i).data * 2, EspecialesPesca(i).data))
                    If specialRoll = 1 Then
                        fishingCatch.ObjIndex = EspecialesPesca(i).ObjIndex
                        fishingCatch.amount = 1
                        If Not MeterItemEnInventario(UserIndex, fishingCatch) Then
                            Call TirarItemAlPiso(.pos, fishingCatch)
                        End If
                        Call WriteLocaleMsg(UserIndex, 1457, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Next i
            End If
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, GRH_FALLO_PESCA))
        End If

        If MapInfo(currentMap).Seguro = 0 Then
            Call SubirSkill(UserIndex, e_Skill.Pescar)
        End If

        If stopWorking Then
            Call WriteWorkRequestTarget(UserIndex, 0)
            Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub
        End If

        .Counters.Trabajando = .Counters.Trabajando + 1
        .Counters.LastTrabajo = Int(IntervaloTrabajarExtraer / 1000)

        If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
            Call WriteMacroTrabajoToggle(UserIndex, True)
        End If
    End With

    Exit Sub
ErrHandler:
    Call LogError("Error in PerformFishing. Error " & Err.Number & " - " & Err.Description & " Line number: " & Erl)
End Sub

Private Function IsValidUserIndex(ByVal UserIndex As Integer) As Boolean
    On Error GoTo InvalidIndex
    Dim lowerBound As Long
    Dim upperBound As Long

    lowerBound = LBound(UserList)
    upperBound = UBound(UserList)

    If UserIndex < lowerBound Or UserIndex > upperBound Then Exit Function

    IsValidUserIndex = True
    Exit Function
InvalidIndex:
    Err.Clear
    IsValidUserIndex = False
End Function

Private Function IsValidObjectIndex(ByVal ObjectIndex As Integer) As Boolean
    On Error GoTo InvalidIndex
    Dim lowerBound As Long
    Dim upperBound As Long

    lowerBound = LBound(ObjData)
    upperBound = UBound(ObjData)

    If ObjectIndex < lowerBound Or ObjectIndex > upperBound Then Exit Function

    IsValidObjectIndex = True
    Exit Function
InvalidIndex:
    Err.Clear
    IsValidObjectIndex = False
End Function

Private Function IsValidMapIndex(ByVal MapIndex As Integer) As Boolean
    On Error GoTo InvalidIndex
    Dim lowerBound As Long
    Dim upperBound As Long

    lowerBound = LBound(MapInfo)
    upperBound = UBound(MapInfo)

    If MapIndex < lowerBound Or MapIndex > upperBound Then Exit Function

    IsValidMapIndex = True
    Exit Function
InvalidIndex:
    Err.Clear
    IsValidMapIndex = False
End Function

Private Function IsValidMapPosition(ByVal MapIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    If Not IsValidMapIndex(MapIndex) Then Exit Function
    If X < XMinMapSize Or X > XMaxMapSize Then Exit Function
    If Y < YMinMapSize Or Y > YMaxMapSize Then Exit Function
    IsValidMapPosition = True
End Function

Private Function ClampFishingLevel(ByVal Level As Long) As Long
    Dim lowerBound As Long
    Dim upperBound As Long

    lowerBound = LBound(FishingLevelBonus)
    upperBound = UBound(FishingLevelBonus)

    If Level < lowerBound Then
        ClampFishingLevel = lowerBound
    ElseIf Level > upperBound Then
        ClampFishingLevel = upperBound
    Else
        ClampFishingLevel = Level
    End If
End Function

Private Function GetFishingChance(ByVal FishingSkill As Integer) As Integer
    If FishingSkill < 20 Then
        GetFishingChance = 20
    ElseIf FishingSkill < 40 Then
        GetFishingChance = 35
    ElseIf FishingSkill < 70 Then
        GetFishingChance = 55
    ElseIf FishingSkill < 100 Then
        GetFishingChance = 68
    Else
        GetFishingChance = 80
    End If
End Function

Private Function HasSpecialFishDefinitions() As Boolean
    On Error GoTo HandleError
    HasSpecialFishDefinitions = (UBound(PecesEspeciales) >= 1)
    Exit Function
HandleError:
    Err.Clear
    HasSpecialFishDefinitions = False
End Function

Private Function HasSpecialFishingRewards() As Boolean
    On Error GoTo HandleError
    HasSpecialFishingRewards = (UBound(EspecialesPesca) >= 1)
    Exit Function
HandleError:
    Err.Clear
    HasSpecialFishingRewards = False
End Function
