Attribute VB_Name = "modFishing"
Option Explicit

Private FishingLevelBonus() As Double
Private FishingBonusesInitialized As Boolean

Public Const OBJ_FISHING_ROD_BASIC                             As Integer = 881
Public Const OBJ_FISHING_ROD_COMMON                            As Integer = 2121
Public Const OBJ_FISHING_ROD_FINE                              As Integer = 2132
Public Const OBJ_FISHING_ROD_ELITE                             As Integer = 2133
Public Const OBJ_BROKEN_FISHING_ROD_BASIC                      As Integer = 3457
Public Const OBJ_BROKEN_FISHING_ROD_COMMON                     As Integer = 3456
Public Const OBJ_BROKEN_FISHING_ROD_FINE                       As Integer = 3459
Public Const OBJ_BROKEN_FISHING_ROD_ELITE                      As Integer = 3458
Public Const OBJ_FISHING_NET_BASIC                             As Integer = 138
Public Const OBJ_FISHING_NET_ELITE                             As Integer = 139
Public Const OBJ_FISHING_LINE                                  As Integer = 2183
Public Const OBJ_FISH_BANK                                     As Integer = 1992
Public Const OBJ_SQUID_BANK                                    As Integer = 1990
Public Const OBJ_SHRIMP_BANK                                   As Integer = 1991
Public Const OBJ_FISH_AREA                                     As Integer = 3740



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

Public Sub PerformFishing(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    ' Early validation
    If Not IsValidUserIndex(UserIndex) Then
        Call TraceError(1001, "Invalid user index in PerformFishing:  " & UserIndex, "modFishing.PerformFishing", Erl)
        Call ResetUserAutomatedActions(UserIndex)
        Exit Sub
    End If
    If Not FishingBonusesInitialized Then
        Call TraceError(1002, "Fishing bonuses were not initialized before use", "modFishing.PerformFishing", Erl)
        Call ResetUserAutomatedActions(UserIndex)
        Call InitializeFishingBonuses
    End If
    With UserList(UserIndex)
        Debug.Assert .AutomatedAction.x <> 0
        Debug.Assert .AutomatedAction.y <> 0
        ' Check stamina
        If Not DecreaseUserStamina(UserIndex, ModAutomatedActions.MIN_STA_REQUIRED) Then
            Exit Sub
        End If
        ' Update animation
        Dim sendTarget As sendTarget
        sendTarget = IIf(MapInfo(.pos.Map).Seguro = 1, ToIndex, ToPCAliveArea)
        Call SendData(sendTarget, UserIndex, PrepareMessageArmaMov(.Char.charindex, 0))
        ' Validate fishing tool
        Dim WorkingToolIndex As Integer
        WorkingToolIndex = .invent.EquippedWorkingToolObjIndex
        If Not IsValidObjectIndex(WorkingToolIndex) Then
            Call TraceError(1004, "Invalid fishing tool index: " & WorkingToolIndex, "modFishing.PerformFishing", Erl)
            Call ResetUserAutomatedActions(UserIndex)
            Exit Sub
        End If
        ' Determine tool type
        Dim IsUsingFishingNet As Boolean
        Select Case ObjData(WorkingToolIndex).Subtipo
            Case e_WorkingToolSubType.FishingRod
                IsUsingFishingNet = False
            Case e_WorkingToolSubType.FishingNet
                IsUsingFishingNet = True
            Case Else
                Call TraceError(1004, "Invalid fishing tool index: " & WorkingToolIndex, "modFishing.PerformFishing", Erl)
                Call ResetUserAutomatedActions(UserIndex)
                Exit Sub
        End Select
        ' Calculate bonuses and rewards
        Dim fishingLevel    As Long
        Dim levelBonus      As Double
        Dim fishingRodBonus As Double
        Dim totalBonus      As Double
        fishingLevel = ClampFishingLevel(.Stats.ELV)
        levelBonus = 1 + FishingLevelBonus(fishingLevel)
        fishingRodBonus = PoderCanas(ObjData(WorkingToolIndex).Power) / 10
        totalBonus = fishingRodBonus * levelBonus * SvrConfig.GetValue("RecoleccionMult")
        If MapInfo(.pos.Map).Seguro <> 0 Then
            totalBonus = totalBonus * PorcentajePescaSegura / 100
        End If
        ' Attempt to catch fish
        Dim fishingChance As Integer
        Dim caughtFish    As Boolean
        fishingChance = GetFishingChance(.Stats.UserSkills(e_Skill.Pescar))
        caughtFish = RandomNumber(1, 100) <= fishingChance
        If Not caughtFish Then
            Call SendData(ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, GRH_FALLO_PESCA))
            GoTo SkillImprovement
        End If
        ' Determine what fish was caught
        Dim fishingCatch As t_Obj
        fishingCatch.ObjIndex = ObtenerPezRandom(ObjData(WorkingToolIndex).Power)
        If .clase = e_Class.Trabajador Then
            If IsUsingFishingNet Then
                fishingCatch.amount = RandomNumber(2, 6)
            Else
                fishingCatch.amount = RandomNumber(1, 3)
            End If
            ' Award experience if enabled
            If IsFeatureEnabled("gain_exp_while_working") Then
                Call GiveExpWhileWorking(UserIndex, WorkingToolIndex, e_JobsTypes.Fisherman)
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)
            End If
        Else
            fishingCatch.amount = 1
        End If
        ' Handle unique map fish replacement
        If IsUniqueMapFish(fishingCatch.ObjIndex) And .pos.Map <> SvrConfig.GetValue("FISHING_MAP_SPECIAL_FISH1_ID") Then
            fishingCatch.ObjIndex = SvrConfig.GetValue("FISHING_SPECIALFISH1_REMPLAZO_ID")
            If MapInfo(.pos.Map).Seguro = 0 Then
                Dim NpcIndex As Integer
                NpcIndex = SpawnNpc(SvrConfig.GetValue("NPC_WATCHMAN_ID"), .pos, True, False)
            End If
        End If
        ' Check if this is a special fish (minigame)
        Dim isSpecialFish As Boolean
        isSpecialFish = False
        If HasSpecialFishDefinitions() Then
            Dim i As Long
            For i = 1 To UBound(PecesEspeciales)
                If PecesEspeciales(i).ObjIndex = fishingCatch.ObjIndex Then
                    isSpecialFish = True
                    Exit For
                End If
            Next i
        End If
        If isSpecialFish Then
            .flags.PescandoEspecial = True
            .Stats.NumObj_PezEspecial = fishingCatch.ObjIndex
            Call WritePelearConPezEspecial(UserIndex)
            Call ResetUserAutomatedActions(UserIndex)
            Exit Sub
        End If
        ' Handle fishing pool depletion (only in non-safe zones)
        If MapInfo(.pos.Map).Seguro = 0 Then
            Dim fishingPoolId As Integer
            Dim TargetX       As Integer
            Dim TargetY       As Integer
            TargetX = .AutomatedAction.x
            TargetY = .AutomatedAction.y
            fishingPoolId = SvrConfig.GetValue("FISHING_POOL_ID")
            If fishingPoolId > 0 And IsValidMapPosition(.pos.Map, TargetX, TargetY) Then
                If fishingPoolId = MapData(.pos.Map, TargetX, TargetY).ObjInfo.ObjIndex Then
                    If fishingCatch.amount > MapData(.pos.Map, TargetX, TargetY).ObjInfo.amount Then
                        fishingCatch.amount = MapData(.pos.Map, TargetX, TargetY).ObjInfo.amount
                        Call CreateFishingPool(.pos.Map)
                        Call EraseObj(MapData(.pos.Map, TargetX, TargetY).ObjInfo.amount, .pos.Map, TargetX, TargetY)
                        Call WriteLocaleMsg(UserIndex, 649, e_FontTypeNames.FONTTYPE_INFO)
                        .AutomatedAction.IsActive = False
                        .Counters.Trabajando = 0
                    End If
                    MapData(.pos.Map, TargetX, TargetY).ObjInfo.amount = MapData(.pos.Map, TargetX, TargetY).ObjInfo.amount - fishingCatch.amount
                End If
            End If
        End If
        ' Show particle effect and give fish to player
        If fishingCatch.ObjIndex = 0 Then
            GoTo SkillImprovement
        End If
        Call SendData(ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, ObjData(fishingCatch.ObjIndex).GrhIndex))
        If Not MeterItemEnInventario(UserIndex, fishingCatch) Then
            Call ResetUserAutomatedActions(UserIndex)
            GoTo SkillImprovement
        End If
        Call WriteTextCharDrop(UserIndex, "+" & fishingCatch.amount, .Char.charindex, vbWhite)
        Call SendData(ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .pos.x, .pos.y))
        ' Try to award bonus items
        If HasSpecialFishingRewards() Then
            Dim specialRoll As Long
            For i = 1 To UBound(EspecialesPesca)
                specialRoll = RandomNumber(1, IIf(IsUsingFishingNet, EspecialesPesca(i).data * 2, EspecialesPesca(i).data))
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
SkillImprovement:
        ' Improve fishing skill (only outside safe zones)
        If MapInfo(.pos.Map).Seguro = 0 Then
            Call SubirSkill(UserIndex, e_Skill.Pescar)
        End If
        .Counters.Trabajando = .Counters.Trabajando + 1
    End With
    Exit Sub
ErrHandler:
    Call LogError("Error in PerformFishing.  Error " & Err.Number & " - " & Err.Description & " Line number: " & Erl)
End Sub

Private Function IsValidUserIndex(ByVal UserIndex As Integer) As Boolean
    On Error GoTo InvalidIndex
    Dim LowerBound As Long
    Dim UpperBound As Long

    LowerBound = LBound(UserList)
    UpperBound = UBound(UserList)

    If UserIndex < LowerBound Or UserIndex > UpperBound Then Exit Function

    IsValidUserIndex = True
    Exit Function
InvalidIndex:
    Err.Clear
    IsValidUserIndex = False
End Function

Private Function IsValidObjectIndex(ByVal objectIndex As Integer) As Boolean
    On Error GoTo InvalidIndex
    Dim LowerBound As Long
    Dim UpperBound As Long

    LowerBound = LBound(ObjData)
    UpperBound = UBound(ObjData)

    If objectIndex < LowerBound Or objectIndex > UpperBound Then Exit Function

    IsValidObjectIndex = True
    Exit Function
InvalidIndex:
    Err.Clear
    IsValidObjectIndex = False
End Function

Private Function IsValidMapIndex(ByVal mapIndex As Integer) As Boolean
    On Error GoTo InvalidIndex
    Dim LowerBound As Long
    Dim UpperBound As Long

    LowerBound = LBound(MapInfo)
    UpperBound = UBound(MapInfo)

    If mapIndex < LowerBound Or mapIndex > UpperBound Then Exit Function

    IsValidMapIndex = True
    Exit Function
InvalidIndex:
    Err.Clear
    IsValidMapIndex = False
End Function

Public Function IsValidMapPosition(ByVal mapIndex As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    If Not IsValidMapIndex(mapIndex) Then Exit Function
    If x < XMinMapSize Or x > XMaxMapSize Then Exit Function
    If y < YMinMapSize Or y > YMaxMapSize Then Exit Function
    IsValidMapPosition = True
End Function

Private Function ClampFishingLevel(ByVal level As Long) As Long
    Dim LowerBound As Long
    Dim UpperBound As Long

    LowerBound = LBound(FishingLevelBonus)
    UpperBound = UBound(FishingLevelBonus)

    If level < LowerBound Then
        ClampFishingLevel = LowerBound
    ElseIf level > UpperBound Then
        ClampFishingLevel = UpperBound
    Else
        ClampFishingLevel = level
    End If
End Function

Private Function GetFishingChance(ByVal FishingSkill As Integer) As Integer
    Select Case FishingSkill
        Case Is < 20
            GetFishingChance = 20
        Case Is < 40
            GetFishingChance = 35
        Case Is < 70
            GetFishingChance = 55
        Case Is < 100
            GetFishingChance = 68
        Case Else
            GetFishingChance = 80
    End Select
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

Public Function ObtenerPezRandom(ByVal PoderCania As Integer) As Long
    On Error GoTo ObtenerPezRandom_Err

    Dim PesoMinimo As Long
    Dim PesoMaximo As Long
    Dim ValorGenerado As Long
    Dim PezIndex As Long

    ' Aseguramos que PoderCania esté dentro del rango válido del array.
    PoderCania = Clamp(PoderCania, LBound(PesoPeces), UBound(PesoPeces))
    
    ' PesoMaximo: suma de pesos acumulados de todos los peces que puede pescar esta caña
    PesoMaximo = PesoPeces(PoderCania)
    
    ' Esto asegura que el aleatorio solo considere los peces que pertenecen al Power actual
    If PoderCania > LBound(PesoPeces) Then
        PesoMinimo = PesoPeces(PoderCania - 1)
    Else
        PesoMinimo = 0
    End If

    ' Generamos un valor aleatorio solo dentro del rango correspondiente
    If PesoMaximo <= PesoMinimo Then
        ValorGenerado = RandomNumber(0, PesoMaximo - 1)
    Else
        ValorGenerado = RandomNumber(PesoMinimo, PesoMaximo - 1)
    End If

    ' Obtenemos el pez correspondiente
    PezIndex = BinarySearchPeces(ValorGenerado) ' BinarySearchPeces() espera un valor en el mismo espacio acumulado que PesoPeces().
    ObtenerPezRandom = Peces(PezIndex).ObjIndex

    Exit Function

ObtenerPezRandom_Err:
    Call TraceError(Err.Number, Err.Description, "modFishing.ObtenerPezRandom", Erl)
End Function
Public Function IsUniqueMapFish(ByVal ObjIndex As Long) As Boolean
On Error GoTo IsUniqueMapFish_Err
    Dim i As Long
    For i = 1 To UniqueMapFishCount
        If UniqueMapFishIDs(i) = ObjIndex Then
            IsUniqueMapFish = True
            Exit Function
        End If
    Next
    Exit Function
IsUniqueMapFish_Err:
    Call TraceError(Err.Number, Err.Description, "modFishing.IsUniqueMapFish", Erl)
End Function

Public Function CanUserFish(ByVal UserIndex As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Boolean
    With UserList(UserIndex)
        Debug.Assert TargetX <> 0
        Debug.Assert TargetY <> 0
        If .invent.EquippedWorkingToolObjIndex = 0 Then
            Exit Function
        End If
        If ObjData(.invent.EquippedWorkingToolObjIndex).OBJType <> e_OBJType.otWorkingTools Then
            Exit Function
        End If
        If Not ValidateFishingPosition(UserIndex, TargetX, TargetY) Then
            Exit Function
        End If
        CanUserFish = True
    End With
End Function

Public Function ValidateFishingPosition(ByVal UserIndex As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Boolean
    ValidateFishingPosition = False
    With UserList(UserIndex)
        ' Check if target position has water flag
        If Not IsValidMapIndex(.pos.Map) Then
            Call TraceError(1003, "Invalid map index in PerformFishing: " & .pos.Map, "modFishing.PerformFishing", Erl)
            Exit Function
        End If
        If (MapData(.pos.Map, TargetX, TargetY).Blocked And FLAG_AGUA) = 0 Then
            ' No water at target position
            Call WriteLocaleMsg(UserIndex, 596, e_FontTypeNames.FONTTYPE_INFO)  ' Zona de pesca no Autorizada
            Exit Function
        End If
        ' Check for invalid fishing trigger
        If MapData(.pos.Map, .pos.x, .pos.y).trigger = e_Trigger.PESCAINVALIDA Then
            Call WriteLocaleMsg(UserIndex, 596, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        Select Case ObjData(.invent.EquippedWorkingToolObjIndex).Subtipo
            Case e_WorkingToolSubType.FishingRod
                If IsStandingOnWater(.pos) Then
                    Call WriteLocaleMsg(UserIndex, 1436, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
                If Not IsAdjacentToWater(.pos) Then
                    ' Msg1021= Acércate a la costa para pescar.
                    Call WriteLocaleMsg(UserIndex, 1021, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
                If UserList(UserIndex).flags.Navegando <> 0 Then
                    Call WriteLocaleMsg(UserIndex, 1436, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            Case e_WorkingToolSubType.FishingNet
                If UserList(UserIndex).flags.Navegando = 0 Then
                    Call WriteLocaleMsg(UserIndex, 1436, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
                If (MapData(.pos.Map, TargetX, TargetY).ObjInfo.ObjIndex <> OBJ_FISH_AREA) And _
                   (MapData(.pos.Map, TargetX, TargetY).ObjInfo.ObjIndex <> OBJ_SHRIMP_BANK) And _
                   (MapData(.pos.Map, TargetX, TargetY).ObjInfo.ObjIndex <> OBJ_SQUID_BANK) And _
                   (MapData(.pos.Map, TargetX, TargetY).ObjInfo.ObjIndex <> OBJ_FISH_BANK) Then
                    Call WriteLocaleMsg(UserIndex, 595, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
                If Not CheckResourceDistance(UserIndex, CLOSE_DISTANCE_EXTRACTION, TargetX, TargetY) Then
                    Call WriteLocaleMsg(UserIndex, 424, e_FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                End If
            Case Else
                Debug.Assert False
                Call TraceError(0, "Invalid fishing tool: " & UserIndex, "modFishing.ValidateFishingPosition", Erl)
                Exit Function
        End Select
        If MapInfo(.pos.Map).zone = "DUNGEON" Then
            Call WriteLocaleMsg(UserIndex, 596, e_FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        ValidateFishingPosition = True
    End With
End Function

' Helper function to check if position has water
Private Function IsStandingOnWater(ByRef pos As t_WorldPos) As Boolean
    IsStandingOnWater = (MapData(pos.Map, pos.x, pos.y).Blocked And FLAG_AGUA) <> 0
End Function

' Helper function to check if any adjacent tile has water
Private Function IsAdjacentToWater(ByRef pos As t_WorldPos) As Boolean
    IsAdjacentToWater = _
        (MapData(pos.Map, pos.x + 1, pos.y).Blocked And FLAG_AGUA) <> 0 Or _
        (MapData(pos.Map, pos.x - 1, pos.y).Blocked And FLAG_AGUA) <> 0 Or _
        (MapData(pos.Map, pos.x, pos.y + 1).Blocked And FLAG_AGUA) <> 0 Or _
        (MapData(pos.Map, pos.x, pos.y - 1).Blocked And FLAG_AGUA) <> 0
End Function

