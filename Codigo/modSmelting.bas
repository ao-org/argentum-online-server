Attribute VB_Name = "modSmelting"
Option Explicit

Public Function CanUserSmelt(ByVal UserIndex As Integer, ByVal ResourceType As e_OBJType, ByVal TargetX As Integer, ByVal TargetY As Integer) As Boolean
    On Error GoTo CanUserSmelt_Err
    CanUserSmelt = False
    With UserList(UserIndex)
        If .clase <> e_Class.Trabajador Then
            Call WriteLocaleMsg(UserIndex, 607, e_FontTypeNames.FONTTYPE_INFO)
            Call ResetUserAutomatedActions(UserIndex)
            Exit Function
        End If
        If .flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios) Then
            Call ResetUserAutomatedActions(UserIndex)
            Exit Function
        End If
        If Not CheckResourceDistance(UserIndex, MEDIUM_DISTANCE_EXTRACTION, TargetX, TargetY) Then
            Call WriteLocaleMsg(UserIndex, 424, e_FontTypeNames.FONTTYPE_INFO)
            Call ResetUserAutomatedActions(UserIndex)
            Exit Function
        End If
        If MapData(.pos.Map, TargetX, TargetY).ObjInfo.ObjIndex = 0 Then
            Call ResetUserAutomatedActions(UserIndex)
            Exit Function
        End If
        If ObjData(MapData(.pos.Map, TargetX, TargetY).ObjInfo.ObjIndex).OBJType <> otForge Then
            Call ResetUserAutomatedActions(UserIndex)
            Exit Function
        End If
        If .flags.TargetObjInvIndex = 0 Then
            Call ResetUserAutomatedActions(UserIndex)
            Exit Function
        End If
    End With
    CanUserSmelt = True
    Exit Function
CanUserSmelt_Err:
    Call TraceError(Err.Number, Err.Description, "ModSmelting.CanUserSmelt", Erl)
End Function

Public Sub SmeltMinerals(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim RequiredSkill As Integer
        RequiredSkill = ObjData(.flags.TargetObjInvIndex).MinSkill
        If RequiredSkill > 100 Then
            Call WriteLocaleMsg(UserIndex, 608, e_FontTypeNames.FONTTYPE_INFO)
            Call ResetUserAutomatedActions(UserIndex)
            Exit Sub
        End If
        If .Stats.UserSkills(e_Skill.Mineria) >= RequiredSkill Then
            Call CraftIngots(UserIndex)
        Else
            Call WriteLocaleMsg(UserIndex, 1449, e_FontTypeNames.FONTTYPE_INFO)  ' Msg1449=No tenés conocimientos de minería suficientes para trabajar este mineral. Necesitas ¬1 puntos en minería.
            Call ResetUserAutomatedActions(UserIndex)
            Exit Sub
        End If
    End With
End Sub

Public Sub CraftIngots(ByVal UserIndex As Integer)
    On Error GoTo CraftIngots_Err
    Dim Slot       As Integer
    Dim obji       As Integer
    Dim cant       As Byte
    Dim necesarios As Integer
    If UserList(UserIndex).Stats.MinSta > 2 Then
        Call QuitarSta(UserIndex, 2)
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareLocalizedChatOverHead(2129, UserList(UserIndex).Char.charindex, vbWhite))
        Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
        Call WriteMacroTrabajoToggle(UserIndex, False)
        Exit Sub
    End If
    Slot = UserList(UserIndex).flags.TargetObjInvSlot
    obji = UserList(UserIndex).invent.Object(Slot).ObjIndex
    cant = RandomNumber(10, 20)
    necesarios = MineralsRequiredPerIngot(obji, cant)
    If UserList(UserIndex).invent.Object(Slot).Amount < MineralsRequiredPerIngot(obji, cant) Or ObjData(obji).OBJType <> e_OBJType.otMinerals Then
        ' Msg645=No tienes suficientes minerales para hacer un lingote.
        Call WriteLocaleMsg(UserIndex, 645, e_FontTypeNames.FONTTYPE_INFO)
        Call ResetUserAutomatedActions(UserIndex)
        Exit Sub
    End If
    UserList(UserIndex).invent.Object(Slot).Amount = UserList(UserIndex).invent.Object(Slot).Amount - MineralsRequiredPerIngot(obji, cant)
    If UserList(UserIndex).invent.Object(Slot).Amount < 1 Then
        UserList(UserIndex).invent.Object(Slot).Amount = 0
        UserList(UserIndex).invent.Object(Slot).ObjIndex = 0
    End If
    Dim nPos  As t_WorldPos
    Dim MiObj As t_Obj
    MiObj.Amount = cant
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)
    End If
    Call UpdateUserInv(False, UserIndex, Slot)
    Call WriteTextCharDrop(UserIndex, "+" & cant, UserList(UserIndex).Char.charindex, vbWhite)
    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(41, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
    Call SubirSkill(UserIndex, e_Skill.Mineria)
    UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
    Exit Sub
CraftIngots_Err:
    Call TraceError(Err.Number, Err.Description, "ModSmelting.CraftIngots", Erl)
End Sub

Private Function MineralsRequiredPerIngot(ByVal Lingote As e_Minerales, ByVal cant As Byte) As Integer
    On Error GoTo MineralsRequiredPerIngot_Err
    Select Case Lingote
        Case e_Minerales.HierroCrudo
            MineralsRequiredPerIngot = 13 * cant
        Case e_Minerales.PlataCruda
            MineralsRequiredPerIngot = 25 * cant
        Case e_Minerales.OroCrudo
            MineralsRequiredPerIngot = 50 * cant
        Case Else
            MineralsRequiredPerIngot = 10000
    End Select
    Exit Function
MineralsRequiredPerIngot_Err:
    Call TraceError(Err.Number, Err.Description, "ModSmelting.MineralsRequiredPerIngot", Erl)
End Function
