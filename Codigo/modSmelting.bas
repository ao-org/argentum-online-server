Attribute VB_Name = "modSmelting"
Option Explicit

Public Function CanUserSmelt(ByVal UserIndex As Integer) As Boolean
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
        If MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.ObjIndex = 0 Then
            Call ResetUserAutomatedActions(UserIndex)
            Exit Function
        End If
        If ObjData(MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.ObjIndex).OBJType <> otForge Then
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
    Call TraceError(Err.Number, Err.Description, "modSmelting.CanUserSmelt", Erl)
End Function

Public Sub SmeltMinerals(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim RequiredSkill As Integer
        RequiredSkill = ObjData(.flags.TargetObjInvIndex).MinSkill
        If ObjData(.flags.TargetObjInvIndex).OBJType = e_OBJType.otMinerals And .Stats.UserSkills(e_Skill.Mineria) >= RequiredSkill Then
            Call CraftIngots(UserIndex)
        ElseIf RequiredSkill > 100 Then
            ' Msg608=Los mortales no pueden fundir este mineral.
            Call WriteLocaleMsg(UserIndex, 608, e_FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteLocaleMsg(UserIndex, 1449, e_FontTypeNames.FONTTYPE_INFO)  ' Msg1449=No tenés conocimientos de minería suficientes para trabajar este mineral. Necesitas ¬1 puntos en minería.
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
        'Msg2129=¡No tengo energía!
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareLocalizedChatOverHead(2129, UserList(UserIndex).Char.charindex, vbWhite))
        'Msg93=Estás muy cansado.
        Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
        Call WriteMacroTrabajoToggle(UserIndex, False)
        Exit Sub
    End If
    Slot = UserList(UserIndex).flags.TargetObjInvSlot
    obji = UserList(UserIndex).invent.Object(Slot).ObjIndex
    cant = RandomNumber(10, 20)
    necesarios = MineralesParaLingote(obji, cant)
    If UserList(UserIndex).invent.Object(Slot).Amount < MineralesParaLingote(obji, cant) Or ObjData(obji).OBJType <> e_OBJType.otMinerals Then
        ' Msg645=No tienes suficientes minerales para hacer un lingote.
        Call WriteLocaleMsg(UserIndex, 645, e_FontTypeNames.FONTTYPE_INFO)
        Call ResetUserAutomatedActions(UserIndex)
        Exit Sub
    End If
    UserList(UserIndex).invent.Object(Slot).Amount = UserList(UserIndex).invent.Object(Slot).Amount - MineralesParaLingote(obji, cant)
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
    Call TraceError(Err.Number, Err.Description, "Trabajo.CraftIngots", Erl)
End Sub

Private Function MineralesParaLingote(ByVal Lingote As e_Minerales, ByVal cant As Byte) As Integer
    On Error GoTo MineralesParaLingote_Err
    Select Case Lingote
        Case e_Minerales.HierroCrudo
            MineralesParaLingote = 13 * cant
        Case e_Minerales.PlataCruda
            MineralesParaLingote = 25 * cant
        Case e_Minerales.OroCrudo
            MineralesParaLingote = 50 * cant
        Case Else
            MineralesParaLingote = 10000
    End Select
    Exit Function
MineralesParaLingote_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.MineralesParaLingote", Erl)
End Function
