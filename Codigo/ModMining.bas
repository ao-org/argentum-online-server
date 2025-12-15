Attribute VB_Name = "ModMining"
Option Explicit

Public Sub MineMinerals(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If Not DecreaseUserStamina(UserIndex, ModAutomatedActions.MIN_STA_REQUIRED) Then
            Exit Sub
        End If
        Dim Suerte     As Integer
        Dim res        As Integer
        Dim Metal      As Integer
        Dim Yacimiento As t_ObjData
        Dim skill      As Integer
        skill = .Stats.UserSkills(e_Skill.Mineria)
        Suerte = Int(-0.00125 * skill * skill - 0.3 * skill + 49)
        res = RandomNumber(1, IIf(MapInfo(UserList(UserIndex).pos.Map).Seguro = 1, Suerte + 2, Suerte))
        If res <= 5 Then
            Dim MiObj As t_Obj
            Dim nPos  As t_WorldPos
            Call ActualizarRecurso(.pos.Map, .AutomatedAction.x, .AutomatedAction.y)
            MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.data = GetTickCountRaw() ' Ultimo uso
            Yacimiento = ObjData(MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.ObjIndex)
            MiObj.ObjIndex = Yacimiento.MineralIndex
            If .clase = Trabajador Then
                MiObj.amount = GetExtractResourceForLevel(.Stats.ELV)
            Else
                MiObj.amount = RandomNumber(1, 2)
            End If
            MiObj.amount = MiObj.amount * SvrConfig.GetValue("RecoleccionMult")
            If MiObj.amount > MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.amount Then
                MiObj.amount = MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.amount
            End If
            MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.amount = MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.amount - MiObj.amount
            If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.pos, MiObj)
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
            Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.charindex, vbWhite)
            Call WriteLocaleMsg(UserIndex, 651, e_FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(e_SoundEffects.OldMiningPickaxeHit, .pos.x, .pos.y))
            If IsFeatureEnabled("gain_exp_while_working") Then
                Call GiveExpWhileWorking(UserIndex, UserList(UserIndex).invent.EquippedWorkingToolObjIndex, e_JobsTypes.Miner)
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)
            End If
        Else
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(e_SoundEffects.FailToExtractOre, .pos.x, .pos.y))
        End If
        Call SubirSkill(UserIndex, e_Skill.Mineria)
    End With
    Exit Sub
End Sub
