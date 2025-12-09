Attribute VB_Name = "ModLumberjack"
Option Explicit

Private Const MIN_STA_REQUIRED As Integer = 5

Public Sub ChopWood(ByRef AutomatedAction As t_AutomatedAction, ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .Stats.MinSta > MIN_STA_REQUIRED Then
            Call QuitarSta(UserIndex, MIN_STA_REQUIRED)
        Else
            Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
            .AutomatedAction.IsActive = False
            Exit Sub
        End If
        Dim skillPoints As Integer
        Dim res         As Integer
        Dim Suerte      As Integer
        skillPoints = .Stats.UserSkills(e_Skill.Talar)
        Suerte = Int(-0.00125 * skillPoints * skillPoints - 0.3 * skillPoints + 49)
        res = RandomNumber(1, IIf(MapInfo(UserList(UserIndex).pos.Map).Seguro = 1, Suerte + 4, Suerte))
        If res < 6 Then
            Dim nPos  As t_WorldPos
            Dim MiObj As t_Obj
            Call ActualizarRecurso(.pos.Map, .AutomatedAction.x, .AutomatedAction.y)
            MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.data = GetTickCountRaw() ' Ultimo uso
            If .clase = Trabajador Then
                MiObj.amount = GetExtractResourceForLevel(.Stats.ELV)
            Else
                MiObj.amount = RandomNumber(1, 2)
            End If
            MiObj.amount = MiObj.amount * SvrConfig.GetValue("RecoleccionMult")
            If ObjData(MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.ObjIndex).Elfico = 1 Then
                MiObj.ObjIndex = ElvenWood
            ElseIf ObjData(MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.ObjIndex).Pino = 1 Then
                MiObj.ObjIndex = PinoWood
            Else
                MiObj.ObjIndex = Wood
            End If
            If MiObj.amount > MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.amount Then
                MiObj.amount = MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.amount
            End If
            MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.amount = MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.amount - MiObj.amount
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.pos, MiObj)
            End If
            Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.charindex, vbWhite)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .pos.x, .pos.y))
            If IsFeatureEnabled("gain_exp_while_working") Then
                Call GiveExpWhileWorking(UserIndex, UserList(UserIndex).invent.EquippedWorkingToolObjIndex, e_JobsTypes.Woodcutter)
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)
            End If
        Else
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(64, .pos.x, .pos.y))
        End If
        Call SubirSkill(UserIndex, e_Skill.Talar)
    End With
End Sub
