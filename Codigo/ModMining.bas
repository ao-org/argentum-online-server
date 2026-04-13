Attribute VB_Name = "ModMining"
' Argentum 20 Game Server
'
'    Copyright (C) 2026 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
Option Explicit

Public Const BLODIUM_PICKAXE_REQUIRED_MSG As Integer = 597

Public Function CanUserExtractMinerals(ByVal UserIndex As Integer, ByVal TargetX As Byte, ByVal TargetY As Byte) As Boolean
    With UserList(UserIndex)
        If .invent.EquippedWeaponObjIndex <= 0 Then Exit Function
        If ObjData(MapData(.pos.Map, TargetX, TargetY).ObjInfo.ObjIndex).Blodium Then
            If Not ObjData(.invent.EquippedWeaponObjIndex).Blodium Then
                Call WriteLocaleMsg(UserIndex, BLODIUM_PICKAXE_REQUIRED_MSG, FONTTYPE_INFO)
                Exit Function
            End If
        End If
    End With
    CanUserExtractMinerals = True
End Function


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
                'dont call to ResetUserAutomatedAction(UserIndex) because .Automated.x and .Automated.y are being used
                .AutomatedAction.IsActive = False
                .Counters.Trabajando = 0
            End If
            MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.amount = MapData(.pos.Map, .AutomatedAction.x, .AutomatedAction.y).ObjInfo.amount - MiObj.amount
            If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.pos, MiObj)
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
            Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.charindex, vbWhite)
            Call WriteLocaleMsg(UserIndex, MSG_EXTRACTED_SOME_MINERALS, e_FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(e_SoundEffects.OldMiningPickaxeHit, .pos.x, .pos.y))
            If IsFeatureEnabled("gain_exp_while_working") Then
                Call GiveExpWhileWorking(UserIndex, MiObj, e_JobsTypes.Miner)
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)
            End If
        Else
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(e_SoundEffects.FailToExtractOre, .pos.x, .pos.y))
        End If
        .Counters.Trabajando = .Counters.Trabajando + 1
        Call SubirSkill(UserIndex, e_Skill.Mineria)
    End With
    Exit Sub
End Sub
