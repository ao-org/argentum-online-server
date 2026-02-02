Attribute VB_Name = "SistemaCombate"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
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
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
Public Const MAXDISTANCIAARCO  As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18

Public Enum AttackType
    Ranged
    Melee
End Enum

Private Function ModificadorPoderAtaqueArmas(ByVal clase As e_Class) As Single
    On Error GoTo ModificadorPoderAtaqueArmas_Err
    ModificadorPoderAtaqueArmas = ModClase(clase).AtaqueArmas
    Exit Function
ModificadorPoderAtaqueArmas_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.ModificadorPoderAtaqueArmas", Erl)
End Function

Private Function ModificadorPoderAtaqueProyectiles(ByVal clase As e_Class) As Single
    On Error GoTo ModificadorPoderAtaqueProyectiles_Err
    ModificadorPoderAtaqueProyectiles = ModClase(clase).AtaqueProyectiles
    Exit Function
ModificadorPoderAtaqueProyectiles_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.ModificadorPoderAtaqueProyectiles", Erl)
End Function

Private Function ModicadorDañoClaseArmas(ByVal clase As e_Class) As Single
    On Error GoTo ModicadorDañoClaseArmas_Err
    ModicadorDañoClaseArmas = ModClase(clase).DañoArmas
    Exit Function
ModicadorDañoClaseArmas_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.ModicadorDañoClaseArmas", Erl)
End Function

Private Function ModicadorApuñalarClase(ByVal clase As e_Class) As Single
    On Error GoTo ModicadorApuñalarClase_Err
    ModicadorApuñalarClase = ModClase(clase).ModApunalar
    Exit Function
ModicadorApuñalarClase_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.ModicadorApuñalarClase", Erl)
End Function

Private Function GetStabbingNPCMinForClass(ByVal clase As e_Class) As Single
    On Error GoTo GetStabbingNPCMinForClass
    GetStabbingNPCMinForClass = ModClase(clase).ModStabbingNPCMin
    Exit Function
GetStabbingNPCMinForClass:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.GetStabbingNPCMinForClass", Erl)
End Function

Private Function GetStabbingNPCMaxForClass(ByVal clase As e_Class) As Single
    On Error GoTo GetStabbingNPCMaxForClass
    GetStabbingNPCMaxForClass = ModClase(clase).ModStabbingNPCMax
    Exit Function
GetStabbingNPCMaxForClass:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.GetStabbingNPCMaxForClass", Erl)
End Function

Private Function ModicadorDañoClaseProyectiles(ByVal clase As e_Class) As Single
    On Error GoTo ModicadorDañoClaseProyectiles_Err
    ModicadorDañoClaseProyectiles = ModClase(clase).DañoProyectiles
    Exit Function
ModicadorDañoClaseProyectiles_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.ModicadorDañoClaseProyectiles", Erl)
End Function

Private Function ModEvasionDeEscudoClase(ByVal clase As e_Class) As Single
    On Error GoTo ModEvasionDeEscudoClase_Err
    ModEvasionDeEscudoClase = ModClase(clase).Escudo
    Exit Function
ModEvasionDeEscudoClase_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.ModEvasionDeEscudoClase", Erl)
End Function

Private Function Minimo(ByVal a As Single, ByVal b As Single) As Single
    On Error GoTo Minimo_Err
    If a > b Then
        Minimo = b
        Else:
        Minimo = a
    End If
    Exit Function
Minimo_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.Minimo", Erl)
End Function

Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    On Error GoTo MinimoInt_Err
    If a > b Then
        MinimoInt = b
        Else:
        MinimoInt = a
    End If
    Exit Function
MinimoInt_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.MinimoInt", Erl)
End Function

Private Function Maximo(ByVal a As Single, ByVal b As Single) As Single
    On Error GoTo Maximo_Err
    If a > b Then
        Maximo = a
        Else:
        Maximo = b
    End If
    Exit Function
Maximo_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.Maximo", Erl)
End Function

Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
    On Error GoTo MaximoInt_Err
    If a > b Then
        MaximoInt = a
        Else:
        MaximoInt = b
    End If
    Exit Function
MaximoInt_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.MaximoInt", Erl)
End Function

Private Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long
    On Error GoTo PoderEvasionEscudo_Err
    With UserList(UserIndex)
        If .invent.EquippedShieldObjIndex <= 0 Then
            PoderEvasionEscudo = 0
            Exit Function
        End If
        Dim itemModifier As Single
        itemModifier = CSng(ObjData(.invent.EquippedShieldObjIndex).Porcentaje) / 100
        PoderEvasionEscudo = ((UserList(UserIndex).Stats.UserSkills(e_Skill.Defensa) * ModEvasionDeEscudoClase(UserList(UserIndex).clase)) / 2) * itemModifier
    End With
    Exit Function
PoderEvasionEscudo_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.PoderEvasionEscudo", Erl)
End Function

Private Function PoderEvasion(ByVal UserIndex As Integer) As Long
    On Error GoTo PoderEvasion_Err
    With UserList(UserIndex)
        PoderEvasion = (.Stats.UserSkills(e_Skill.Tacticas) + (3 * .Stats.UserSkills(e_Skill.Tacticas) / 100) * .Stats.UserAtributos(Agilidad)) * ModClase(.clase).Evasion
        PoderEvasion = PoderEvasion + (2.5 * Maximo(.Stats.ELV - 12, 0))
        PoderEvasion = PoderEvasion + UserMod.GetEvasionBonus(UserList(UserIndex))
    End With
    Exit Function
PoderEvasion_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.PoderEvasion", Erl)
End Function

Private Function AttackPower(ByVal UserIndex, ByVal Skill As Integer, ByVal skillModifier As Single) As Long
    On Error GoTo AttackPower_Err
    Dim TempAttackPower As Long
    With UserList(UserIndex)
        TempAttackPower = ((.Stats.UserSkills(Skill) + ((3 * .Stats.UserSkills(Skill) / 100) * .Stats.UserAtributos(e_Atributos.Agilidad))) * skillModifier)
        AttackPower = (TempAttackPower + (2.5 * Maximo(CInt(.Stats.ELV) - 12, 0)))
        AttackPower = AttackPower + UserMod.GetHitBonus(UserList(UserIndex))
    End With
    Exit Function
AttackPower_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.AttackPower", Erl)
End Function

Private Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long
    On Error GoTo PoderAtaqueArma_Err
    PoderAtaqueArma = AttackPower(UserIndex, e_Skill.Armas, ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
    Exit Function
PoderAtaqueArma_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.PoderAtaqueArma", Erl)
End Function

Private Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long
    On Error GoTo PoderAtaqueProyectil_Err
    PoderAtaqueProyectil = AttackPower(UserIndex, e_Skill.Proyectiles, ModificadorPoderAtaqueProyectiles(UserList(UserIndex).clase))
    Exit Function
PoderAtaqueProyectil_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.PoderAtaqueProyectil", Erl)
End Function

Private Function AttackPowerDaggers(ByVal UserIndex As Integer) As Long
    On Error GoTo AttackPowerDaggers_Err:
    AttackPowerDaggers = AttackPower(UserIndex, e_Skill.Apuñalar, ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
    Exit Function
AttackPowerDaggers_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.AttackPowerDaggers", Erl)
End Function

Private Function PoderAtaqueWrestling(ByVal UserIndex As Integer) As Long
    On Error GoTo PoderAtaqueWrestling_Err
    PoderAtaqueWrestling = AttackPower(UserIndex, e_Skill.Wrestling, ModificadorPoderAtaqueArmas(UserList(UserIndex).clase))
    Exit Function
PoderAtaqueWrestling_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.PoderAtaqueWrestling", Erl)
End Function

Private Function UserImpactoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal aType As AttackType) As Boolean
    On Error GoTo UserImpactoNpc_Err
    Dim PoderAtaque As Long
    Dim Arma        As Integer
    Dim ProbExito   As Long
    Arma = UserList(UserIndex).invent.EquippedWeaponObjIndex
    Dim RequiredSkill As e_Skill
    RequiredSkill = GetSkillRequiredForWeapon(Arma)
    If RequiredSkill = Wrestling Then
        PoderAtaque = PoderAtaqueWrestling(UserIndex)
    ElseIf RequiredSkill = Armas Then
        PoderAtaque = PoderAtaqueArma(UserIndex)
    ElseIf RequiredSkill = Proyectiles Then
        PoderAtaque = PoderAtaqueProyectil(UserIndex)
    Else
        PoderAtaque = PoderAtaqueWrestling(UserIndex)
    End If
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - NpcList(NpcIndex).PoderEvasion) * 0.4)))
    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    If UserImpactoNpc Then
        Call SubirSkillDeArmaActual(UserIndex)
    End If
    If Not IsSet(NpcList(NpcIndex).flags.StatusMask, eTaunted) Then
        Call SetUserRef(NpcList(NpcIndex).TargetUser, UserIndex)
    End If
    Exit Function
UserImpactoNpc_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.UserImpactoNpc", Erl)
End Function

Private Function NpcImpacto(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Boolean
    On Error GoTo NpcImpacto_Err
    Dim Rechazo           As Boolean
    Dim ProbRechazo       As Long
    Dim ProbExito         As Long
    Dim UserEvasion       As Long
    Dim NpcPoderAtaque    As Long
    Dim PoderEvasioEscudo As Long
    Dim SkillTacticas     As Long
    Dim SkillDefensa      As Long
    UserEvasion = PoderEvasion(UserIndex)
    NpcPoderAtaque = NpcList(NpcIndex).PoderAtaque + NPCs.GetHitBonus(NpcList(NpcIndex))
    PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)
    SkillTacticas = UserList(UserIndex).Stats.UserSkills(e_Skill.Tacticas)
    SkillDefensa = UserList(UserIndex).Stats.UserSkills(e_Skill.Defensa)
    'Esta usando un escudo ???
    If UserList(UserIndex).invent.EquippedShieldObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
    ProbExito = Maximo(10, Minimo(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
    ' el usuario esta usando un escudo ???
    If UserList(UserIndex).invent.EquippedShieldObjIndex > 0 Then
        If ObjData(UserList(UserIndex).invent.EquippedShieldObjIndex).Porcentaje > 0 Then
            If Not NpcImpacto Then
                If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
                    ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
                    Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                    If Rechazo = True Then
                        'Se rechazo el ataque con el escudo
                        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                        If UserList(UserIndex).ChatCombate = 1 Then
                            Call Write_BlockedWithShieldUser(UserIndex)
                        End If
                    End If
                End If
            End If
            Call SubirSkill(UserIndex, Defensa)
        End If
    End If
    Exit Function
NpcImpacto_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.NpcImpacto", Erl)
End Function

Private Function GetUserDamage(ByVal UserIndex As Integer) As Long
    On Error GoTo GetUserDamge_Err
    With UserList(UserIndex)
        GetUserDamage = GetUserDamageWithItem(UserIndex, .invent.EquippedWeaponObjIndex, .invent.EquippedMunitionObjIndex) + UserMod.GetLinearDamageBonus(UserIndex)
    End With
    Exit Function
GetUserDamge_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.GetUserDamge", Erl)
End Function

Public Function GetClassAttackModifier(ByRef ObjData As t_ObjData, ByVal Class As e_Class) As Single
    If ObjData.Proyectil > 0 Then
        GetClassAttackModifier = ModicadorDañoClaseProyectiles(Class)
    ElseIf ObjData.WeaponType = eKnuckle Then
        GetClassAttackModifier = ModClase(Class).DañoWrestling
    Else
        GetClassAttackModifier = ModicadorDañoClaseArmas(Class)
    End If
End Function

Public Function GetUserDamageWithItem(ByVal UserIndex As Integer, ByVal WeaponObjIndex As Integer, ByVal AmunitionObjIndex As Integer) As Long
    On Error GoTo GetUserDamageWithItem_Err
    Dim UserDamage As Long, WeaponDamage As Long, MaxWeaponDamage As Long, ClassModifier As Single
    With UserList(UserIndex)
        ' Daño base del usuario
        UserDamage = RandomNumber(.Stats.MinHIT, .Stats.MaxHit)
        ' Daño con arma
        If WeaponObjIndex > 0 Then
            Dim Arma As t_ObjData
            Arma = ObjData(WeaponObjIndex)
            ClassModifier = GetClassAttackModifier(Arma, .clase)
            ' Calculamos el daño del arma
            WeaponDamage = RandomNumber(Arma.MinHIT, Arma.MaxHit)
            ' Daño máximo del arma
            MaxWeaponDamage = Arma.MaxHit
            ' Si lanza proyectiles
            If Arma.Proyectil > 0 Then
                ' Si requiere munición
                If Arma.Municion > 0 And AmunitionObjIndex > 0 Then
                    Dim Municion As t_ObjData
                    Municion = ObjData(AmunitionObjIndex)
                    ' Agregamos el daño de la munición al daño del arma
                    WeaponDamage = WeaponDamage + RandomNumber(Municion.MinHIT, Municion.MaxHit)
                    MaxWeaponDamage = Arma.MaxHit + Municion.MaxHit
                End If
            End If
            ' Daño con puños
        Else
            ' Modificador de combate sin armas
            ClassModifier = ModClase(.clase).DañoWrestling
        End If
        ' Base damage
        GetUserDamageWithItem = (3 * WeaponDamage + MaxWeaponDamage * 0.2 * Maximo(0, .Stats.UserAtributos(Fuerza) - 15) + UserDamage) * ClassModifier
        ' Ship bonus
        If .flags.Navegando = 1 And .invent.EquippedShipObjIndex > 0 Then
            GetUserDamageWithItem = GetUserDamageWithItem + RandomNumber(ObjData(.invent.EquippedShipObjIndex).MinHIT, ObjData(.invent.EquippedShipObjIndex).MaxHit)
            ' mount bonus
        ElseIf .flags.Montado = 1 And .invent.EquippedSaddleObjIndex > 0 Then
            GetUserDamageWithItem = GetUserDamageWithItem + RandomNumber(ObjData(.invent.EquippedSaddleObjIndex).MinHIT, ObjData(.invent.EquippedSaddleObjIndex).MaxHit)
        End If
    End With
    Exit Function
GetUserDamageWithItem_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.GetUserDamageWithItem", Erl)
End Function

Private Sub UserDamageNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal aType As AttackType)
    On Error GoTo UserDamageNpc_Err
    With UserList(UserIndex)
        Dim Damage As Long, DamageBase As Long, DamageExtra As Long, Color As Long, DamageStr As String
        If .invent.EquippedWeaponObjIndex = EspadaMataDragonesIndex And NpcList(NpcIndex).npcType = DRAGON Then
            ' Espada MataDragones
            DamageBase = NpcList(NpcIndex).Stats.MinHp + NpcList(NpcIndex).Stats.def
            ' La pierde una vez usada
            Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)
            'registramos quien mato y uso la MD
            Call LogGM(.name, " Mato un Dragon Rojo ")
        Else
            ' Daño normal o elemental
            DamageBase = GetUserDamage(UserIndex)
            ' NPC de pruebas
            If NpcList(NpcIndex).npcType = DummyTarget Then
                Call DummyTargetAttacked(NpcIndex)
            End If
        End If
        ' Color por defecto rojo
        Color = vbRed
        Dim NpcDef As Integer
        NpcDef = NpcList(NpcIndex).Stats.def + NPCs.GetDefenseBonus(NpcIndex)
        NpcDef = max(0, NpcDef - GetArmorPenetration(UserIndex, NpcDef))
        ' Defensa del NPC
        Damage = DamageBase - NpcDef
        Damage = Damage * UserMod.GetPhysicalDamageModifier(UserList(UserIndex))
        Damage = Damage * NPCs.GetPhysicDamageReduction(NpcList(NpcIndex))
        If IsFeatureEnabled("elemental_tags") Then
            Call CalculateElementalTagsModifiers(UserIndex, NpcIndex, Damage)
        End If
        If Damage < 0 Then Damage = 0
        If IsFeatureEnabled("healers_and_tanks") And .clase = e_Class.Warrior Then
            Dim Calc As Integer
            Calc = Damage * WarriorLifeStealOnHitMultiplier
            .Stats.MinHp = .Stats.MinHp + Calc
            If .Stats.MinHp > .Stats.MaxHp Then
                .Stats.MinHp = .Stats.MaxHp
            End If
            Call WriteUpdateHP(UserIndex)
            'no wrapper senddata because of extra params
            Call modSendData.SendData(ToIndex, UserIndex, PrepareMessageTextOverTile(Calc, .pos.x, .pos.y, vbGreen, 1300, -10, True))
        End If
        ' Golpe crítico
        If PuedeGolpeCritico(UserIndex) Then
            ' Si acertó - Doble chance contra NPCs
            If RandomNumber(1, 100) <= GetCriticalHitChanceBase(UserIndex) Then
                ' Daño del golpe crítico (usamos el daño base)
                DamageExtra = DamageBase * 0.33
                DamageExtra = DamageExtra * UserMod.GetPhysicalDamageModifier(UserList(UserIndex))
                DamageExtra = DamageExtra * NPCs.GetPhysicDamageReduction(NpcList(NpcIndex))
                ' Mostramos en consola el daño
                If .ChatCombate = 1 Then
                    Call WriteLocaleMsg(UserIndex, 383, e_FontTypeNames.FONTTYPE_INFOBOLD, PonerPuntos(Damage) & "¬" & (DamageExtra))
                End If
                ' Color naranja
                Color = RGB(225, 165, 0)
            End If
            ' Stab
        ElseIf PuedeApuñalar(UserIndex) Then
            ' Si acertó - Doble chance contra NPCs
            If RandomNumber(1, 100) <= GetStabbingChanceBase(UserIndex) Then
                Dim min_stab_npc As Double
                Dim max_stab_npc As Double
                min_stab_npc = GetStabbingNPCMinForClass(UserList(UserIndex).clase)
                max_stab_npc = GetStabbingNPCMaxForClass(UserList(UserIndex).clase)
                ' Daño del apunalamiento (formula con valor oscilante en contra de NPCs)
                DamageExtra = Damage * (Rnd * (max_stab_npc - min_stab_npc) + min_stab_npc)
                ' Mostramos en consola el daño
                If .ChatCombate = 1 Then
                    Call WriteLocaleMsg(UserIndex, 212, e_FontTypeNames.FONTTYPE_INFOBOLD, PonerPuntos(Damage) & "¬" & PonerPuntos(DamageExtra))
                End If
                ' Color amarillo
                Color = vbYellow
            End If
            ' Sube skills en apuñalar
            Call SubirSkill(UserIndex, Apuñalar)
        End If
        If DamageExtra > 0 Then
            Damage = Damage + DamageExtra
        End If
        ' Restamos el daño al NPC
        If NPCs.DoDamageOrHeal(NpcIndex, UserIndex, eUser, -Damage, e_phisical, .invent.EquippedWeaponObjIndex, Color) = eStillAlive Then
            'efectos
            Dim ArmaObjInd, ObjInd As Integer
            ObjInd = 0
            ArmaObjInd = .invent.EquippedWeaponObjIndex
            If ArmaObjInd > 0 Then
                If ObjData(ArmaObjInd).Municion = 0 Then
                    ObjInd = ArmaObjInd
                Else
                    ObjInd = .invent.EquippedMunitionObjIndex
                End If
            End If
            Dim rangeStun  As Boolean
            Dim stunChance As Byte
            If ObjInd > 0 Then
                rangeStun = ObjData(ObjInd).Subtipo = 2 And aType = Ranged
                stunChance = ObjData(ObjInd).Porcentaje
            End If
            If rangeStun And NpcList(NpcIndex).flags.AfectaParalisis = 0 Then
                If (RandomNumber(1, 100) < stunChance) Then
                    With NpcList(NpcIndex)
                        Call StunNPc(.Contadores)
                        Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageCreateFX(.Char.charindex, 142, 6))
                    End With
                End If
            End If
        End If
    End With
    Exit Sub
UserDamageNpc_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.UserDañoNpc", Erl)
End Sub

Public Function UserDamageToNpc(ByVal attackerIndex As Integer, _
                                ByVal TargetIndex As Integer, _
                                ByVal Damage As Long, _
                                ByVal Source As e_DamageSourceType, _
                                ByVal ObjIndex As Integer) As e_DamageResult
    Damage = Damage * UserMod.GetPhysicalDamageModifier(UserList(attackerIndex))
    Damage = Damage * NPCs.GetPhysicDamageReduction(NpcList(TargetIndex))
    UserDamageToNpc = NPCs.DoDamageOrHeal(TargetIndex, attackerIndex, e_ReferenceType.eUser, -Damage, Source, ObjIndex)
End Function

Public Function GetNpcDamage(ByVal NpcIndex As Integer) As Long
    GetNpcDamage = RandomNumber(NpcList(NpcIndex).Stats.MinHIT, NpcList(NpcIndex).Stats.MaxHit) + NPCs.GetLinearDamageBonus(NpcIndex)
End Function

Private Function NpcDamage(ByVal NpcIndex As Integer, ByVal UserIndex As Integer) As Long
    On Error GoTo NpcDamage_Err
    NpcDamage = -1
    Dim Damage   As Integer, Lugar As Integer, absorbido As Integer
    Dim defbarco As Integer
    Dim obj      As t_ObjData
    Damage = GetNpcDamage(NpcIndex)
    If UserList(UserIndex).flags.Navegando = 1 And UserList(UserIndex).invent.EquippedShipObjIndex > 0 Then
        obj = ObjData(UserList(UserIndex).invent.EquippedShipObjIndex)
        defbarco = RandomNumber(obj.MinDef, obj.MaxDef)
    End If
    Dim defMontura As Integer
    If UserList(UserIndex).flags.Montado = 1 And UserList(UserIndex).invent.EquippedSaddleObjIndex > 0 Then
        obj = ObjData(UserList(UserIndex).invent.EquippedSaddleObjIndex)
        defMontura = RandomNumber(obj.MinDef, obj.MaxDef)
    End If
    Lugar = RandomNumber(1, 6)
    Select Case Lugar
            ' 1/6 de chances de que sea a la cabeza
        Case e_PartesCuerpo.bCabeza
            'Si tiene casco absorbe el golpe
            If UserList(UserIndex).invent.EquippedHelmetObjIndex > 0 Then
                Dim Casco As t_ObjData
                Casco = ObjData(UserList(UserIndex).invent.EquippedHelmetObjIndex)
                absorbido = absorbido + RandomNumber(Casco.MinDef, Casco.MaxDef)
            End If
        Case Else
            'Si tiene armadura absorbe el golpe
            If UserList(UserIndex).invent.EquippedArmorObjIndex > 0 Then
                Dim Armadura As t_ObjData
                Armadura = ObjData(UserList(UserIndex).invent.EquippedArmorObjIndex)
                absorbido = absorbido + RandomNumber(Armadura.MinDef, Armadura.MaxDef)
            End If
            'Si tiene escudo absorbe el golpe
            If UserList(UserIndex).invent.EquippedShieldObjIndex > 0 Then
                Dim Escudo As t_ObjData
                Escudo = ObjData(UserList(UserIndex).invent.EquippedShieldObjIndex)
                absorbido = absorbido + RandomNumber(Escudo.MinDef, Escudo.MaxDef)
            End If
    End Select
    Damage = Damage - absorbido - defbarco - defMontura - UserMod.GetDefenseBonus(UserIndex)
    Damage = Damage * NPCs.GetPhysicalDamageModifier(NpcList(NpcIndex))
    Damage = Damage * UserMod.GetPhysicDamageReduction(UserList(UserIndex))
    If Damage < 0 Then Damage = 0
    If UserList(UserIndex).ChatCombate = 1 Then
        Call WriteNPCHitUser(UserIndex, Lugar, Damage)
    End If
    If UserList(UserIndex).flags.Privilegios And e_PlayerType.User Then Call UserMod.DoDamageOrHeal(UserIndex, NpcIndex, eNpc, -Damage, e_phisical, 0)
    If UserList(UserIndex).flags.Meditando Then
        If Damage > Fix(UserList(UserIndex).Stats.MinHp / 100 * UserList(UserIndex).Stats.UserAtributos(e_Atributos.Inteligencia) * UserList(UserIndex).Stats.UserSkills( _
                e_Skill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
            UserList(UserIndex).flags.Meditando = False
            UserList(UserIndex).Char.FX = 0
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.charindex, 0))
        End If
    End If
    NpcDamage = Damage
    Exit Function
NpcDamage_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.NpcDamage", Erl)
End Function

Public Function NpcDoDamageToUser(ByVal attackerIndex As Integer, _
                                  ByVal TargetIndex As Integer, _
                                  ByVal Damage As Long, _
                                  ByVal Source As e_DamageSourceType, _
                                  ByVal ObjIndex As Integer) As e_DamageResult
    Damage = Damage * NPCs.GetPhysicalDamageModifier(NpcList(attackerIndex))
    Damage = Damage * UserMod.GetPhysicDamageReduction(UserList(TargetIndex))
    NpcDoDamageToUser = UserMod.DoDamageOrHeal(TargetIndex, attackerIndex, e_ReferenceType.eNpc, -Damage, Source, ObjIndex)
    If UserList(TargetIndex).ChatCombate = 1 Then
        Call WriteNPCHitUser(TargetIndex, bTorso, Damage)
    End If
End Function

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Heading As e_Heading) As Boolean
    On Error GoTo NpcAtacaUser_Err
    If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
    If UserList(UserIndex).flags.Muerto = 1 Then Exit Function
    If (Not UserList(UserIndex).flags.Privilegios And e_PlayerType.User) <> 0 And Not UserList(UserIndex).flags.AdminPerseguible Then Exit Function
    ' El npc puede atacar ???
    If Not IntervaloPermiteAtacarNPC(NpcIndex) Then
        NpcAtacaUser = False
        Exit Function
    End If
    If ((MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).Blocked And 2 ^ (Heading - 1)) <> 0) Then
        NpcAtacaUser = False
        Exit Function
    End If
    NpcAtacaUser = True
    Call AllMascotasAtacanNPC(NpcIndex, UserIndex)
    Call ResetUserAutomatedActions(UserIndex)
    UserList(UserIndex).Counters.EnCombate = IntervaloEnCombate
    If Not IsValidUserRef(NpcList(NpcIndex).TargetUser) Then
        Call SetUserRef(NpcList(NpcIndex).TargetUser, UserIndex)
    End If
    If Not IsValidNpcRef(UserList(UserIndex).flags.AtacadoPorNpc) And UserList(UserIndex).flags.AtacadoPorUser = 0 Then Call SetNpcRef(UserList(UserIndex).flags.AtacadoPorNpc, _
            NpcIndex)
    If NpcList(NpcIndex).flags.Snd1 > 0 Then
        Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessagePlayWave(NpcList(NpcIndex).flags.Snd1, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y))
    End If
    Call CancelExit(UserIndex)
    If NpcList(NpcIndex).flags.Inmovilizado = 0 And NpcList(NpcIndex).flags.AttackedBy <> UserList(UserIndex).name Then
        NpcList(NpcIndex).flags.AttackedBy = vbNullString
    End If
    Dim danio As Long
    danio = -1
    If NpcImpacto(NpcIndex, UserIndex) Then
        danio = NpcDamage(NpcIndex, UserIndex)
        '¿Puede envenenar?
        If NpcList(NpcIndex).Veneno > 0 Then Call NpcEnvenenarUser(UserIndex, NpcList(NpcIndex).Veneno)
    End If
    Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageCharAtaca(NpcList(NpcIndex).Char.charindex, UserList(UserIndex).Char.charindex, danio, NpcList( _
            NpcIndex).Char.Ataque1))
    If NpcList(NpcIndex).Char.WeaponAnim > 0 Then
        Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageArmaMov(NpcList(NpcIndex).Char.charindex, 0))
    End If
    '-----Tal vez suba los skills------
    Call SubirSkill(UserIndex, Tacticas)
    'Controla el nivel del usuario
    Call CheckUserLevel(UserIndex)
    Exit Function
NpcAtacaUser_Err:
    Call TraceError(Err.Number, Err.Description & " Linea---> " & Erl, "SistemaCombate.NpcAtacaUser", Erl)
End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
    On Error GoTo NpcImpactoNpc_Err
    Dim PoderAtt  As Long, PoderEva As Long
    Dim ProbExito As Long
    PoderAtt = NpcList(Atacante).PoderAtaque
    PoderEva = NpcList(Victima).PoderEvasion
    ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtt - PoderEva) * 0.4)))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    Exit Function
NpcImpactoNpc_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.NpcImpactoNpc", Erl)
End Function

Public Function NpcDamageNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Long
    Dim Damage As Long

    With NpcList(Atacante)
        Damage = RandomNumber(.Stats.MinHIT, .Stats.MaxHit) _
                 + NPCs.GetLinearDamageBonus(Atacante) _
                 - NPCs.GetDefenseBonus(Victima) _
                 - NpcList(Victima).Stats.def
    End With

    ' Evitamos valores negativos
    If Damage < 0 Then Damage = 0

    ' Aplicamos el daño real en el juego (usa la lógica existente)
    Call NpcDamageToNpc(Atacante, Victima, CInt(Damage))

    ' Devolvemos el daño para que el caller lo mande al cliente
    NpcDamageNpc = Damage
End Function

Public Function NpcDamageToNpc(ByVal attackerIndex As Integer, _
                               ByVal TargetIndex As Integer, _
                               ByVal Damage As Integer) As e_DamageResult
    On Error GoTo NpcDamageNpc_Err

    With NpcList(attackerIndex)
        ' Ojo: aquí se recalcula el Damage interno con modificadores
        Dim finalDamage As Long

        finalDamage = Damage
        finalDamage = finalDamage * NPCs.GetPhysicalDamageModifier(NpcList(attackerIndex))
        finalDamage = finalDamage * NPCs.GetPhysicDamageReduction(NpcList(TargetIndex))

        NpcDamageToNpc = NPCs.DoDamageOrHeal(TargetIndex, attackerIndex, eNpc, -finalDamage, e_phisical, 0)

        If NpcDamageToNpc = eDead Then
            If Not IsValidUserRef(NpcList(attackerIndex).MaestroUser) Then
                Call SetMovement(attackerIndex, .flags.OldMovement)
                If LenB(.flags.AttackedBy) <> 0 Then
                    .Hostile = .flags.OldHostil
                End If
            End If
        End If
    End With
    Exit Function
NpcDamageNpc_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.NpcDamageNpc")
End Function


Public Function NpcPerformAttackNpc(ByVal attackerIndex As Integer, ByVal TargetIndex As Integer) As Boolean
    Dim danio As Long
    Dim impacto As Boolean

    danio = -1 ' -1 = miss by default

    If NpcList(attackerIndex).flags.Snd1 > 0 Then
        Call SendData( _
            SendTarget.ToNPCAliveArea, _
            attackerIndex, _
            PrepareMessagePlayWave(NpcList(attackerIndex).flags.Snd1, NpcList(attackerIndex).pos.x, NpcList(attackerIndex).pos.y))
    End If

    If NpcList(attackerIndex).Char.WeaponAnim > 0 Then
        Call SendData(SendTarget.ToNPCAliveArea, attackerIndex, PrepareMessageArmaMov(NpcList(attackerIndex).Char.charindex, 0))
    End If

    impacto = NpcImpactoNpc(attackerIndex, TargetIndex)

    If impacto Then
        If NpcList(attackerIndex).flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, TargetIndex, PrepareMessagePlayWave(NpcList(attackerIndex).flags.Snd1, NpcList(TargetIndex).pos.x, NpcList(TargetIndex).pos.y))
        Else
            Call SendData(SendTarget.ToNPCAliveArea, TargetIndex, PrepareMessagePlayWave(SND_IMPACTO2, NpcList(TargetIndex).pos.x, NpcList(TargetIndex).pos.y))
        End If
        danio = NpcDamageNpc(attackerIndex, TargetIndex)
    Else
        Call SendData(SendTarget.ToNPCAliveArea, attackerIndex, PrepareMessageCharSwing(NpcList(attackerIndex).Char.charindex, False, True))
    End If

    Call SendData( _
        SendTarget.ToNPCAliveArea, _
        attackerIndex, _
        PrepareMessageCharAtaca(NpcList(attackerIndex).Char.charindex, NpcList(TargetIndex).Char.charindex, danio, NpcList(attackerIndex).Char.Ataque1))

    NpcPerformAttackNpc = impacto
End Function

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMovimiento As Boolean = True)
    On Error GoTo NpcAtacaNpc_Err
    If Not IntervaloPermiteAtacarNPC(Atacante) Then Exit Sub
    Dim Heading As e_Heading
    ' Determina hacia dónde debe mirar el atacante
    Heading = GetHeadingFromWorldPos(NpcList(Atacante).pos, NpcList(Victima).pos)
    ' Si no está mirando y está paralizado, no puede girar ni atacar
    If Heading <> NpcList(Atacante).Char.Heading Then
        If NpcList(Atacante).flags.Paralizado = 1 Then
            Call ClearNpcRef(NpcList(Atacante).TargetNPC)
            Call SetMovement(Atacante, e_TipoAI.MueveAlAzar)
            Exit Sub
        End If
    End If
    ' Si puede girar, lo hace
    Call ChangeNPCChar(Atacante, NpcList(Atacante).Char.body, NpcList(Atacante).Char.head, Heading)
    ' La víctima podría reaccionar
    Heading = GetHeadingFromWorldPos(NpcList(Victima).pos, NpcList(Atacante).pos)
    If Heading <> NpcList(Victima).Char.Heading Then
        If NpcList(Victima).flags.Paralizado = 1 Then
            cambiarMovimiento = False ' Si está paralizado, no puede reaccionar
        End If
    End If
    If cambiarMovimiento Then
        Call SetNpcRef(NpcList(Victima).TargetNPC, Atacante)
        Call SetMovement(Victima, e_TipoAI.NpcAtacaNpc)
    End If
    If NpcList(Atacante).flags.Snd1 > 0 Then
        Call SendData(SendTarget.ToNPCAliveArea, Atacante, PrepareMessagePlayWave(NpcList(Atacante).flags.Snd1, NpcList(Atacante).pos.x, NpcList(Atacante).pos.y))
    End If
    ' Ejecuta el ataque real
    Call NpcPerformAttackNpc(Atacante, Victima)
    Exit Sub
NpcAtacaNpc_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.NpcAtacaNpc", Erl)
End Sub

Public Sub UsuarioAtacaNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal aType As AttackType)
    On Error GoTo UsuarioAtacaNpc_Err
    'Si el npc es solo atacable para clanes y el usuario no tiene clan, le avisa y sale de la funcion
    If NpcList(NpcIndex).OnlyForGuilds = 1 And UserList(UserIndex).GuildIndex <= 0 Then
        'Msg2001=Debes pertenecer a un clan para atacar a este NPC
        Call WriteLocaleMsg(UserIndex, 2001, e_FontTypeNames.FONTTYPE_WARNING)
        Exit Sub
    End If
    Dim UserAttackInteractionResult As t_AttackInteractionResult
    UserAttackInteractionResult = UserCanAttackNpc(UserIndex, NpcIndex)
    Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResult.Result)
    If UserAttackInteractionResult.CanAttack Then
        If UserAttackInteractionResult.TurnPK Then Call VolverCriminal(UserIndex)
    Else
        Exit Sub
    End If
    Call AllMascotasAtacanNPC(NpcIndex, UserIndex)
    If UserList(UserIndex).flags.invisible = 0 Then Call NPCAtacado(NpcIndex, UserIndex)
    Call EffectsOverTime.TargetWillAttack(UserList(UserIndex).EffectOverTime, NpcIndex, eNpc, e_phisical)
    If UserImpactoNpc(UserIndex, NpcIndex, aType) Then
        ' Suena el Golpe en el cliente.
        If NpcList(NpcIndex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessagePlayWave(NpcList(NpcIndex).flags.Snd2, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y))
        Else
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_IMPACTO2, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y))
        End If
        ' Golpe Paralizador
        If UserList(UserIndex).flags.Paraliza = 1 And NpcList(NpcIndex).flags.Paralizado = 0 Then
            If RandomNumber(1, 4) = 1 Then
                If NpcList(NpcIndex).flags.AfectaParalisis = 0 Then
                    NpcList(NpcIndex).flags.Paralizado = 1
                    NpcList(NpcIndex).Contadores.Paralisis = (IntervaloParalizado / 3) * 7
                    If UserList(UserIndex).ChatCombate = 1 Then
                        Call WriteLocaleMsg(UserIndex, 136, e_FontTypeNames.FONTTYPE_FIGHT)
                    End If
                    UserList(UserIndex).Counters.timeFx = 3
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(NpcList(NpcIndex).Char.charindex, 8, 0, UserList(UserIndex).pos.x, UserList( _
                            UserIndex).pos.y))
                Else
                    If UserList(UserIndex).ChatCombate = 1 Then
                        Call WriteLocaleMsg(UserIndex, 381, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End If
        ' Cambiamos el objetivo del NPC si uno le pega cuerpo a cuerpo.
        If Not IsSet(NpcList(NpcIndex).flags.StatusMask, eTaunted) And (Not IsValidUserRef(NpcList(NpcIndex).TargetUser) Or NpcList(NpcIndex).TargetUser.ArrayIndex <> UserIndex) _
                Then
            Call SetUserRef(NpcList(NpcIndex).TargetUser, UserIndex)
        End If
        ' Si te mimetizaste en forma de bicho y le pegas al chobi, el chobi te va a pegar.
        If UserList(UserIndex).flags.Mimetizado = e_EstadoMimetismo.FormaBicho Then
            UserList(UserIndex).flags.Mimetizado = e_EstadoMimetismo.FormaBichoSinProteccion
        End If
        ' Resta la vida del NPC
        Call UserDamageNpc(UserIndex, NpcIndex, aType)
        Dim Arma          As Integer: Arma = UserList(UserIndex).invent.EquippedWeaponObjIndex
        Dim municionIndex As Integer: municionIndex = UserList(UserIndex).invent.EquippedMunitionObjIndex
        Dim Particula     As Integer
        Dim Tiempo        As Long
        If Arma > 0 Then
            If municionIndex > 0 And ObjData(Arma).Proyectil Then
                If ObjData(municionIndex).CreaFX <> 0 Then
                    UserList(UserIndex).Counters.timeFx = 3
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(NpcList(NpcIndex).Char.charindex, ObjData(municionIndex).CreaFX, 0, UserList( _
                            UserIndex).pos.x, UserList(UserIndex).pos.y))
                End If
                If ObjData(municionIndex).CreaParticula <> "" Then
                    Particula = val(ReadField(1, ObjData(municionIndex).CreaParticula, Asc(":")))
                    Tiempo = val(ReadField(2, ObjData(municionIndex).CreaParticula, Asc(":")))
                    UserList(UserIndex).Counters.timeFx = 3
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(NpcList(NpcIndex).Char.charindex, Particula, Tiempo, False, , UserList( _
                            UserIndex).pos.x, UserList(UserIndex).pos.y))
                End If
            End If
        End If
        Call EffectsOverTime.TargetDidHit(UserList(UserIndex).EffectOverTime, NpcIndex, eNpc, e_phisical)
    Else
        Call EffectsOverTime.TargetFailedAttack(UserList(UserIndex).EffectOverTime, NpcIndex, eNpc, e_phisical)
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.charindex, , , IIf(UserList(UserIndex).flags.invisible + UserList( _
                UserIndex).flags.Oculto > 0, False, True)))
    End If
    Exit Sub
UsuarioAtacaNpc_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtacaNpc", Erl)
End Sub

Public Sub UserAttackPosition(ByVal UserIndex As Integer, ByRef TargetPos As t_WorldPos, Optional ByVal IsExtraHit As Boolean = False)
    'Exit if not legal
    If TargetPos.x >= XMinMapSize And TargetPos.x <= XMaxMapSize And TargetPos.y >= YMinMapSize And TargetPos.y <= YMaxMapSize Then
        If ((MapData(TargetPos.Map, TargetPos.x, TargetPos.y).Blocked And 2 ^ (UserList(UserIndex).Char.Heading - 1)) <> 0) Then
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.charindex, True, False))
            Exit Sub
        End If
        Dim Index As Integer
        Index = MapData(TargetPos.Map, TargetPos.x, TargetPos.y).UserIndex
        'Look for user
        If Index > 0 Then
            Call UsuarioAtacaUsuario(UserIndex, Index, Melee)
            'Look for NPC
        ElseIf MapData(TargetPos.Map, TargetPos.x, TargetPos.y).NpcIndex > 0 Then
            Index = MapData(TargetPos.Map, TargetPos.x, TargetPos.y).NpcIndex
            If NpcList(Index).Attackable Then
                If IsValidUserRef(NpcList(Index).MaestroUser) And MapInfo(NpcList(Index).pos.Map).Seguro = 1 Then
                    'Msg1041= No podés atacar mascotas en zonas seguras
                    Call WriteLocaleMsg(UserIndex, 1041, e_FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                Call UsuarioAtacaNpc(UserIndex, Index, Melee)
            Else
                'Msg1042= No podés atacar a este NPC
                Call WriteLocaleMsg(UserIndex, 1042, e_FontTypeNames.FONTTYPE_FIGHT)
            End If
            Exit Sub
        Else
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.charindex, True, False))
            With UserList(UserIndex)
                If Not IsExtraHit And .flags.Inmovilizado + .flags.Paralizado > 0 Then
                    .Counters.Inmovilizado = max(0, .Counters.Inmovilizado - AirHitReductParalisisTime)
                    .Counters.Paralisis = max(0, .Counters.Paralisis - AirHitReductParalisisTime)
                End If
            End With
        End If
    Else
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCharSwing(UserList(UserIndex).Char.charindex, True, False))
    End If
End Sub

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)
    On Error GoTo UsuarioAtaca_Err
    'Check bow's interval
    If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
    'Check Spell-Attack interval
    If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Sub
    'Check Attack interval
    If Not IntervaloPermiteAtacar(UserIndex) Then Exit Sub
    With UserList(UserIndex)
        'Quitamos stamina
        If .Stats.MinSta < 10 Then
            'Msg93=Estás muy cansado
            Call WriteLocaleMsg(UserIndex, 93, e_FontTypeNames.FONTTYPE_INFO)
            'Msg2129=¡No tengo energía!
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareLocalizedChatOverHead(2129, UserList(UserIndex).Char.charindex, vbWhite))
            Exit Sub
        End If
        Call QuitarSta(UserIndex, RandomNumber(1, 10))
        If .Counters.Trabajando Then
            Call WriteMacroTrabajoToggle(UserIndex, False)
        End If
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
        'Movimiento de arma, solo lo envio si no es GM invisible.
        If .flags.AdminInvisible = 0 Then
            If IsSet(.flags.StatusMask, e_StatusMask.eTransformed) Then
                If .Char.Ataque1 > 0 Then
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageDoAnimation(.Char.charindex, .Char.Ataque1))
                End If
            Else
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageArmaMov(.Char.charindex))
            End If
        End If
        Dim AttackPos As t_WorldPos
        AttackPos = UserList(UserIndex).pos
        Call HeadtoPos(.Char.Heading, AttackPos)
        Call EffectsOverTime.TargetWillAttackPosition(UserList(UserIndex).EffectOverTime, AttackPos)
        If .flags.Cleave > 0 Then
            Call IncreaseSingle(.Modifiers.PhysicalDamageBonus, -0.25) 'Front target gets 75% damage
            Call UserAttackPosition(UserIndex, AttackPos)
            Call IncreaseSingle(.Modifiers.PhysicalDamageBonus, -0.25) 'Side targets gets 50% damage
            AttackPos = UserList(UserIndex).pos
            Call GetHeadingRight(.Char.Heading, AttackPos)
            Call UserAttackPosition(UserIndex, AttackPos, True)
            AttackPos = UserList(UserIndex).pos
            Call GetHeadingLeft(.Char.Heading, AttackPos)
            Call UserAttackPosition(UserIndex, AttackPos, True)
            Call IncreaseSingle(.Modifiers.PhysicalDamageBonus, 0.5) 'return to prev state
        Else
            Call UserAttackPosition(UserIndex, AttackPos)
        End If
    End With
    Exit Sub
UsuarioAtaca_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtaca", Erl)
End Sub

Private Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer, ByVal aType As AttackType) As Boolean
    On Error GoTo UsuarioImpacto_Err
    Dim ProbRechazo            As Long
    Dim Rechazo                As Boolean
    Dim ProbExito              As Long
    Dim PoderAtaque            As Long
    Dim UserPoderEvasion       As Long
    Dim Arma                   As Integer
    Dim SkillTacticas          As Long
    Dim SkillDefensa           As Long
    Dim ProbEvadir             As Long
    Dim ShieldChancePercentage As Long
    If UserList(AtacanteIndex).flags.GolpeCertero = 1 Then
        UsuarioImpacto = True
        UserList(AtacanteIndex).flags.GolpeCertero = 0
        Exit Function
    End If
    SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(e_Skill.Tacticas)
    SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(e_Skill.Defensa)
    Arma = UserList(AtacanteIndex).invent.EquippedWeaponObjIndex
    Dim RequiredSkill As e_Skill
    RequiredSkill = GetSkillRequiredForWeapon(Arma)
    Select Case RequiredSkill
        Case e_Skill.Wrestling
            PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
        Case e_Skill.Armas
            PoderAtaque = PoderAtaqueArma(AtacanteIndex)
        Case e_Skill.Proyectiles
            PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
        Case e_Skill.Apuñalar
            PoderAtaque = AttackPowerDaggers(AtacanteIndex)
        Case Else
            PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
    End Select
    'Calculamos el poder de evasion...
    UserPoderEvasion = PoderEvasion(VictimaIndex)
    If UserList(VictimaIndex).invent.EquippedShieldObjIndex > 0 Then
        ShieldChancePercentage = ObjData(UserList(VictimaIndex).invent.EquippedShieldObjIndex).Porcentaje
        If ShieldChancePercentage > 0 Then
            UserPoderEvasion = UserPoderEvasion + PoderEvasionEscudo(VictimaIndex)
            If SkillDefensa > 0 Then
                ProbRechazo = Maximo(10, Minimo(90, (ShieldChancePercentage * (SkillDefensa / (Maximo(SkillDefensa + SkillTacticas, 1))))))
            Else
                ProbRechazo = 10
            End If
        Else
            ProbRechazo = 0
        End If
    Else
        ProbRechazo = 0
    End If
    Dim WeaponHitModifier As Integer
    WeaponHitModifier = 0
    If UserList(AtacanteIndex).invent.EquippedWeaponObjIndex > 0 And IsFeatureEnabled("Improved-Hit-Chance") Then
        If aType = Melee Then
            WeaponHitModifier = ObjData(UserList(AtacanteIndex).invent.EquippedWeaponObjIndex).ImprovedMeleeHitChance
        Else
            WeaponHitModifier = ObjData(UserList(AtacanteIndex).invent.EquippedWeaponObjIndex).ImprovedRangedHitChance
        End If
    End If
    ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - UserPoderEvasion) * 0.4) + WeaponHitModifier))
    ' Se reduce la evasion un 25%
    If UserList(VictimaIndex).flags.Meditando Then
        ProbEvadir = (100 - ProbExito) * 0.75
        ProbExito = MinimoInt(90, 100 - ProbEvadir)
    End If
    UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
    If UsuarioImpacto Then
        Call SubirSkillDeArmaActual(AtacanteIndex)
    Else ' Falló
        If RandomNumber(1, 100) <= ProbRechazo Then
            'Se rechazo el ataque con el escudo
            Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(VictimaIndex).pos.x, UserList(VictimaIndex).pos.y))
            Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageEscudoMov(UserList(VictimaIndex).Char.charindex))
            If UserList(AtacanteIndex).ChatCombate = 1 Then
                Call Write_BlockedWithShieldOther(AtacanteIndex)
            End If
            If UserList(VictimaIndex).ChatCombate = 1 Then
                Call Write_BlockedWithShieldUser(VictimaIndex)
            End If
            UserList(VictimaIndex).Counters.timeFx = 3
            Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, 88, 0, UserList(VictimaIndex).pos.x, UserList( _
                    VictimaIndex).pos.y))
            Call SubirSkill(VictimaIndex, e_Skill.Defensa)
        Else
            Call WriteConsoleMsg(VictimaIndex, PrepareMessageLocaleMsg(1930, UserList(AtacanteIndex).name, e_FontTypeNames.FONTTYPE_FIGHT)) ' Msg1930=¡¬1 te atacó y falló!
            'Msg1043= ¡Has fallado el golpe!
            Call WriteLocaleMsg(AtacanteIndex, "1043", e_FontTypeNames.FONTTYPE_FIGHT)
        End If
    End If
    Exit Function
UsuarioImpacto_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.UsuarioImpacto", Erl)
End Function

Public Sub UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer, ByVal aType As AttackType)
    On Error GoTo UsuarioAtacaUsuario_Err
    Dim sendto As SendTarget
    If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Sub
    If Distancia(UserList(AtacanteIndex).pos, UserList(VictimaIndex).pos) > MAXDISTANCIAARCO Then
        Call WriteLocaleMsg(AtacanteIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
    Call EffectsOverTime.TargetWillAttack(UserList(AtacanteIndex).EffectOverTime, VictimaIndex, eUser, e_phisical)
    Call ResetUserAutomatedActions(VictimaIndex)
    If UsuarioImpacto(AtacanteIndex, VictimaIndex, aType) Then
        If UserList(VictimaIndex).flags.Navegando = 0 Or UserList(VictimaIndex).flags.Montado = 0 Then
            UserList(VictimaIndex).Counters.timeFx = 3
            Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, FXSANGRE, 0, UserList(VictimaIndex).pos.x, _
                    UserList(VictimaIndex).pos.y))
        End If
        Select Case UserList(AtacanteIndex).clase
            Case e_Class.Hunter
                'if i have an armor equipped
                If UserList(AtacanteIndex).invent.EquippedArmorObjIndex > 0 Then
                    'and the armor has the camouflage property and i have 100 in stealth skill
                    If ObjData(UserList(AtacanteIndex).invent.EquippedArmorObjIndex).Camouflage And UserList(AtacanteIndex).Stats.UserSkills(e_Skill.Ocultarse) = 100 Then
                        'dont remove invisibility
                    Else
                        Call RemoveUserInvisibility(AtacanteIndex)
                    End If
                End If
            Case Else
                Call RemoveUserInvisibility(AtacanteIndex)
        End Select
        Call UserDamageToUser(AtacanteIndex, VictimaIndex, aType)
        'Call EffectsOverTime.TargetDidHit(UserList(AtacanteIndex).EffectOverTime, VictimaIndex, eUser, e_phisical)
        Call RegisterNewAttack(VictimaIndex, AtacanteIndex)
    Else
        Select Case UserList(AtacanteIndex).clase
            Case e_Class.Bandit
                If Not UserList(AtacanteIndex).Stats.UserSkills(e_Skill.Ocultarse) = 100 Then
                    Call RemoveUserInvisibility(AtacanteIndex)
                End If
            Case e_Class.Hunter
                'if i have an armor equipped
                If UserList(AtacanteIndex).invent.EquippedArmorObjIndex > 0 Then
                    'and the armor has the camouflage property and i have 100 in stealth skill
                    If ObjData(UserList(AtacanteIndex).invent.EquippedArmorObjIndex).Camouflage And UserList(AtacanteIndex).Stats.UserSkills(e_Skill.Ocultarse) = 100 Then
                        'dont remove invisibility
                    Else
                        Call RemoveUserInvisibility(AtacanteIndex)
                    End If
                End If
            Case Else
                Call RemoveUserInvisibility(AtacanteIndex)
        End Select
        Call EffectsOverTime.TargetFailedAttack(UserList(AtacanteIndex).EffectOverTime, VictimaIndex, eUser, e_phisical)
        If UserList(AtacanteIndex).flags.invisible Or UserList(AtacanteIndex).flags.Oculto Then
            sendto = SendTarget.ToIndex
        Else
            sendto = SendTarget.ToPCAliveArea
        End If
        Call SendData(sendto, AtacanteIndex, PrepareMessageCharSwing(UserList(AtacanteIndex).Char.charindex, , , IIf(UserList(AtacanteIndex).flags.invisible + UserList( _
                AtacanteIndex).flags.Oculto > 0, False, True)))
    End If
    Exit Sub
UsuarioAtacaUsuario_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtacaUsuario", Erl)
End Sub

Private Sub UserDamageToUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer, ByVal aType As AttackType)
    On Error GoTo UserDañoUser_Err
    With UserList(VictimaIndex)
        Dim Damage As Long, BaseDamage As Long, BonusDamage As Long, Defensa As Long, Color As Long, DamageStr As String, Lugar As e_PartesCuerpo
        ' Daño normal
        BaseDamage = GetUserDamage(AtacanteIndex)
        ' Color por defecto rojo
        Color = vbRed
        ' Elegimos al azar una parte del cuerpo
        Lugar = RandomNumber(1, 8)
        Select Case Lugar
                ' 1/6 de chances de que sea a la cabeza
            Case e_PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If .invent.EquippedHelmetObjIndex > 0 Then
                    Dim Casco As t_ObjData
                    Casco = ObjData(.invent.EquippedHelmetObjIndex)
                    Defensa = Defensa + RandomNumber(Casco.MinDef, Casco.MaxDef)
                End If
            Case Else
                If Lugar > bTorso Then
                    Lugar = RandomNumber(bPiernaIzquierda, bTorso)
                End If
                'Si tiene armadura absorbe el golpe
                If .invent.EquippedArmorObjIndex > 0 Then
                    Dim Armadura As t_ObjData
                    Armadura = ObjData(.invent.EquippedArmorObjIndex)
                    Defensa = Defensa + RandomNumber(Armadura.MinDef, Armadura.MaxDef)
                End If
                'Si tiene escudo absorbe el golpe
                If .invent.EquippedShieldObjIndex > 0 Then
                    Dim Escudo As t_ObjData
                    Escudo = ObjData(.invent.EquippedShieldObjIndex)
                    Defensa = Defensa + RandomNumber(Escudo.MinDef, Escudo.MaxDef)
                End If
        End Select
        ' Defensa del barco de la víctima
        If .invent.EquippedShipObjIndex > 0 Then
            Dim Barco As t_ObjData
            Barco = ObjData(.invent.EquippedShipObjIndex)
            Defensa = Defensa + RandomNumber(Barco.MinDef, Barco.MaxDef)
            ' Defensa de la montura de la víctima
        ElseIf .invent.EquippedSaddleObjIndex > 0 Then
            Dim Montura As t_ObjData
            Montura = ObjData(.invent.EquippedSaddleObjIndex)
            Defensa = Defensa + RandomNumber(Montura.MinDef, Montura.MaxDef)
        End If
        Defensa = Defensa + UserMod.GetDefenseBonus(VictimaIndex)
        Defensa = max(0, Defensa - GetArmorPenetration(AtacanteIndex, Defensa))
        ' Restamos la defensa
        Damage = BaseDamage - Defensa
        Damage = Damage * UserMod.GetPhysicalDamageModifier(UserList(AtacanteIndex))
        Damage = Damage * UserMod.GetPhysicDamageReduction(UserList(VictimaIndex))
        If Damage < 0 Then Damage = 0
        DamageStr = PonerPuntos(Damage)
        ' Mostramos en consola el golpe al atacante solo si tiene activado el chat de combate
        If UserList(AtacanteIndex).ChatCombate = 1 Then
            Call WriteUserHittedUser(AtacanteIndex, Lugar, .Char.charindex, DamageStr)
        End If
        ' Mostramos en consola el golpe a la victima independientemente de la configuración de chat
        Call WriteUserHittedByUser(VictimaIndex, Lugar, UserList(AtacanteIndex).Char.charindex, DamageStr)
        ' Golpe crítico (ignora defensa)
        If PuedeGolpeCritico(AtacanteIndex) Then
            ' Si acertó
            If RandomNumber(1, 100) <= GetCriticalHitChanceAgainstUsers(AtacanteIndex, VictimaIndex) Then
                ' Daño del golpe crítico (usamos el daño base)
                BonusDamage = Damage * CriticalHitDmgModifier
                DamageStr = PonerPuntos(BonusDamage)
                ' Mostramos en consola el daño al atacante
                If UserList(AtacanteIndex).ChatCombate = 1 Then
                    Call WriteLocaleMsg(AtacanteIndex, 383, e_FontTypeNames.FONTTYPE_INFOBOLD, Damage & "¬" & DamageStr)
                End If
                ' Y a la víctima
                If .ChatCombate = 1 Then
                    Call WriteLocaleMsg(VictimaIndex, 385, e_FontTypeNames.FONTTYPE_INFOBOLD, UserList(AtacanteIndex).name & "¬" & DamageStr)
                End If
                Call SendData(SendTarget.ToPCAliveArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO_CRITICO, UserList(AtacanteIndex).pos.x, UserList(AtacanteIndex).pos.y))
                ' Color naranja
                Color = RGB(225, 165, 0)
            End If
            ' Apuñalar (le afecta la defensa)
        ElseIf PuedeApuñalar(AtacanteIndex) Then
            If RandomNumber(1, 100) <= GetStabbingChanceAgainstUsers(AtacanteIndex, VictimaIndex) Then
                ' Daño del apuñalamiento
                BonusDamage = Damage * ModicadorApuñalarClase(UserList(AtacanteIndex).clase)
                DamageStr = PonerPuntos(BonusDamage)
                ' Mostramos en consola el golpe al atacante solo si tiene activado el chat de combate
                If UserList(AtacanteIndex).ChatCombate = 1 Then
                    Call WriteLocaleMsg(AtacanteIndex, "210", e_FontTypeNames.FONTTYPE_INFOBOLD, .name & "¬" & DamageStr)
                End If
                ' Mostramos en consola el golpe a la victima independientemente de la configuración de chat
                Call WriteLocaleMsg(VictimaIndex, "211", e_FontTypeNames.FONTTYPE_INFOBOLD, UserList(AtacanteIndex).name & "¬" & DamageStr)
                'Fx de apuñalar
                Call SendData(SendTarget.ToPCAliveArea, AtacanteIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, FX_STABBING, 0, UserList( _
                        AtacanteIndex).pos.x, UserList(AtacanteIndex).pos.y))
                'Sonido de apuñalar
                Call SendData(SendTarget.ToPCAliveArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO_APU, UserList(AtacanteIndex).pos.x, UserList(AtacanteIndex).pos.y))
                ' Color amarillo
                Color = vbYellow
                ' Efecto en la víctima
                UserList(VictimaIndex).Counters.timeFx = 3
                Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, 89, 0, UserList(VictimaIndex).pos.x, UserList( _
                        VictimaIndex).pos.y))
                ' Efecto en pantalla a ambos
                Call WriteFlashScreen(VictimaIndex, &H3C3CFF, 200, True)
                Call WriteFlashScreen(AtacanteIndex, &H3C3CFF, 150, True)
                Call SendData(SendTarget.ToPCAliveArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(AtacanteIndex).pos.x, UserList(AtacanteIndex).pos.y))
            End If
            ' Sube skills en apuñalar
            Call SubirSkill(AtacanteIndex, Apuñalar)
        End If
        If PuedeDesequiparDeUnGolpe(AtacanteIndex) Then
            If RandomNumber(1, 100) <= ProbabilidadDesequipar(AtacanteIndex) Then
                Call DesequiparObjetoDeUnGolpe(AtacanteIndex, VictimaIndex, Lugar)
            End If
        End If
        If BonusDamage > 0 Then
            Damage = Damage + BonusDamage
            ' Solo si la victima se encuentra en vida completa, generamos la condicion
            If .Stats.MinHp = .Stats.MaxHp Then
                ' Si el daño total es superior a su vida maxima, la victima muere
                If Damage >= .Stats.MaxHp Then
                    Damage = .Stats.MinHp ' Esto simula la muerte (vida minima)
                End If
            End If
        End If
        If UserMod.DoDamageOrHeal(VictimaIndex, AtacanteIndex, e_ReferenceType.eUser, -Damage, e_DamageSourceType.e_phisical, .invent.EquippedWeaponObjIndex, -1, -1, Color) = _
                eStillAlive Then
            'Sonido del golpe
            Call SendData(SendTarget.ToPCAliveArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(AtacanteIndex).pos.x, UserList(AtacanteIndex).pos.y))
            'Fx de sangre del golpe
            Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, FX_BLOOD, 0, UserList(VictimaIndex).pos.x, _
                    UserList(VictimaIndex).pos.y))
            ' Intentamos aplicar algún efecto de estado
            Call UserDañoEspecial(AtacanteIndex, VictimaIndex, aType)
        End If
    End With
    Exit Sub
UserDañoUser_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.UserDañoUser", Erl)
End Sub

Public Function UserDoDamageToUser(ByVal attackerIndex As Integer, _
                                   ByVal TargetIndex As Integer, _
                                   ByVal Damage As Long, _
                                   ByVal Source As e_DamageSourceType, _
                                   ByVal ObjIndex As Integer) As e_DamageResult
    Damage = Damage * UserMod.GetPhysicalDamageModifier(UserList(attackerIndex))
    Damage = Damage * UserMod.GetPhysicDamageReduction(UserList(TargetIndex))
    UserDoDamageToUser = UserMod.DoDamageOrHeal(TargetIndex, attackerIndex, e_ReferenceType.eUser, -Damage, Source, ObjIndex)
    Dim DamageStr As String
    DamageStr = PonerPuntos(Damage)
    If UserList(attackerIndex).ChatCombate = 1 Then
        Call WriteUserHittedUser(attackerIndex, bTorso, UserList(TargetIndex).Char.charindex, DamageStr)
    End If
    Call WriteUserHittedByUser(TargetIndex, bTorso, UserList(attackerIndex).Char.charindex, DamageStr)
End Function

Private Sub DesequiparObjetoDeUnGolpe(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer, ByVal parteDelCuerpo As e_PartesCuerpo)
    On Error GoTo DesequiparObjetoDeUnGolpe_Err
    Dim desequiparCasco As Boolean, desequiparArma As Boolean, desequiparEscudo As Boolean
    With UserList(VictimIndex)
        Select Case parteDelCuerpo
            Case e_PartesCuerpo.bCabeza
                ' Si pega en la cabeza, desequipamos el casco si tiene
                desequiparCasco = .invent.EquippedHelmetObjIndex > 0
                ' Si no tiene casco, intentaremos desequipar otra cosa porque un golpe en la cabeza
                ' algo te tiene que desequipar.
                desequiparArma = (Not desequiparCasco) And (.invent.EquippedWeaponObjIndex > 0)
                desequiparEscudo = (Not desequiparCasco) And (Not desequiparArma) And (.invent.EquippedShieldObjIndex > 0)
            Case e_PartesCuerpo.bBrazoDerecho, e_PartesCuerpo.bBrazoIzquierdo, e_PartesCuerpo.bTorso
                desequiparArma = (.invent.EquippedWeaponObjIndex > 0)
                desequiparEscudo = (Not desequiparArma) And (.invent.EquippedShieldObjIndex > 0)
                desequiparCasco = (Not desequiparEscudo) And (Not desequiparArma) And (.invent.EquippedHelmetObjIndex > 0)
            Case e_PartesCuerpo.bPiernaDerecha, e_PartesCuerpo.bPiernaIzquierda
                desequiparEscudo = (.invent.EquippedShieldObjIndex > 0)
                desequiparArma = (Not desequiparEscudo) And (.invent.EquippedWeaponObjIndex > 0)
                desequiparCasco = (Not desequiparEscudo) And (Not desequiparArma) And (.invent.EquippedHelmetObjIndex > 0)
        End Select
        If desequiparCasco Then
            Call Desequipar(VictimIndex, .invent.EquippedHelmetSlot)
            Call WriteCombatConsoleMsg(attackerIndex, "Has logrado desequipar el casco de tu oponente!")
            Call WriteCombatConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha desequipado el casco.")
            Call CreateUnequip(VictimIndex, eUser, e_InventorySlotMask.eHelm)
        ElseIf desequiparArma Then
            Call Desequipar(VictimIndex, .invent.EquippedWeaponSlot)
            Call WriteCombatConsoleMsg(attackerIndex, "Has logrado desarmar a tu oponente!")
            Call WriteCombatConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha desarmado.")
            Call CreateUnequip(VictimIndex, eUser, e_InventorySlotMask.eWeapon)
        ElseIf desequiparEscudo Then
            Call Desequipar(VictimIndex, .invent.EquippedShieldSlot)
            Call WriteCombatConsoleMsg(attackerIndex, "Has logrado desequipar el escudo de " & .name & ".")
            Call WriteCombatConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha desequipado el escudo.")
            Call CreateUnequip(VictimIndex, eUser, e_InventorySlotMask.eShiled)
        Else
            Call WriteCombatConsoleMsg(attackerIndex, "No has logrado desequipar ningun item a tu oponente!")
        End If
    End With
    Exit Sub
DesequiparObjetoDeUnGolpe_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.DesequiparObjetoDeUnGolpe", Erl)
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer)
    '***************************************************
    'Autor: Unknown
    'Last Modification: 10/01/08
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
    '***************************************************
    On Error GoTo UsuarioAtacadoPorUsuario_Err
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(VictimIndex)
    If UserList(VictimIndex).flags.Meditando Then
        UserList(VictimIndex).flags.Meditando = False
        UserList(VictimIndex).Char.FX = 0
        Call SendData(SendTarget.ToPCAliveArea, VictimIndex, PrepareMessageMeditateToggle(UserList(VictimIndex).Char.charindex, 0))
    End If
    If PeleaSegura(attackerIndex, VictimIndex) Then Exit Sub
    Dim EraCriminal As Byte
    UserList(VictimIndex).Counters.EnCombate = IntervaloEnCombate
    UserList(attackerIndex).Counters.EnCombate = IntervaloEnCombate
    'Si es ciudadano
    If esCiudadano(attackerIndex) Then
        If (esCiudadano(VictimIndex) Or esArmada(VictimIndex)) Then
            Call VolverCriminal(attackerIndex)
        End If
    End If
    EraCriminal = Status(attackerIndex)
    If EraCriminal = 2 And Status(attackerIndex) < 2 Then
        Call RefreshCharStatus(attackerIndex)
    ElseIf EraCriminal < 2 And Status(attackerIndex) = 2 Then
        Call RefreshCharStatus(attackerIndex)
    End If
    If Status(attackerIndex) = e_Facciones.Caos Then If UserList(attackerIndex).Faccion.Status = e_Facciones.Armada Then Call ExpulsarFaccionReal(attackerIndex)
    Call AllMascotasAtacanUser(VictimIndex, attackerIndex)
    Call AllMascotasAtacanUser(attackerIndex, VictimIndex)
    Exit Sub
UsuarioAtacadoPorUsuario_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.UsuarioAtacadoPorUsuario", Erl)
End Sub

Public Function PuedeAtacar(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
    On Error GoTo PuedeAtacar_Err
    '***************************************************
    'Autor: Unknown
    'Last Modification: 24/01/2007
    'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
    '24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
    '***************************************************
    Dim t    As e_Trigger6
    Dim rank As Integer
    'MUY importante el orden de estos "IF"...
    'Estas muerto no podes atacar
    If UserList(attackerIndex).flags.Muerto = 1 Then
        Call WriteLocaleMsg(attackerIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    If UserList(attackerIndex).flags.EnReto Then
        If Retos.Salas(UserList(attackerIndex).flags.SalaReto).TiempoItems > 0 Then
            'Msg1044= No podés atacar en este momento.
            Call WriteLocaleMsg(attackerIndex, "1044", e_FontTypeNames.FONTTYPE_INFO)
            PuedeAtacar = False
            Exit Function
        End If
    End If
    'No podes atacar a alguien muerto
    If UserList(VictimIndex).flags.Muerto = 1 Then
        'Msg1045= No podés atacar a un espiritu.
        Call WriteLocaleMsg(attackerIndex, "1045", e_FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    If UserList(attackerIndex).Grupo.Id > 0 And UserList(VictimIndex).Grupo.Id > 0 And UserList(attackerIndex).Grupo.Id = UserList(VictimIndex).Grupo.Id Then
        'Msg1046= No podés atacar a un miembro de tu grupo.
        Call WriteLocaleMsg(attackerIndex, "1046", e_FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    ' No podes atacar si estas en consulta
    If UserList(attackerIndex).flags.EnConsulta Then
        'Msg1047= No podés atacar usuarios mientras estás en consulta.
        Call WriteLocaleMsg(attackerIndex, "1047", e_FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    ' No podes atacar si esta en consulta
    If UserList(VictimIndex).flags.EnConsulta Then
        'Msg1048= No podés atacar usuarios mientras estan en consulta.
        Call WriteLocaleMsg(attackerIndex, "1048", e_FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    If UserList(attackerIndex).flags.Maldicion = 1 Then
        'Msg1049= ¡Estás maldito! No podes atacar.
        Call WriteLocaleMsg(attackerIndex, "1049", e_FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    If UserList(attackerIndex).flags.Montado = 1 Then
        'Msg1050= No podés atacar usando una montura.
        Call WriteLocaleMsg(attackerIndex, "1050", e_FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False
        Exit Function
    End If
    If Not MapInfo(UserList(VictimIndex).pos.Map).FriendlyFire And UserList(VictimIndex).flags.CurrentTeam > 0 And UserList(VictimIndex).flags.CurrentTeam = UserList( _
            attackerIndex).flags.CurrentTeam Then
        'Msg1051= No podes atacar un miembro de tu equipo.
        Call WriteLocaleMsg(attackerIndex, "1051", e_FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    'Admins y GMs no pueden ser atacados
    rank = e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero
    If (UserList(VictimIndex).flags.Privilegios And rank) > (UserList(attackerIndex).flags.Privilegios And rank) Then
        'Msg1053= El ser es demasiado poderoso
        Call WriteLocaleMsg(attackerIndex, "1053", e_FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    'Estamos en una Arena? o un trigger zona segura?
    t = TriggerZonaPelea(attackerIndex, VictimIndex)
    If t = e_Trigger6.TRIGGER6_PERMITE Then
        PuedeAtacar = True
        Exit Function
    ElseIf PeleaSegura(attackerIndex, VictimIndex) Then
        PuedeAtacar = True
        Exit Function
    ElseIf t = e_Trigger6.TRIGGER6_PROHIBE Then
        PuedeAtacar = False
        Exit Function
    End If
    'Solo administradores pueden atacar a usuarios (PARA TESTING)
    If (UserList(attackerIndex).flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Admin)) = 0 Then
        PuedeAtacar = False
        Exit Function
    End If
    ' Seguro Clan
    If UserList(attackerIndex).GuildIndex > 0 Then
        If UserList(attackerIndex).flags.SeguroClan And NivelDeClan(UserList(attackerIndex).GuildIndex) >= RequiredGuildLevelSafe Then
            If UserList(attackerIndex).GuildIndex = UserList(VictimIndex).GuildIndex Then
                'Msg1054= No podes atacar a un miembro de tu clan.
                Call WriteLocaleMsg(attackerIndex, "1054", e_FontTypeNames.FONTTYPE_INFOIAO)
                PuedeAtacar = False
                Exit Function
            End If
        End If
    End If
    ' Es armada?
    If esArmada(attackerIndex) Then
        ' Si ataca otro armada
        If esArmada(VictimIndex) Then
            'Msg1055= Los miembros del Ejercito Real tienen prohibido atacarse entre sí.
            Call WriteLocaleMsg(attackerIndex, "1055", e_FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = False
            Exit Function
            ' Si ataca un ciudadano
        ElseIf esCiudadano(VictimIndex) Then
            'Msg1056= Los miembros del Ejercito Real tienen prohibido atacar ciudadanos.
            Call WriteLocaleMsg(attackerIndex, "1056", e_FontTypeNames.FONTTYPE_WARNING)
            PuedeAtacar = False
            Exit Function
        End If
        ' No es armada
    Else
        'Tenes puesto el seguro?
        If (esCiudadano(attackerIndex)) Then
            If (UserList(attackerIndex).flags.Seguro) Then
                If esCiudadano(VictimIndex) Then
                    'Msg1057= No podés atacar ciudadanos, para hacerlo debes desactivar el seguro.
                    Call WriteLocaleMsg(attackerIndex, "1057", e_FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = False
                    Exit Function
                ElseIf esArmada(VictimIndex) Then
                    'Msg1058= No podés atacar miembros del Ejercito Real, para hacerlo debes desactivar el seguro.
                    Call WriteLocaleMsg(attackerIndex, "1058", e_FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = False
                    Exit Function
                End If
            End If
        ElseIf esCaos(attackerIndex) And esCaos(VictimIndex) Then
            If Not (UserList(attackerIndex).flags.LegionarySecure) Then
                PuedeAtacar = True
            ElseIf MapInfo(UserList(VictimIndex).pos.Map).Seguro <> 1 Then
                'Msg1059= Los miembros de las Fuerzas del Caos no se pueden atacar entre sí.
                Call WriteLocaleMsg(attackerIndex, "1059", e_FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function
            End If
        End If
    End If
    'Estas en un Mapa Seguro?
    If MapInfo(UserList(VictimIndex).pos.Map).Seguro = 1 Then
        If esArmada(attackerIndex) Then
            If UserList(attackerIndex).Faccion.RecompensasReal >= 3 Then
                If UserList(VictimIndex).pos.Map = 58 Or UserList(VictimIndex).pos.Map = 59 Or UserList(VictimIndex).pos.Map = 60 Then
                    'Msg1060= Huye de la ciudad! estas siendo atacado y no podrás defenderte.
                    Call WriteLocaleMsg(VictimIndex, "1060", e_FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = True 'Beneficio de Armadas que atacan en su ciudad.
                    Exit Function
                End If
            End If
        End If
        If esCaos(attackerIndex) Then
            If UserList(attackerIndex).Faccion.RecompensasCaos >= 3 Then
                If UserList(VictimIndex).pos.Map = 195 Or UserList(VictimIndex).pos.Map = 196 Then
                    'Msg1061= Huye de la ciudad! estas siendo atacado y no podrás defenderte.
                    Call WriteLocaleMsg(VictimIndex, "1061", e_FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = True 'Beneficio de Caos que atacan en su ciudad.
                    Exit Function
                End If
            End If
        End If
        'Msg1062= Esta es una zona segura, aqui no podes atacar otros usuarios.
        Call WriteLocaleMsg(attackerIndex, "1062", e_FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
    If MapData(UserList(VictimIndex).pos.Map, UserList(VictimIndex).pos.x, UserList(VictimIndex).pos.y).trigger = e_Trigger.ZonaSegura Or MapData(UserList( _
            attackerIndex).pos.Map, UserList(attackerIndex).pos.x, UserList(attackerIndex).pos.y).trigger = e_Trigger.ZonaSegura Then
        'Msg1063= No podes pelear aqui.
        Call WriteLocaleMsg(attackerIndex, "1063", e_FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False
        Exit Function
    End If
    PuedeAtacar = True
    Exit Function
PuedeAtacar_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.PuedeAtacar", Erl)
End Function

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
    On Error GoTo CalcularDarExp_Err
    If NpcList(NpcIndex).MaestroUser.ArrayIndex <> 0 Then
        Exit Sub
    End If
    If UserList(UserIndex).Grupo.EnGrupo Then
        Call CalcularDarExpGrupal(UserIndex, NpcIndex, ElDaño)
    Else
        Call GetExpForUser(UserIndex, NpcIndex, ElDaño)
    End If
    Exit Sub
CalcularDarExp_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.CalcularDarExp", Erl)
End Sub

Private Sub GetExpForUser(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
    On Error GoTo GetExpForUser_Err
    Dim ExpaDar As Double
    With UserList(UserIndex)
        'Chekeamos que las variables sean validas para las operaciones
        If ElDaño <= 0 Then ElDaño = 0
        If NpcList(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
        'La experiencia a dar es la porcion de vida quitada * toda la experiencia
        ExpaDar = CDbl(ElDaño) * CDbl(NpcList(NpcIndex).GiveEXP) / NpcList(NpcIndex).Stats.MaxHp
        If ExpaDar <= 0 Then Exit Sub
        'Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
        'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
        'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
        If ExpaDar > NpcList(NpcIndex).flags.ExpCount Then
            ExpaDar = NpcList(NpcIndex).flags.ExpCount
            NpcList(NpcIndex).flags.ExpCount = 0
        Else
            NpcList(NpcIndex).flags.ExpCount = NpcList(NpcIndex).flags.ExpCount - ExpaDar
        End If
        If SvrConfig.GetValue("ExpMult") > 0 Then
            ExpaDar = ExpaDar * SvrConfig.GetValue("ExpMult")
        End If
        If ExpaDar > 0 Then
            If NpcList(NpcIndex).nivel Then
                Dim DeltaLevel As Integer
                DeltaLevel = .Stats.ELV - NpcList(NpcIndex).nivel
                If DeltaLevel > CInt(SvrConfig.GetValue("NpcDeltaLevelPenalties")) Then
                    Dim Penalty As Single
                    Penalty = GetExpPenalty(UserIndex, NpcIndex, DeltaLevel)
                    ExpaDar = ExpaDar * Penalty
                    ' Si tiene el chat activado, enviamos el mensaje
                    If UserList(UserIndex).ChatCombate = 1 Then
                        ' Mostrar porcentaje final de experiencia como número entero
                        Dim PorcentajeFinal As Integer
                        PorcentajeFinal = Penalty * 100
                        'Msg1467=Debido a tu nivel, obtienes el ¬1% de la experiencia.
                        Call WriteLocaleMsg(UserIndex, 1467, e_FontTypeNames.FONTTYPE_WARNING, PorcentajeFinal)
                    End If
                End If
            End If
            If .Stats.ELV < STAT_MAXELV Then
                .Stats.Exp = .Stats.Exp + ExpaDar
                If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
                Call WriteUpdateExp(UserIndex)
                Call CheckUserLevel(UserIndex)
            End If
            Call WriteTextOverTile(UserIndex, "+" & PonerPuntos(ExpaDar), .pos.x, .pos.y, RGB(0, 169, 255))
        End If
    End With
    Exit Sub
GetExpForUser_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.GetExpForUser", Erl)
End Sub

Private Sub CalcularDarExpGrupal(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
    On Error GoTo CalcularDarExpGrupal_Err
    Dim ExpaDar                 As Long
    Dim BonificacionGrupo       As Single
    Dim CantidadMiembrosValidos As Integer
    Dim i                       As Long
    Dim Index                   As Integer
    'If UserList(UserIndex).Grupo.EnGrupo Then
    'Chekeamos que las variables sean validas para las operaciones
    If NpcIndex = 0 Then Exit Sub
    If UserIndex = 0 Then Exit Sub
    If ElDaño <= 0 Then ElDaño = 0
    If NpcList(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
    If ElDaño > NpcList(NpcIndex).Stats.MinHp Then ElDaño = NpcList(NpcIndex).Stats.MinHp
    'La experiencia a dar es la porcion de vida quitada * toda la experiencia
    ExpaDar = CLng((ElDaño) * (NpcList(NpcIndex).GiveEXP / NpcList(NpcIndex).Stats.MaxHp))
    If ExpaDar <= 0 Then Exit Sub
    'Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
    'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
    'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
    If ExpaDar > NpcList(NpcIndex).flags.ExpCount Then
        ExpaDar = NpcList(NpcIndex).flags.ExpCount
        NpcList(NpcIndex).flags.ExpCount = 0
    Else
        NpcList(NpcIndex).flags.ExpCount = NpcList(NpcIndex).flags.ExpCount - ExpaDar
    End If
    With UserList(UserIndex)
        If Not IsValidUserRef(.Grupo.Lider) Then Exit Sub
        Dim LiderIndex As Integer
        LiderIndex = .Grupo.Lider.ArrayIndex
        For i = 1 To UserList(LiderIndex).Grupo.CantidadMiembros
            If IsValidUserRef(UserList(LiderIndex).Grupo.Miembros(i)) Then
                Index = UserList(LiderIndex).Grupo.Miembros(i).ArrayIndex
                If UserList(Index).flags.Muerto = 0 Then
                    If .pos.Map = UserList(Index).pos.Map Then
                        If Distancia(.pos, UserList(Index).pos) < 20 Then
                            CantidadMiembrosValidos = CantidadMiembrosValidos + 1
                        End If
                    End If
                End If
            End If
        Next
        ' Verificar si el líder está en otro mapa
        If UserList(LiderIndex).pos.Map <> .pos.Map Then
            CantidadMiembrosValidos = CantidadMiembrosValidos + 1 ' Se cuenta como un miembro más para dividir la exp
            ' Avisamos a los miembros del grupo
            For i = 1 To UserList(LiderIndex).Grupo.CantidadMiembros
                If IsValidUserRef(UserList(LiderIndex).Grupo.Miembros(i)) Then
                    Index = UserList(LiderIndex).Grupo.Miembros(i).ArrayIndex
                    ' Enviar el mensaje solo si el miembro no está muerto y tiene el chat de combate activado
                    If UserList(Index).flags.Muerto = 0 And UserList(Index).ChatCombate = 1 Then
                        'Msg1437=El líder del grupo está demasiado lejos, su experiencia se pierde.
                        Call WriteLocaleMsg(Index, "1437", e_FontTypeNames.FONTTYPE_EXP)
                    End If
                End If
            Next i
        End If
        If CantidadMiembrosValidos = 0 Then Exit Sub
        If SvrConfig.GetValue("ExpMult") > 0 Then
            ExpaDar = ExpaDar * SvrConfig.GetValue("ExpMult")
        End If
        ExpaDar = ExpaDar / CantidadMiembrosValidos
        Dim ExpUser As Long, DeltaLevel As Integer, ExpBonusForUser As Double
        If ExpaDar > 0 Then
            For i = 1 To UserList(LiderIndex).Grupo.CantidadMiembros
                If IsValidUserRef(UserList(LiderIndex).Grupo.Miembros(i)) Then
                    Index = UserList(LiderIndex).Grupo.Miembros(i).ArrayIndex
                    If UserList(Index).flags.Muerto = 0 Then
                        If Distancia(.pos, UserList(Index).pos) < 20 Then
                            ExpUser = ExpaDar
                            If UserList(Index).Stats.ELV < STAT_MAXELV Then
                                If NpcList(NpcIndex).nivel Then
                                    DeltaLevel = UserList(Index).Stats.ELV - NpcList(NpcIndex).nivel
                                    If DeltaLevel > CInt(SvrConfig.GetValue("NpcDeltaLevelPenalties")) Then
                                        Dim Penalty As Single
                                        Penalty = GetExpPenalty(Index, NpcIndex, DeltaLevel)
                                        ExpUser = ExpUser * Penalty
                                        ' Si tiene el chat activado, enviamos el mensaje
                                        If UserList(Index).ChatCombate = 1 Then
                                            ' Mostrar porcentaje final de experiencia como número entero
                                            Dim PorcentajeFinal As Integer
                                            PorcentajeFinal = Penalty * 100
                                            'Msg1467=Debido a tu nivel, obtienes el ¬1% de la experiencia.
                                            Call WriteLocaleMsg(Index, "1467", e_FontTypeNames.FONTTYPE_WARNING, PorcentajeFinal)
                                        End If
                                    End If
                                End If
                                If (UserList(Index).Stats.UserSkills(e_Skill.liderazgo) >= (15 - UserList(Index).Stats.UserAtributos(e_Atributos.Carisma) / 2)) Then
                                    ExpBonusForUser = ExpUser * SvrConfig.GetValue("LeadershipExpPartyBonus")
                                    UserList(Index).Stats.Exp = UserList(Index).Stats.Exp + ExpBonusForUser
                                Else
                                    UserList(Index).Stats.Exp = UserList(Index).Stats.Exp + ExpUser
                                End If
                                If UserList(Index).Stats.Exp > MAXEXP Then UserList(Index).Stats.Exp = MAXEXP
                                If UserList(Index).ChatCombate = 1 Then
                                    Call WriteLocaleMsg(Index, "141", e_FontTypeNames.FONTTYPE_EXP, ExpUser)
                                End If
                                Call WriteUpdateExp(Index)
                                Call CheckUserLevel(Index)
                            End If
                        Else
                            If UserList(Index).ChatCombate = 1 Then
                                Call WriteLocaleMsg(Index, "69", e_FontTypeNames.FONTTYPE_New_GRUPO)
                            End If
                        End If
                    Else
                        If UserList(Index).ChatCombate = 1 Then
                            'Msg1064= Estás muerto, no has ganado experencia del grupo.
                            Call WriteLocaleMsg(Index, "1064", e_FontTypeNames.FONTTYPE_New_GRUPO)
                        End If
                    End If
                End If
            Next i
        End If
    End With
    Exit Sub
CalcularDarExpGrupal_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.CalcularDarExpGrupal", Erl)
End Sub

Function GetExpPenalty(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, DeltaLevel As Integer) As Single
    On Error GoTo GetExpPenalty_Err
    '    This function computes an experience-gain multiplier (between 0.0 and 1.0) based on how far above the NPC’s level the player is.
    '    Why “DeltaLevel – 4”? No penalty for small over-leveling (up to 4 levels). Beyond that, each extra level reduces your XP by the configured percentage.
    '    Summary
    '    Output: a multiplier from 1.0 down to 0.0.
    '    Use it to scale whatever EXP your damage routine calculated.
    Dim NivelesExtra As Integer
    NivelesExtra = DeltaLevel - 4
    ' Calculamos el porcentaje de penalización
    Dim Penalizacion As Single
    Penalizacion = 1 - (CSng(SvrConfig.GetValue("PenaltyExpUserPerLevel")) * NivelesExtra)
    ' Nos aseguramos de que nunca sea menos del 0%
    If Penalizacion < 0 Then Penalizacion = 0
    GetExpPenalty = Penalizacion
    Exit Function
GetExpPenalty_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.GetExpPenalty", Erl)
End Function

Private Sub CalcularDarOroGrupal(ByVal UserIndex As Integer, ByVal GiveGold As Long)
    On Error GoTo CalcularDarOroGrupal_Err
    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/09/06 Nacho
    'Reescribi gran parte del Sub
    'Ahora, da toda la experiencia del npc mientras este vivo.
    '***************************************************
    Dim OroDar As Long
    OroDar = GiveGold * SvrConfig.GetValue("GoldMult")
    Dim orobackup As Long
    orobackup = OroDar
    Dim i     As Byte
    Dim Index As Byte
    Dim Lider As Integer
    Lider = UserList(UserIndex).Grupo.Lider.ArrayIndex
    OroDar = OroDar / UserList(UserList(UserIndex).Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros
    For i = 1 To UserList(Lider).Grupo.CantidadMiembros
        If IsValidUserRef(UserList(Lider).Grupo.Miembros(i)) Then
            Index = UserList(Lider).Grupo.Miembros(i).ArrayIndex
            If UserList(Index).flags.Muerto = 0 Then
                If UserList(UserIndex).pos.Map = UserList(Index).pos.Map Then
                    If OroDar > 0 Then
                        UserList(Index).Stats.GLD = UserList(Index).Stats.GLD + OroDar
                        If UserList(Index).ChatCombate = 1 Then
                            Call WriteConsoleMsg(Index, PrepareMessageLocaleMsg(1980, PonerPuntos(OroDar), e_FontTypeNames.FONTTYPE_New_GRUPO)) ' Msg1780=¡El grupo ha ganado ¬1 monedas de oro!
                        End If
                        Call WriteUpdateGold(Index)
                    End If
                End If
            End If
        End If
    Next i
    Exit Sub
CalcularDarOroGrupal_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.CalcularDarOroGrupal", Erl)
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As e_Trigger6
    On Error GoTo ErrHandler
    Dim tOrg As e_Trigger
    Dim tDst As e_Trigger
    tOrg = MapData(UserList(Origen).pos.Map, UserList(Origen).pos.x, UserList(Origen).pos.y).trigger
    tDst = MapData(UserList(Destino).pos.Map, UserList(Destino).pos.x, UserList(Destino).pos.y).trigger
    If tOrg = e_Trigger.ZONAPELEA Or tDst = e_Trigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If
    Exit Function
ErrHandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.Description)
End Function

Public Function PeleaSegura(ByVal Source As Integer, ByVal dest As Integer) As Boolean
    If MapInfo(UserList(Source).pos.Map).SafeFightMap Then
        PeleaSegura = True
    Else
        PeleaSegura = TriggerZonaPelea(Source, dest) = TRIGGER6_PERMITE
    End If
End Function

Private Sub UserDañoEspecial(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer, ByVal aType As AttackType)
    On Error GoTo UserDañoEspecial_Err
    Dim ArmaObjInd As Integer, ObjInd As Integer
    ArmaObjInd = UserList(AtacanteIndex).invent.EquippedWeaponObjIndex
    ObjInd = 0
    ' Preguntamos una vez mas, si no tiene Nudillos o Arma, no tiene sentido seguir.
    If ArmaObjInd = 0 Then
        Exit Sub
    End If
    If ObjData(ArmaObjInd).Proyectil = 0 Or ObjData(ArmaObjInd).Municion = 0 Then
        ObjInd = ArmaObjInd
    Else
        ObjInd = UserList(AtacanteIndex).invent.EquippedMunitionObjIndex
    End If
    Dim puedeEnvenenar, puedeEstupidizar, puedeIncinierar, puedeParalizar, rangeStun As Boolean
    Dim stunChance As Byte
    puedeEnvenenar = (UserList(AtacanteIndex).flags.Envenena > 0) Or (ObjInd > 0 And ObjData(ObjInd).Envenena)
    puedeEstupidizar = (UserList(AtacanteIndex).flags.Estupidiza > 0) Or (ObjInd > 0 And ObjData(ObjInd).Estupidiza)
    puedeIncinierar = (UserList(AtacanteIndex).flags.incinera > 0) Or (ObjInd > 0 And ObjData(ObjInd).incinera)
    puedeParalizar = (UserList(AtacanteIndex).flags.Paraliza > 0) Or (ObjInd > 0 And ObjData(ObjInd).Paraliza)
    If ObjInd > 0 Then
        rangeStun = ObjData(ObjInd).Subtipo = 2 And aType = Ranged
        stunChance = ObjData(ObjInd).Porcentaje
    End If
    If puedeEnvenenar And (UserList(VictimaIndex).flags.Envenenado = 0) Then
        If RandomNumber(1, 100) < 30 Then
            UserList(VictimaIndex).flags.Envenenado = ObjData(ObjInd).Envenena
            Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha envenenado!")
            Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has envenenado a " & UserList(VictimaIndex).name & "!")
            Exit Sub
        End If
    End If
    If puedeIncinierar And (UserList(VictimaIndex).flags.Incinerado = 0) Then
        If RandomNumber(1, 100) < 10 Then
            UserList(VictimaIndex).flags.Incinerado = 1
            UserList(VictimaIndex).Counters.Incineracion = 1
            Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha Incinerado!")
            Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has Incinerado a " & UserList(VictimaIndex).name & "!")
            Exit Sub
        End If
    End If
    If puedeParalizar And (UserList(VictimaIndex).flags.Paralizado = 0) And Not IsSet(UserList(VictimaIndex).flags.StatusMask, eCCInmunity) Then
        If RandomNumber(1, 100) < 10 Then
            UserList(VictimaIndex).flags.Paralizado = 1
            UserList(VictimaIndex).Counters.Paralisis = 6
            Call WriteParalizeOK(VictimaIndex)
            UserList(VictimaIndex).Counters.timeFx = 3
            Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, 8, 0, UserList(VictimaIndex).pos.x, UserList( _
                    VictimaIndex).pos.y))
            Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha paralizado!")
            Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has paralizado a " & UserList(VictimaIndex).name & "!")
            Exit Sub
        End If
    End If
    If puedeEstupidizar And (UserList(VictimaIndex).flags.Estupidez = 0) Then
        If RandomNumber(1, 100) < 13 Then
            UserList(VictimaIndex).flags.Estupidez = 1
            UserList(VictimaIndex).Counters.Estupidez = 3 ' segundos?
            Call WriteDumb(VictimaIndex)
            UserList(VictimaIndex).Counters.timeFx = 3
            Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageParticleFX(UserList(VictimaIndex).Char.charindex, 30, 30, False, , UserList(VictimaIndex).pos.x, _
                    UserList(VictimaIndex).pos.y))
            Call WriteCombatConsoleMsg(VictimaIndex, "¡" & UserList(AtacanteIndex).name & " te ha estupidizado!")
            Call WriteCombatConsoleMsg(AtacanteIndex, "¡Has estupidizado a " & UserList(VictimaIndex).name & "!")
            Exit Sub
        End If
    End If
    If rangeStun And Not IsSet(UserList(VictimaIndex).flags.StatusMask, eCCInmunity) Then
        If (RandomNumber(1, 100) < stunChance) Then
            With UserList(VictimaIndex)
                If StunPlayer(VictimaIndex, .Counters) Then
                    Call WriteStunStart(VictimaIndex, PlayerStunTime)
                    Call WritePosUpdate(VictimaIndex)
                    Call SendData(SendTarget.ToPCAliveArea, VictimaIndex, PrepareMessageCreateFX(.Char.charindex, 142, 1))
                End If
            End With
        End If
    End If
    Exit Sub
UserDañoEspecial_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.UserDañoEspecial", Erl)
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
    'Reaccion de las mascotas
    On Error GoTo AllMascotasAtacanUser_Err
    Dim iCount       As Long
    Dim mascotaIndex As Integer
    With UserList(Maestro)
        For iCount = 1 To MAXMASCOTAS
            mascotaIndex = .MascotasIndex(iCount).ArrayIndex
            If mascotaIndex > 0 Then
                If IsValidNpcRef(.MascotasIndex(iCount)) Then
                    If IsSet(NpcList(mascotaIndex).flags.BehaviorFlags, e_BehaviorFlags.eAttackUsers) Then
                        NpcList(mascotaIndex).flags.AttackedBy = UserList(victim).name
                        NpcList(mascotaIndex).flags.AttackedTime = GlobalFrameTime
                        Call SetUserRef(NpcList(mascotaIndex).TargetUser, victim)
                        Call SetMovement(mascotaIndex, e_TipoAI.NpcDefensa)
                        NpcList(mascotaIndex).Hostile = 0
                        NpcList(mascotaIndex).flags.OldHostil = 0
                    End If
                Else
                    Call ClearNpcRef(.MascotasIndex(iCount))
                End If
            End If
        Next iCount
    End With
    Exit Sub
AllMascotasAtacanUser_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.AllMascotasAtacanUser", Erl)
End Sub

Public Sub AllMascotasAtacanNPC(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    On Error GoTo AllMascotasAtacanNPC_Err
    Dim j                           As Long
    Dim mascotaIdx                  As Integer
    Dim UserAttackInteractionResult As t_AttackInteractionResult
    UserAttackInteractionResult = UserCanAttackNpc(UserIndex, NpcIndex)
    Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResult.Result)
    If UserAttackInteractionResult.CanAttack Then
        If UserAttackInteractionResult.TurnPK Then Call VolverCriminal(UserIndex)
    Else
        Exit Sub
    End If
    For j = 1 To MAXMASCOTAS
        If IsValidNpcRef(UserList(UserIndex).MascotasIndex(j)) Then
            mascotaIdx = UserList(UserIndex).MascotasIndex(j).ArrayIndex
            If mascotaIdx > 0 And mascotaIdx <> NpcIndex Then
                With NpcList(mascotaIdx)
                    If IsSet(.flags.BehaviorFlags, e_BehaviorFlags.eAttackNpc) And Not IsValidNpcRef(.TargetNPC) Then
                        Call SetNpcRef(.TargetNPC, NpcIndex)
                        Call SetMovement(mascotaIdx, e_TipoAI.NpcAtacaNpc)
                        NpcList(NpcIndex).flags.AttackedBy = UserList(UserIndex).name
                        NpcList(NpcIndex).flags.AttackedTime = GlobalFrameTime
                        Call SetNpcRef(UserList(UserIndex).flags.NPCAtacado, NpcIndex)
                    End If
                End With
            End If
        End If
    Next j
    Exit Sub
AllMascotasAtacanNPC_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.AllMascotasAtacanNPC", Erl)
End Sub

Private Function PuedeDesequiparDeUnGolpe(ByVal UserIndex As Integer) As Boolean
    On Error GoTo PuedeDesequiparDeUnGolpe_Err
    With UserList(UserIndex)
        If .invent.EquippedWeaponObjIndex > 0 Then
            If ObjData(.invent.EquippedWeaponObjIndex).WeaponType <> eKnuckle Then
                PuedeDesequiparDeUnGolpe = False
                Exit Function
            End If
        End If
        Select Case .clase
            Case e_Class.Bandit, e_Class.Thief
                ' PuedeDesequiparDeUnGolpe = (.Stats.UserSkills(e_Skill.Wrestling) >= 100)
                ' Shugar: Hago que pueda desequipar desde nivel 1 y modifico
                ' la probabilidad de desequipar en ProbabilidadDesequipar
                PuedeDesequiparDeUnGolpe = True
            Case Else
                PuedeDesequiparDeUnGolpe = False
        End Select
    End With
    Exit Function
PuedeDesequiparDeUnGolpe_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.PuedeDesequiparDeUnGolpe", Erl)
End Function

Private Function PuedeApuñalar(ByVal UserIndex As Integer) As Boolean
    On Error GoTo PuedeApuñalar_Err
    With UserList(UserIndex)
        If .invent.EquippedWeaponObjIndex > 0 Then
            PuedeApuñalar = (.clase = e_Class.Assasin Or .Stats.UserSkills(e_Skill.Apuñalar) >= MIN_APUÑALAR) And ObjData(.invent.EquippedWeaponObjIndex).Apuñala = 1
        End If
    End With
    Exit Function
PuedeApuñalar_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.PuedeApuñalar", Erl)
End Function

Private Function PuedeGolpeCritico(ByVal UserIndex As Integer) As Boolean
    ' Autor: WyroX - 16/01/2021
    On Error GoTo PuedeGolpeCritico_Err
    With UserList(UserIndex)
        If .invent.EquippedWeaponObjIndex > 0 Then
            PuedeGolpeCritico = .clase = e_Class.Bandit And ObjData(.invent.EquippedWeaponObjIndex).WeaponType = e_WeaponType.eKnuckle
        End If
    End With
    Exit Function
PuedeGolpeCritico_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.PuedeGolpeCritico", Erl)
End Function

Private Function GetSkillRequiredForWeapon(ByVal ObjId As Integer) As e_Skill
    If ObjId = 0 Then
        GetSkillRequiredForWeapon = e_Skill.Wrestling
    Else
        Select Case ObjData(ObjId).WeaponType
            Case e_WeaponType.eKnuckle
                GetSkillRequiredForWeapon = e_Skill.Wrestling
            Case e_WeaponType.eBow
                GetSkillRequiredForWeapon = e_Skill.Proyectiles
            Case e_WeaponType.eGunPowder
                GetSkillRequiredForWeapon = e_Skill.Proyectiles
            Case e_WeaponType.eDagger
                GetSkillRequiredForWeapon = e_Skill.Apuñalar
            Case Else
                GetSkillRequiredForWeapon = e_Skill.Armas
        End Select
    End If
End Function

Private Function ProbabilidadDesequipar(ByVal UserIndex As Integer) As Integer
    On Error GoTo ProbabilidadDesequipar_Err
    With UserList(UserIndex)
        Select Case .clase
            Case e_Class.Bandit
                If IsFeatureEnabled("bandit_unequip_bonus") Then
                    ' Shugar: Hago que la probabilidad de desequipar sea proporcional a los skills
                    ' requeridos por el arma, en este caso combate sin armas para nudillos
                    ProbabilidadDesequipar = 0.2 * UserList(UserIndex).Stats.UserSkills(GetSkillRequiredForWeapon(UserList(UserIndex).invent.EquippedWeaponObjIndex))
                Else
                    ProbabilidadDesequipar = 0.15 * UserList(UserIndex).Stats.UserSkills(GetSkillRequiredForWeapon(UserList(UserIndex).invent.EquippedWeaponObjIndex))
                End If
            Case e_Class.Thief
                ProbabilidadDesequipar = 0.33 * 100
            Case Else
                ProbabilidadDesequipar = 0
        End Select
    End With
    Exit Function
ProbabilidadDesequipar_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.ProbabilidadDesequipar", Erl)
End Function

' Helper function to simplify the code. Keep private!
Private Sub WriteCombatConsoleMsg(ByVal UserIndex As Integer, ByVal Message As String)
    On Error GoTo WriteCombatConsoleMsg_Err
    If UserList(UserIndex).ChatCombate = 1 Then
        Call WriteConsoleMsg(UserIndex, Message, e_FontTypeNames.FONTTYPE_FIGHT)
    End If
    Exit Sub
WriteCombatConsoleMsg_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.WriteCombatConsoleMsg", Erl)
End Sub

Public Function MultiShot(ByVal UserIndex As Integer, ByRef TargetPos As t_WorldPos) As Boolean
    On Error GoTo MultiShot_Err
    With UserList(UserIndex)
        Dim ArrowSlot As Integer
        ArrowSlot = .invent.EquippedMunitionSlot
        If ArrowSlot = 0 Then
            Call WriteLocaleMsg(UserIndex, MsgEquipedArrowRequired, FONTTYPE_INFO)
            Exit Function
        End If
        If ArrowSlot = 0 Then
            Call WriteLocaleMsg(UserIndex, MsgEquipedArrowRequired, FONTTYPE_INFO)
            Exit Function
        End If
        If ObjData(.invent.Object(ArrowSlot).ObjIndex).Subtipo <> ObjData(.invent.EquippedWeaponObjIndex).Municion Then
            Call WriteLocaleMsg(UserIndex, MsgEquipedArrowRequired, FONTTYPE_INFO)
            Exit Function
        End If
        If .invent.Object(ArrowSlot).amount < 5 Then
            Call WriteLocaleMsg(UserIndex, MsgNotEnoughtAmunitions, FONTTYPE_INFO)
            Exit Function
        End If
        Dim Direction  As t_Vector
        Dim RotatedDir As t_Vector
        Direction.x = TargetPos.x - .pos.x
        Direction.y = TargetPos.y - .pos.y
        If Direction.x = 0 And Direction.y = 0 Then
            Call WriteLocaleMsg(UserIndex, MsgCantAttackYourself, FONTTYPE_INFO)
            Exit Function
        End If
        Direction = GetNormal(Direction)
        Call IncreaseSingle(.Modifiers.PhysicalDamageBonus, -MultiShotReduction) 'reduce the damage we deal for each arrow (we can hit multiple time the same target)
        Call ThrowArrowToTargetDir(UserIndex, Direction, 10)
        RotatedDir = RotateVector(Direction, ToRadians(-15))
        Call ThrowArrowToTargetDir(UserIndex, RotatedDir, 10)
        RotatedDir = RotateVector(Direction, ToRadians(-30))
        Call ThrowArrowToTargetDir(UserIndex, RotatedDir, 10)
        RotatedDir = RotateVector(Direction, ToRadians(15))
        Call ThrowArrowToTargetDir(UserIndex, RotatedDir, 10)
        RotatedDir = RotateVector(Direction, ToRadians(30))
        Call ThrowArrowToTargetDir(UserIndex, RotatedDir, 10)
        Call IncreaseSingle(.Modifiers.PhysicalDamageBonus, MultiShotReduction)  'back to normal
    End With
    MultiShot = True
    Exit Function
MultiShot_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.MultiShot", Erl)
End Function

Public Sub ThrowArrowToTargetDir(ByVal UserIndex As Integer, ByRef Direction As t_Vector, ByVal Distance As Integer)
    On Error GoTo ThrowArrowToTargetDir_Err
    Dim currentPos        As t_WorldPos
    Dim TargetPoint       As t_Vector
    Dim TargetTranslation As t_Vector
    Dim TargetPos         As t_WorldPos
    Dim TranslationDiff   As Double
    Dim Tanslation        As Integer
    currentPos = UserList(UserIndex).pos
    TargetPos.Map = currentPos.Map
    Dim step As Integer
    For step = 1 To Distance
        TargetPoint.x = Direction.x * (step) + UserList(UserIndex).pos.x
        TargetPoint.y = Direction.y * (step) + UserList(UserIndex).pos.y
        TargetTranslation.x = TargetPoint.x - currentPos.x
        TargetTranslation.y = TargetPoint.y - currentPos.y
        TranslationDiff = Abs(TargetTranslation.x) - Abs(TargetTranslation.y)
        If Abs(TranslationDiff) < 0.3 Then 'if they are similar we are close to 45% let move in both directions
            TargetPos.x = currentPos.x + Sgn(TargetTranslation.x)
            TargetPos.y = currentPos.y + Sgn(TargetTranslation.y)
        ElseIf TranslationDiff > 0 Then 'x axis is bigger than
            TargetPos.x = currentPos.x + Sgn(TargetTranslation.x)
            TargetPos.y = currentPos.y
        Else
            TargetPos.x = currentPos.x
            TargetPos.y = currentPos.y + Sgn(TargetTranslation.y)
        End If
        If ThrowArrowToTile(UserIndex, TargetPos) Then
            Exit Sub
        End If
        currentPos = TargetPos
    Next step
    Call ConsumeAmunition(UserIndex)
    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, TargetPos.x, TargetPos.y, GetProjectileView( _
            UserList(UserIndex))))
    Exit Sub
ThrowArrowToTargetDir_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.ThrowArrowToTargetDir", Erl)
End Sub

Public Function ThrowArrowToTile(ByVal UserIndex As Integer, ByRef TargetPos As t_WorldPos) As Boolean
    On Error GoTo ThrowArrowToTile_Err
    ThrowArrowToTile = False
    If MapData(TargetPos.Map, TargetPos.x, TargetPos.y).UserIndex > 0 Then
        If UserMod.CanAttackUser(UserIndex, UserList(UserIndex).VersionId, MapData(TargetPos.Map, TargetPos.x, TargetPos.y).UserIndex, UserList(MapData(TargetPos.Map, _
                TargetPos.x, TargetPos.y).UserIndex).VersionId) = eCanAttack Then
            Call ThrowProjectileToTarget(UserIndex, MapData(TargetPos.Map, TargetPos.x, TargetPos.y).UserIndex, eUser)
            ThrowArrowToTile = True
        End If
    ElseIf MapData(TargetPos.Map, TargetPos.x, TargetPos.y).NpcIndex > 0 Then
        Dim UserAttackInteractionResult As t_AttackInteractionResult
        UserAttackInteractionResult = UserCanAttackNpc(UserIndex, MapData(TargetPos.Map, TargetPos.x, TargetPos.y).NpcIndex)
        Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResult.Result)
        If UserAttackInteractionResult.CanAttack Then
            If UserAttackInteractionResult.TurnPK Then Call VolverCriminal(UserIndex)
            Call ThrowProjectileToTarget(UserIndex, MapData(TargetPos.Map, TargetPos.x, TargetPos.y).NpcIndex, eNpc)
            ThrowArrowToTile = True
        Else
            Exit Function
        End If
    End If
    Exit Function
ThrowArrowToTile_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.ThrowArrowToTile", Erl)
End Function

Public Sub ThrowProjectileToTarget(ByVal UserIndex As Integer, ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType)
    Dim WeaponData          As t_ObjData
    Dim ProjectileType      As Byte
    Dim AmunitionState      As Integer
    Dim DidConsumeAmunition As Boolean
    With UserList(UserIndex).invent
        If .EquippedWeaponObjIndex < 1 Then Exit Sub
        WeaponData = ObjData(.EquippedWeaponObjIndex)
        ProjectileType = GetProjectileView(UserList(UserIndex))
        If WeaponData.Proyectil = 1 And WeaponData.Municion = 0 Then
            AmunitionState = 0
        ElseIf .EquippedWeaponObjIndex = 0 Then
            AmunitionState = 1
        ElseIf .EquippedWeaponSlot < 1 Or .EquippedWeaponSlot > UserList(UserIndex).CurrentInventorySlots Then
            AmunitionState = 1
        ElseIf .EquippedMunitionSlot < 1 Or .EquippedMunitionSlot > UserList(UserIndex).CurrentInventorySlots Then
            AmunitionState = 1
        ElseIf .EquippedMunitionObjIndex = 0 Then
            AmunitionState = 1
        ElseIf ObjData(.EquippedWeaponObjIndex).Proyectil <> 1 Then
            AmunitionState = 2
        ElseIf ObjData(.EquippedMunitionObjIndex).OBJType <> e_OBJType.otArrows Then
            AmunitionState = 1
        ElseIf .Object(.EquippedMunitionSlot).amount < 1 Then
            AmunitionState = 1
        End If
        If AmunitionState <> 0 Then
            If AmunitionState = 1 Then
                ' Msg709=No tenés municiones.
                Call WriteLocaleMsg(UserIndex, 709, e_FontTypeNames.FONTTYPE_INFO)
            End If
            Call Desequipar(UserIndex, .EquippedMunitionSlot)
            Call WriteWorkRequestTarget(UserIndex, 0)
            Exit Sub
        End If
        If TargetType = eUser Then
            Dim backup    As Byte
            Dim envie     As Boolean
            Dim Particula As Integer
            Dim Tiempo    As Long
            ' Porque no es HandleAttack ???
            Call UsuarioAtacaUsuario(UserIndex, TargetIndex, Ranged)
            Dim FX As Integer
            If .EquippedMunitionObjIndex Then
                FX = ObjData(.EquippedMunitionObjIndex).CreaFX
            End If
            If FX <> 0 Then
                UserList(TargetIndex).Counters.timeFx = 3
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(UserList(TargetIndex).Char.charindex, FX, 0, UserList(TargetIndex).pos.x, UserList( _
                        TargetIndex).pos.y))
            End If
            If ProjectileType > 0 And UserList(UserIndex).flags.Oculto = 0 Then
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, UserList(TargetIndex).pos.x, _
                        UserList(TargetIndex).pos.y, ProjectileType))
            End If
            'Si no es GM invisible, le envio el movimiento del arma.
            If UserList(UserIndex).flags.AdminInvisible = 0 Then
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageArmaMov(UserList(UserIndex).Char.charindex, 1))
            End If
            If .EquippedMunitionObjIndex > 0 Then
                If ObjData(.EquippedMunitionObjIndex).CreaParticula <> "" Then
                    Particula = val(ReadField(1, ObjData(.EquippedMunitionObjIndex).CreaParticula, Asc(":")))
                    Tiempo = val(ReadField(2, ObjData(.EquippedMunitionObjIndex).CreaParticula, Asc(":")))
                    UserList(TargetIndex).Counters.timeFx = 3
                    Call SendData(SendTarget.ToPCAliveArea, TargetIndex, PrepareMessageParticleFX(UserList(TargetIndex).Char.charindex, Particula, Tiempo, False, , UserList( _
                            TargetIndex).pos.x, UserList(TargetIndex).pos.y))
                End If
            End If
            DidConsumeAmunition = True
        Else
            Call UsuarioAtacaNpc(UserIndex, TargetIndex, Ranged)
            DidConsumeAmunition = True
            If ProjectileType > 0 And UserList(UserIndex).flags.Oculto = 0 Then
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, NpcList(TargetIndex).pos.x, _
                        NpcList(TargetIndex).pos.y, ProjectileType))
            End If
            'Si no es GM invisible, le envio el movimiento del arma.
            If UserList(UserIndex).flags.AdminInvisible = 0 Then
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageArmaMov(UserList(UserIndex).Char.charindex, 1))
            End If
        End If
    End With
    If DidConsumeAmunition And Not IsConsumableFreeZone(UserIndex) Then
        Call ConsumeAmunition(UserIndex)
    End If
End Sub

Public Function GetProjectileView(ByRef User As t_User) As Integer
    Dim WeaponData     As t_ObjData
    Dim ProjectileType As Byte
    With User.invent
        If .EquippedWeaponObjIndex < 1 Then Exit Function
        WeaponData = ObjData(.EquippedWeaponObjIndex)
        If WeaponData.Proyectil = 1 And WeaponData.Municion = 0 Then
            GetProjectileView = WeaponData.ProjectileType
        ElseIf .EquippedMunitionObjIndex > 0 Then
            GetProjectileView = ObjData(.EquippedMunitionObjIndex).ProjectileType
        End If
    End With
End Function

Public Sub ConsumeAmunition(ByVal UserIndex As Integer)
    With UserList(UserIndex).invent
        Dim AmunitionSlot As Integer
        AmunitionSlot = .EquippedMunitionSlot
        If AmunitionSlot > 0 Then
            Call QuitarUserInvItem(UserIndex, AmunitionSlot, 1)
            If .Object(AmunitionSlot).amount > 0 Then
                'QuitarUserInvItem unequipps the ammo, so we equip it again
                .EquippedMunitionSlot = AmunitionSlot
                .EquippedMunitionObjIndex = .Object(AmunitionSlot).ObjIndex
                .Object(AmunitionSlot).Equipped = 1
            Else
                .EquippedMunitionSlot = 0
                .EquippedMunitionObjIndex = 0
            End If
            Call UpdateUserInv(False, UserIndex, AmunitionSlot)
        End If
    End With
End Sub

Public Sub CalculateElementalTagsModifiers(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByRef DmgAcumulator As Long)
    Dim attackerElementMask As Long
    Dim defenderElementMask As Long
    Dim attackerBit         As Long
    Dim defenderBit         As Long
    Dim attackerIndex       As Long
    Dim defenderIndex       As Long
    ' Get the bitmask of elements from the equipped weapon if the weapon naturally has tags OR if the user has enchanted his own weapon
    With UserList(UserIndex).invent
        If .EquippedWeaponObjIndex = 0 Or .EquippedWeaponSlot = 0 Then
            Exit Sub
        End If
        attackerElementMask = ObjData(.EquippedWeaponObjIndex).ElementalTags Or .Object(.EquippedWeaponSlot).ElementalTags
    End With
    ' Get the bitmask of elements from the NPC
    defenderElementMask = NpcList(NpcIndex).flags.ElementalTags
    If attackerElementMask = 0 Or defenderElementMask = 0 Then
        ' No elemental tags to process
        Exit Sub
    End If
    ' Loop over each possible attacker element (0 to 31)
    For attackerIndex = 0 To MAX_ELEMENT_TAGS - 1
        ' Create a bitmask for the current attacker element
        attackerBit = ShiftLeft(1, attackerIndex)
        ' Ensure shift is valid and safe (only 0 to 31)
        If attackerBit <> 0 And IsSet(attackerElementMask, attackerBit) Then
            ' Loop over each possible defender element
            For defenderIndex = 0 To MAX_ELEMENT_TAGS - 1
                defenderBit = ShiftLeft(1, defenderIndex)
                If defenderBit <> 0 And IsSet(defenderElementMask, defenderBit) Then
                    ' Multiply the accumulated damage by the matrix value
                    ' Matrix is 1-based, so we add 1 to both indexes
                    DmgAcumulator = DmgAcumulator * ElementalMatrixForNpcs(attackerIndex + 1, defenderIndex + 1)
                End If
            Next defenderIndex
        End If
    Next attackerIndex
End Sub
Public Function GetStabbingChanceBase(ByVal UserIndex As Integer) As Single
    On Error GoTo GetStabbingChanceBase_Err:
    Dim skill As Integer
    With UserList(UserIndex)
        skill = .Stats.UserSkills(e_Skill.Apuñalar)
        Select Case .clase
            Case e_Class.Assasin
                GetStabbingChanceBase = skill * AssasinStabbingChance
            Case e_Class.Bard
                GetStabbingChanceBase = skill * BardStabbingChance
            Case e_Class.Hunter
                GetStabbingChanceBase = skill * HunterStabbingChance
            Case Else
                GetStabbingChanceBase = skill * GenericStabbingChance
        End Select
    End With
    GetStabbingChanceBase = ClampChance(GetStabbingChanceBase)
    Exit Function
GetStabbingChanceBase_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.GetStabbingChanceBase", Erl)
End Function

Private Function GetBackHitBonusChanceAgainstUsers(ByVal UserIndex As Integer, ByVal targetUserIndex As Integer) As Single
    On Error GoTo GetBackHitBonusChanceAgainstUsers_Err:
    If UserList(UserIndex).Char.Heading = UserList(targetUserIndex).Char.Heading And Distancia(UserList(UserIndex).pos, UserList(targetUserIndex).pos) <= 1 Then
        GetBackHitBonusChanceAgainstUsers = ExtraBackstabChance
    End If
    Exit Function
GetBackHitBonusChanceAgainstUsers_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.GetBackHitBonusChanceAgainstUsers", Erl)
End Function

Private Function GetCriticalHitChanceBase(ByVal UserIndex As Integer) As Single
    On Error GoTo GetCriticalHitChanceBase_Err:
    Dim skill As Integer
    With UserList(UserIndex)
        skill = .Stats.UserSkills(e_Skill.Wrestling)
        GetCriticalHitChanceBase = skill * BanditCriticalHitChance
    End With
    GetCriticalHitChanceBase = ClampChance(GetCriticalHitChanceBase)
    Exit Function
GetCriticalHitChanceBase_Err:
    Call TraceError(Err.Number, Err.Description, "SistemaCombate.GetCriticalHitChanceBase", Erl)
End Function

Private Function GetStabbingChanceAgainstUsers(ByVal UserIndex As Integer, ByVal targetUserIndex As Integer) As Single
    GetStabbingChanceAgainstUsers = ClampChance(GetStabbingChanceBase(UserIndex) + GetBackHitBonusChanceAgainstUsers(UserIndex, targetUserIndex))
End Function

Private Function GetCriticalHitChanceAgainstUsers(ByVal UserIndex As Integer, ByVal targetUserIndex As Integer) As Single
    GetCriticalHitChanceAgainstUsers = ClampChance(GetCriticalHitChanceBase(UserIndex) + GetBackHitBonusChanceAgainstUsers(UserIndex, targetUserIndex))
End Function

