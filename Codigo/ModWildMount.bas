Attribute VB_Name = "ModWildMount"
Option Explicit

Private Const WildMountTameOddsDivisor As Byte = 5

Private Type t_WildMountConfig
    NpcNumber As Integer
    TamedObjCount As Integer
    TamedObjIndexes() As Integer
    TamedObjWeights() As Integer
    TamedObjWeightTotal As Long
    MapPoolCount As Integer
    MapPool() As Integer
    IntervaloMinSpawnSeg As Long
    IntervaloMaxSpawnSeg As Long
    MaxLevel As Byte
    ExpPerLevel As Long
    BonusPerLevel As Single
    ExpStealBasePercent As Single
    ExpStealPercentPerLevel As Single
    ExpStealMaxPercent As Single
    EnableSpawn As Boolean
    MinTameChance As Single
    MaxTameChance As Single
End Type

Private Type t_WildMountState
    npc As t_NpcReference
    CounterSeg As Long
    NextSpawnSeg As Long
End Type

Private WildMountConfigs()     As t_WildMountConfig
Private WildMountStates()      As t_WildMountState
Private WildMountCount         As Integer
Private WildMountConfigLoaded  As Boolean

Public Sub LoadWildMountConfig()
    On Error GoTo LoadWildMountConfig_Err
    Dim Ini As New clsIniManager
    Call Ini.Initialize(DatPath & "WildMount.dat")
    WildMountCount = val(Ini.GetValue("INIT", "NumMounts"))
    If WildMountCount <= 0 Then
        Call LogError("WildMount.dat sin monturas configuradas (NumMounts=0), sistema deshabilitado.")
        WildMountConfigLoaded = False
        Set Ini = Nothing
        Exit Sub
    End If
    ReDim WildMountConfigs(1 To WildMountCount)
    ReDim WildMountStates(1 To WildMountCount)
    Dim i As Long, SectionName As String
    For i = 1 To WildMountCount
        SectionName = "MOUNT" & i
        With WildMountConfigs(i)
            .NpcNumber = val(Ini.GetValue(SectionName, "NpcNumber"))
            .TamedObjCount = val(Ini.GetValue(SectionName, "TamedObjCount"))
            If .TamedObjCount > 0 Then
                ReDim .TamedObjIndexes(1 To .TamedObjCount)
                ReDim .TamedObjWeights(1 To .TamedObjCount)
                .TamedObjWeightTotal = 0
                Dim k As Long
                For k = 1 To .TamedObjCount
                    .TamedObjIndexes(k) = val(Ini.GetValue(SectionName, "TamedObj" & k))
                    .TamedObjWeights(k) = val(Ini.GetValue(SectionName, "TamedObjWeight" & k))
                    .TamedObjWeightTotal = .TamedObjWeightTotal + .TamedObjWeights(k)
                Next k
            End If
            .IntervaloMinSpawnSeg = val(Ini.GetValue(SectionName, "IntervaloMinSpawnMinutos")) * 60
            .IntervaloMaxSpawnSeg = val(Ini.GetValue(SectionName, "IntervaloMaxSpawnMinutos")) * 60
            .MapPoolCount = val(Ini.GetValue(SectionName, "MapPoolCount"))
            If .MapPoolCount > 0 Then
                ReDim .MapPool(1 To .MapPoolCount)
                Dim j As Long
                For j = 1 To .MapPoolCount
                    .MapPool(j) = val(Ini.GetValue(SectionName, "Map" & j))
                Next j
            End If
            .MaxLevel = val(Ini.GetValue(SectionName, "MaxLevel"))
            .ExpPerLevel = val(Ini.GetValue(SectionName, "ExpPerLevel"))
            .BonusPerLevel = val(Ini.GetValue(SectionName, "BonusPerLevel"))
            .ExpStealBasePercent = val(Ini.GetValue(SectionName, "ExpStealBasePercent"))
            .ExpStealPercentPerLevel = val(Ini.GetValue(SectionName, "ExpStealPercentPerLevel"))
            .ExpStealMaxPercent = val(Ini.GetValue(SectionName, "ExpStealMaxPercent"))
            If .ExpStealMaxPercent <= 0 Then .ExpStealMaxPercent = 1 ' Sin techo configurado, no limita (cuidado)
            .EnableSpawn = (val(Ini.GetValue(SectionName, "EnableSpawn")) <> 0)
            If .MaxLevel <= 0 Then .MaxLevel = 1 ' Sin config válida, la montura no sube de nivel
            .MinTameChance = val(Ini.GetValue(SectionName, "MinTameChance"))
            .MaxTameChance = val(Ini.GetValue(SectionName, "MaxTameChance"))
            If .MaxTameChance <= 0 Then .MaxTameChance = 1
        End With
                If WildMountConfigs(i).NpcNumber <= 0 Or WildMountConfigs(i).MapPoolCount <= 0 Or WildMountConfigs(i).TamedObjCount <= 0 Or WildMountConfigs(i).TamedObjWeightTotal <= 0 Then
            Call LogError("WildMount.dat: " & SectionName & " mal configurado, se ignora esa entrada.")
            WildMountConfigs(i).NpcNumber = 0 ' Marca la entrada como inválida
        End If
        WildMountStates(i).NextSpawnSeg = RandomNumber(WildMountConfigs(i).IntervaloMinSpawnSeg, WildMountConfigs(i).IntervaloMaxSpawnSeg)
        WildMountStates(i).CounterSeg = 0
        Call ClearNpcRef(WildMountStates(i).npc)
    Next i
    Set Ini = Nothing
    WildMountConfigLoaded = True
    Exit Sub
LoadWildMountConfig_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.LoadWildMountConfig", Erl)
End Sub

Public Sub CheckWildMountSpawn()
    On Error GoTo CheckWildMountSpawn_Err
    If Not WildMountConfigLoaded Then Exit Sub
    Dim i As Long
    For i = 1 To WildMountCount
        If WildMountConfigs(i).NpcNumber > 0 Then
            Call CheckSingleWildMountSpawn(i)
        End If
    Next i
    Exit Sub
CheckWildMountSpawn_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.CheckWildMountSpawn", Erl)
End Sub

Private Sub CheckSingleWildMountSpawn(ByVal ConfigIndex As Integer)
    On Error GoTo CheckSingleWildMountSpawn_Err
    If Not WildMountConfigs(ConfigIndex).EnableSpawn Then Exit Sub
    With WildMountStates(ConfigIndex)
        If IsValidNpcRef(.npc) Then
            If NpcList(.npc.ArrayIndex).flags.NPCActive Then Exit Sub
        End If
        If .CounterSeg < .NextSpawnSeg Then
            .CounterSeg = .CounterSeg + 1
            Exit Sub
        End If
        Call SpawnWildMount(ConfigIndex)
        .CounterSeg = 0
        .NextSpawnSeg = RandomNumber(WildMountConfigs(ConfigIndex).IntervaloMinSpawnSeg, WildMountConfigs(ConfigIndex).IntervaloMaxSpawnSeg)
    End With
    Exit Sub
CheckSingleWildMountSpawn_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.CheckSingleWildMountSpawn", Erl)
End Sub

Private Sub SpawnWildMount(ByVal ConfigIndex As Integer)
    On Error GoTo SpawnWildMount_Err
    Dim MapIndex  As Integer
    Dim OrigPos   As t_WorldPos
    Dim NewNpcIdx As Integer
    With WildMountConfigs(ConfigIndex)
        MapIndex = .MapPool(RandomNumber(1, .MapPoolCount))
        OrigPos.Map = MapIndex
        NewNpcIdx = CrearNPC(.NpcNumber, MapIndex, OrigPos)
    End With
    If NewNpcIdx <> 0 Then
        Call SetNpcRef(WildMountStates(ConfigIndex).npc, NewNpcIdx)
    Else
        Call ClearNpcRef(WildMountStates(ConfigIndex).npc)
    End If
    Exit Sub
SpawnWildMount_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.SpawnWildMount", Erl)
End Sub

' Devuelve el índice de configuración (1..WildMountCount) si el NPC es una montura salvaje, o 0 si no lo es.
Private Function FindWildMountConfigIndex(ByVal npcIndex As Integer) As Integer
    If Not WildMountConfigLoaded Then Exit Function
    Dim i As Long
    Dim NpcNumero As Integer
    NpcNumero = NpcList(npcIndex).Numero
    For i = 1 To WildMountCount
        If WildMountConfigs(i).NpcNumber > 0 And WildMountConfigs(i).NpcNumber = NpcNumero Then
            FindWildMountConfigIndex = i
            Exit Function
        End If
    Next i
End Function

Public Function IsWildMountNpc(ByVal npcIndex As Integer) As Boolean
    IsWildMountNpc = FindWildMountConfigIndex(npcIndex) > 0
End Function

Public Sub DoTameWildMount(ByVal UserIndex As Integer, ByVal npcIndex As Integer)
    On Error GoTo DoTameWildMount_Err
    Dim ConfigIndex As Integer
    Dim puntosDomar As Long
    ConfigIndex = FindWildMountConfigIndex(npcIndex)
    If ConfigIndex = 0 Then Exit Sub
    With UserList(UserIndex)
        If .flags.Privilegios And e_PlayerType.Consejero Then Exit Sub
        If .flags.Muerto = 1 Then Exit Sub
        If NpcList(npcIndex).MinTameLevel > .Stats.ELV Then
            Call WriteLocaleMsg(UserIndex, MSG_DEBES_NIVEL_SUPERIOR_DOMAR_CRIATURA, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja, NpcList(npcIndex).MinTameLevel)
            Exit Sub
        End If
        puntosDomar = CLng(.Stats.UserSkills(e_Skill.Domar))
        Dim TameChance As Single
        With WildMountConfigs(ConfigIndex)
            TameChance = .MinTameChance + (CSng(puntosDomar) / 100) * (.MaxTameChance - .MinTameChance)
        End With
        If RandomNumber(1, 10000) <= CLng(TameChance * 10000) Then
            Call OnWildMountTameSuccess(UserIndex, npcIndex, ConfigIndex)
        Else
            Call OnWildMountTameFailure(UserIndex, npcIndex)
        End If
        Call SubirSkill(UserIndex, e_Skill.Domar)
    End With
    Exit Sub
DoTameWildMount_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.DoTameWildMount", Erl)
End Sub

Private Sub OnWildMountTameFailure(ByVal UserIndex As Integer, ByVal npcIndex As Integer)
    On Error GoTo OnWildMountTameFailure_Err
    If Not UserList(UserIndex).flags.UltimoMensaje = MSG_NO_TAME_FAILED Then
        Call WriteLocaleMsg(UserIndex, MSG_NO_TAME_FAILED, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
        UserList(UserIndex).flags.UltimoMensaje = MSG_NO_TAME_FAILED
    End If
    With NpcList(npcIndex)
        If .Hostile = 0 Then
            .Hostile = 1
            .flags.OldHostil = 1
            .flags.AttackedBy = UserList(UserIndex).name
            .flags.AttackedTime = GlobalFrameTime
            Call SetUserRef(.targetUser, UserIndex)
            Call SetMovement(npcIndex, e_TipoAI.NpcDefensa)
        End If
    End With
    Exit Sub
OnWildMountTameFailure_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.OnWildMountTameFailure", Erl)
End Sub

Private Sub OnWildMountTameSuccess(ByVal UserIndex As Integer, ByVal npcIndex As Integer, ByVal ConfigIndex As Integer)
    On Error GoTo OnWildMountTameSuccess_Err
    Dim MiObj As t_Obj
    MiObj.ObjIndex = PickRandomTamedObj(ConfigIndex)
    MiObj.amount = 1
    If Not HayLugarEnInventario(UserIndex, MiObj.ObjIndex, 1) Then
        Call WriteLocaleMsg(UserIndex, MSG_WILD_MOUNT_ESCAPED_NO_SPACE, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Naranja)
        Call QuitarNPC(npcIndex, eTame)
        Exit Sub
    End If
    Call WriteLocaleMsg(UserIndex, MSG_TAMED_WILD_MOUNT_SUCCESS, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_PROMEDIO_MAYOR)
    Call QuitarNPC(npcIndex, eTame)
    Call MeterItemEnInventario(UserIndex, MiObj)
    Call InitializeMountLevelInSlot(UserIndex, MiObj.ObjIndex)
    Exit Sub
OnWildMountTameSuccess_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.OnWildMountTameSuccess", Erl)
End Sub

Private Sub InitializeMountLevelInSlot(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    On Error GoTo InitializeMountLevelInSlot_Err
    Dim i As Integer
    For i = 1 To UBound(UserList(UserIndex).invent.Object)
        If UserList(UserIndex).invent.Object(i).ObjIndex = ObjIndex Then
            If UserList(UserIndex).invent.Object(i).MountLevel = 0 Then
                UserList(UserIndex).invent.Object(i).MountLevel = 1
                UserList(UserIndex).invent.Object(i).MountExp = 0
            End If
            Exit For
        End If
    Next i
    Exit Sub
InitializeMountLevelInSlot_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.InitializeMountLevelInSlot", Erl)
End Sub

Public Function GetWildMountExpToNextLevel(ByVal ConfigIndex As Integer, ByVal CurrentLevel As Byte) As Long
    On Error GoTo GetWildMountExpToNextLevel_Err
    GetWildMountExpToNextLevel = CLng(WildMountConfigs(ConfigIndex).ExpPerLevel) * CLng(CurrentLevel)
    Exit Function
GetWildMountExpToNextLevel_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.GetWildMountExpToNextLevel", Erl)
End Function

Public Sub GrantWildMountExp(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal ExpAmount As Long)
    On Error GoTo GrantWildMountExp_Err
    With UserList(UserIndex).invent.Object(Slot)
        Dim ConfigIndex As Integer
        ConfigIndex = FindWildMountConfigIndexByObjIndex(.ObjIndex)
        If ConfigIndex = 0 Then Exit Sub
        If .MountLevel >= WildMountConfigs(ConfigIndex).MaxLevel Then Exit Sub ' Ya está al tope
        .MountExp = .MountExp + ExpAmount
        Dim ExpNeeded As Long
        ExpNeeded = GetWildMountExpToNextLevel(ConfigIndex, .MountLevel)
        Do While .MountExp >= ExpNeeded And .MountLevel < WildMountConfigs(ConfigIndex).MaxLevel
            .MountExp = .MountExp - ExpNeeded
            .MountLevel = .MountLevel + 1
            Dim strExtra As String
            strExtra = ObjData(.ObjIndex).name & Chr$(172) & .MountLevel
            Call WriteLocaleMsg(UserIndex, MSG_WILD_MOUNT_LEVEL_UP, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_PROMEDIO_MAYOR, strExtra)
            UserList(UserIndex).Counters.timeFx = 3
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.charindex, 106, 0, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
            If .MountLevel >= WildMountConfigs(ConfigIndex).MaxLevel Then
                .MountExp = 0 ' Tope alcanzado, no acumula más
                Exit Do
            End If
            ExpNeeded = GetWildMountExpToNextLevel(ConfigIndex, .MountLevel)
        Loop
    End With
    Exit Sub
GrantWildMountExp_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.GrantWildMountExp", Erl)
End Sub

Private Function FindWildMountConfigIndexByObjIndex(ByVal ObjIndex As Integer) As Integer
    If Not WildMountConfigLoaded Then Exit Function
    Dim i As Long, k As Long
    For i = 1 To WildMountCount
        For k = 1 To WildMountConfigs(i).TamedObjCount
            If WildMountConfigs(i).TamedObjIndexes(k) = ObjIndex Then
                FindWildMountConfigIndexByObjIndex = i
                Exit Function
            End If
        Next k
    Next i
End Function

Public Function GetWildMountEffectiveBonus(ByVal BaseBonus As Single, ByVal MountLevel As Byte, ByVal ObjIndexEquipped As Integer) As Single
    On Error GoTo GetWildMountEffectiveBonus_Err
    Dim ConfigIndex As Integer
    ConfigIndex = FindWildMountConfigIndexByObjIndex(ObjIndexEquipped)
    If ConfigIndex = 0 Then
        GetWildMountEffectiveBonus = BaseBonus
        Exit Function
    End If
    GetWildMountEffectiveBonus = BaseBonus + (CSng(MountLevel) - 1) * WildMountConfigs(ConfigIndex).BonusPerLevel
    Exit Function
GetWildMountEffectiveBonus_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.GetWildMountEffectiveBonus", Erl)
End Function

Public Sub ShowMountStatusMessage(ByVal UserIndex As Integer, ByVal Slot As Byte)
    On Error GoTo ShowMountStatusMessage_Err
    Dim ConfigIndex As Integer
    Dim CurrentLevel As Byte
    Dim CurrentExp As Long
    Dim ExpNeeded As Long
    Dim strExtra As String
    Dim ObjIndexEquipped As Integer
    With UserList(UserIndex).invent.Object(Slot)
        ConfigIndex = FindWildMountConfigIndexByObjIndex(.ObjIndex)
        If ConfigIndex = 0 Then Exit Sub
        CurrentLevel = .MountLevel
        CurrentExp = .MountExp
        ObjIndexEquipped = .ObjIndex
    End With
    Dim MsgActive As Integer
    Dim MsgMaxLevel As Integer
    Select Case ObjData(ObjIndexEquipped).NpcDamageBonusCategory
        Case e_WildMountBonusCategory.eWildMountMagic
            MsgActive = MSG_WILD_MOUNT_BONUS_ACTIVE_MAGIC
            MsgMaxLevel = MSG_WILD_MOUNT_BONUS_ACTIVE_MAGIC_MAXLEVEL
        Case e_WildMountBonusCategory.eWildMountWeapon
            MsgActive = MSG_WILD_MOUNT_BONUS_ACTIVE_WEAPON
            MsgMaxLevel = MSG_WILD_MOUNT_BONUS_ACTIVE_WEAPON_MAXLEVEL
        Case e_WildMountBonusCategory.eWildMountKnuckle
            MsgActive = MSG_WILD_MOUNT_BONUS_ACTIVE_KNUCKLE
            MsgMaxLevel = MSG_WILD_MOUNT_BONUS_ACTIVE_KNUCKLE_MAXLEVEL
        Case e_WildMountBonusCategory.eWildMountDagger
            MsgActive = MSG_WILD_MOUNT_BONUS_ACTIVE_DAGGER
            MsgMaxLevel = MSG_WILD_MOUNT_BONUS_ACTIVE_DAGGER_MAXLEVEL
        Case e_WildMountBonusCategory.eWildMountBow
            MsgActive = MSG_WILD_MOUNT_BONUS_ACTIVE_BOW
            MsgMaxLevel = MSG_WILD_MOUNT_BONUS_ACTIVE_BOW_MAXLEVEL
        Case Else
            Exit Sub ' Categoría desconocida, no mostramos nada
    End Select
    If CurrentLevel >= WildMountConfigs(ConfigIndex).MaxLevel Then
        Call WriteLocaleMsg(UserIndex, MsgMaxLevel, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Blanco, CStr(CurrentLevel))
    Else
        ExpNeeded = GetWildMountExpToNextLevel(ConfigIndex, CurrentLevel)
        strExtra = CurrentLevel & Chr$(172) & CurrentExp & Chr$(172) & ExpNeeded
        Call WriteLocaleMsg(UserIndex, MsgActive, e_TextChannel.TEXTCHANNEL_SYSTEM, e_FontTypeNames.FONTTYPE_New_Blanco, strExtra)
    End If
    Exit Sub
ShowMountStatusMessage_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.ShowMountStatusMessage", Erl)
End Sub

Public Function GetWildMountExpSteal(ByVal UserIndex As Integer, ByVal ExpAmount As Double) As Double
    On Error GoTo GetWildMountExpSteal_Err
    With UserList(UserIndex)
        If .flags.Montado <> 1 Then Exit Function
        If .invent.EquippedSaddleObjIndex <= 0 Then Exit Function
        If ObjData(.invent.EquippedSaddleObjIndex).OBJType <> e_OBJType.otWildMount Then Exit Function
        Dim ConfigIndex As Integer
        ConfigIndex = FindWildMountConfigIndexByObjIndex(.invent.EquippedSaddleObjIndex)
        If ConfigIndex = 0 Then Exit Function
        Dim MountedLevel As Byte
        MountedLevel = .invent.Object(.invent.EquippedSaddleSlot).MountLevel
        If MountedLevel >= WildMountConfigs(ConfigIndex).MaxLevel Then Exit Function ' Nivel máximo, no roba más EXP
        Dim StealPercent As Single
        With WildMountConfigs(ConfigIndex)
            StealPercent = .ExpStealBasePercent + (CSng(MountedLevel) - 1) * .ExpStealPercentPerLevel
            If StealPercent > .ExpStealMaxPercent Then StealPercent = .ExpStealMaxPercent
        End With
        If StealPercent <= 0 Then Exit Function
        GetWildMountExpSteal = ExpAmount * StealPercent
    End With
    Exit Function
GetWildMountExpSteal_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.GetWildMountExpSteal", Erl)
End Function

Private Function PickRandomTamedObj(ByVal ConfigIndex As Integer) As Integer
    On Error GoTo PickRandomTamedObj_Err
    With WildMountConfigs(ConfigIndex)
        Dim Roll As Long
        Roll = RandomNumber(1, .TamedObjWeightTotal)
        Dim Accum As Long
        Dim k As Long
        Accum = 0
        For k = 1 To .TamedObjCount
            Accum = Accum + .TamedObjWeights(k)
            If Roll <= Accum Then
                PickRandomTamedObj = .TamedObjIndexes(k)
                Exit Function
            End If
        Next k
        ' Fallback de seguridad, no debería llegar acá si los pesos están bien cargados
        PickRandomTamedObj = .TamedObjIndexes(.TamedObjCount)
    End With
    Exit Function
PickRandomTamedObj_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.PickRandomTamedObj", Erl)
End Function

Public Sub RestoreMountProgressInSlot(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, ByVal MountLevel As Byte, ByVal MountExp As Long)
    On Error GoTo RestoreMountProgressInSlot_Err
    Dim i As Integer
    For i = 1 To UBound(UserList(UserIndex).invent.Object)
        If UserList(UserIndex).invent.Object(i).ObjIndex = ObjIndex And UserList(UserIndex).invent.Object(i).MountLevel = 0 Then
            UserList(UserIndex).invent.Object(i).MountLevel = MountLevel
            UserList(UserIndex).invent.Object(i).MountExp = MountExp
            Exit For
        End If
    Next i
    Exit Sub
RestoreMountProgressInSlot_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.RestoreMountProgressInSlot", Erl)
End Sub

Public Function HasMagicWeaponEquipped(ByVal UserIndex As Integer) As Boolean
    On Error GoTo HasMagicWeaponEquipped_Err
    With UserList(UserIndex).invent
        If .EquippedWeaponObjIndex > 0 Then
            If ObjData(.EquippedWeaponObjIndex).WeaponType = e_WeaponType.eStaff Then
                HasMagicWeaponEquipped = True
                Exit Function
            End If
        End If
        If .EquippedRingAccesoryObjIndex > 0 Then
            If ObjData(.EquippedRingAccesoryObjIndex).OBJType = e_OBJType.otMagicalInstrument Then
                HasMagicWeaponEquipped = True
                Exit Function
            End If
        End If
    End With
    Exit Function
HasMagicWeaponEquipped_Err:
    Call TraceError(Err.Number, Err.Description, "ModWildMount.HasMagicWeaponEquipped", Erl)
End Function

