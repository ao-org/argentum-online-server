Attribute VB_Name = "Fixtures"
' Deterministic fixtures only; never linked into Server.VBP.
Option Explicit
Public Type FixtureFlags
    UsingItemSlot As Integer
    UsingItemObjIndex As Integer
    Muerto As Byte
    AfectaParalisis As Byte
End Type
Public Type t_UserOBJ
    ObjIndex As Integer
    amount As Integer
End Type
Public Type FixtureInventory
    Object(1 To 10) As t_UserOBJ
End Type
Public Type FixtureStats
    MinHp As Integer
    MinSta As Integer
    tipoUsuario As Integer
End Type
Public Type FixtureChar
    charindex As Integer
End Type
Public Type t_User
    flags As FixtureFlags
    pos As t_WorldPos
    invent As FixtureInventory
    Stats As FixtureStats
    Char As FixtureChar
    EffectOverTime As t_EffectOverTimeList
    CurrentInventorySlots As Integer
    ChatCombate As Byte
    VersionId As Integer
    name As String
End Type
Public Type FixtureNpc
    flags As FixtureFlags
    Char As FixtureChar
    EffectOverTime As t_EffectOverTimeList
End Type
Public Type FixtureTile
    Trap As clsTrap
    Blocked As Byte
    NpcIndex As Integer
    UserIndex As Integer
    ObjInfo As t_UserOBJ
End Type
Public Type FixtureEffect
    Type As e_EffectOverTimeType
    SubType As Integer
    TickTime As Long
    ClientEffectTypeId As Integer
    buffType As e_EffectType
    SharedTypeId As Integer
    OnHitWav As Integer
    OnHitFx As Integer
    ApplyEffectId As Integer
    TickPowerMin As Integer
    TickPowerMax As Integer
End Type
Public Type FixtureObject
    name As String
    OBJType As e_OBJType
    Subtipo As Integer
    EfectoMagico As Integer
    cdType As Integer
    MinHp As Integer
    MinSta As Integer
End Type
Public Type FixtureAttack
    CanAttack As Boolean
End Type
Public UserList(1 To 2) As t_User
Public NpcList(1 To 2) As FixtureNpc
Public MapData(1 To 100, 1 To 100, 1 To 2) As FixtureTile
Public EffectOverTime(1 To 9) As FixtureEffect
Public ObjData(1 To 5) As FixtureObject
Public Writer As Network.Writer
Public Packets As Collection
Public SvrConfig As TrapConfig
Public AllowAttack As Boolean
Public Cooldown As Boolean
Public SoundCalls As Long, FxCalls As Long, DamageCalls As Long
Public ParalysisCalls As Long, PoisonCalls As Long, Consumed As Long
Public CdCalls As Long, CdClears As Long, Recycled As Long
Public Const MsgFallIntoTrap As Integer = 1
Public Const MsgTrapInmo As Integer = 2
Public Const MsgYourTrapDidDamangeTo As Integer = 3
Public Const MsgTrapDidDamageToYou As Integer = 4
Public Const MsgTrapPoison As Integer = 5
Public Const MsgRequiresMoreHealth As Integer = 6
Public Const MSG_NO_ENERGY As Integer = 7
Public Const MsgTiredToPerformAction As Integer = 8
Public Const MSG_AT_LEAST_NOBLE_TO_FOUND_CASTLE As Integer = 9
Public Const MsgToFar As Integer = 10
Public Const MsgInvalidTile As Integer = 11

Public Sub ResetFixture()
    Dim x As Integer, y As Integer, Map As Integer, i As Integer
    For Map = 1 To 2
        For x = 1 To 100
            For y = 1 To 100
                Set MapData(x, y, Map).Trap = Nothing
                MapData(x, y, Map).Blocked = 0
            Next y
        Next x
    Next Map
    Erase UserList
    Erase NpcList
    Erase ObjData
    Erase EffectOverTime
    Call ResetEffects()
    Set Writer = New Network.Writer
    Set Packets = New Collection
    Set SvrConfig = New TrapConfig
    AllowAttack = True: Cooldown = False
    SoundCalls = 0: FxCalls = 0: DamageCalls = 0: ParalysisCalls = 0
    PoisonCalls = 0: Consumed = 0: CdCalls = 0: CdClears = 0: Recycled = 0
    With UserList(1)
        .VersionId = 1
        .CurrentInventorySlots = 10
        .flags.UsingItemSlot = 1
        .flags.UsingItemObjIndex = 1
        .invent.Object(1).ObjIndex = 1
        .invent.Object(1).amount = 20
        .pos.Map = 1: .pos.x = 10: .pos.y = 10
        .Stats.MinHp = 100: .Stats.MinSta = 100
    End With
    ObjData(1).OBJType = e_OBJType.otUsableOntarget
    ObjData(1).Subtipo = e_UssableOnTarget.eTrap
    ObjData(1).EfectoMagico = 6
    For i = 6 To 9
        EffectOverTime(i).Type = e_EffectOverTimeType.eTrap
        EffectOverTime(i).SubType = 1
        EffectOverTime(i).TickTime = 10000
        EffectOverTime(i).ClientEffectTypeId = 5
        EffectOverTime(i).OnHitFx = 8
        EffectOverTime(i).buffType = eInformativeBuff
    Next i
    EffectOverTime(7).SubType = 2: EffectOverTime(7).TickTime = 120000
    EffectOverTime(7).OnHitWav = 162
    EffectOverTime(7).TickPowerMin = -100: EffectOverTime(7).TickPowerMax = -150
    EffectOverTime(8).Type = eHealthModifier
    EffectOverTime(9).SubType = 3: EffectOverTime(9).TickTime = 120000
    EffectOverTime(9).ApplyEffectId = 8: EffectOverTime(9).OnHitFx = 0
End Sub

Public Sub SetRef(ByRef Ref As t_AnyReference, ByVal Index As Integer, ByVal Kind As e_ReferenceType)
    Ref.ArrayIndex = Index: Ref.RefType = Kind
    Ref.VersionId = UserList(Index).VersionId
End Sub
Public Function IsValidRef(ByRef Ref As t_AnyReference) As Boolean
    If Ref.RefType = eNone Then Exit Function
    If Ref.ArrayIndex < 1 Or Ref.ArrayIndex > 2 Then Exit Function
    IsValidRef = Ref.VersionId = UserList(Ref.ArrayIndex).VersionId
End Function
Public Function UserCanAttackNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As FixtureAttack
    UserCanAttackNpc.CanAttack = AllowAttack And UserList(UserIndex).flags.Muerto = 0
End Function
Public Function IsValidMapPosition(ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
    IsValidMapPosition = Map >= 1 And Map <= 2 And x >= 1 And x <= 100 And y >= 1 And y <= 100
End Function
Public Function IsItemInCooldown(ByRef User As t_User, ByRef obj As t_UserOBJ) As Boolean
    IsItemInCooldown = Cooldown
End Function
Public Function Distance(ByVal x As Integer, ByVal y As Integer, ByVal x2 As Integer, ByVal y2 As Integer) As Double
    Distance = Sqr((x - x2) ^ 2 + (y - y2) ^ 2)
End Function
Public Function IsArrayInitialized(ByRef Items() As IBaseEffectOverTime) As Boolean
    On Error GoTo EmptyArray
    IsArrayInitialized = UBound(Items) >= LBound(Items)
EmptyArray:
End Function
Public Sub SetMask(ByRef Mask As Long, ByVal Bits As Long)
    Mask = Mask Or Bits
End Sub
Public Sub RemoveItemFromInventory(ByVal UserIndex As Integer, ByVal Slot As Integer)
    UserList(UserIndex).invent.Object(Slot).amount = UserList(UserIndex).invent.Object(Slot).amount - 1
    Consumed = Consumed + 1
End Sub
Public Sub UpdateCd(ByVal UserIndex As Integer, ByVal Kind As Integer)
    CdCalls = CdCalls + 1
End Sub
Public Sub WriteSendSkillCdUpdate(ByVal UserIndex As Integer, ByVal Kind As Integer, ByVal Id As Long, ByVal Total As Long, ByVal Remaining As Long, ByVal Buff As Integer)
    If Total = 0 Then CdClears = CdClears + 1
End Sub
Public Function PrepareMessagePlayWave(ByVal Wave As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal LocalSound As Boolean)
    SoundCalls = SoundCalls + 1
    Call Writer.WriteInt16(1)
End Function
Public Function PrepareMessageCreateFX(ByVal Character As Integer, ByVal Fx As Integer, ByVal Loops As Integer, ByVal x As Integer, ByVal y As Integer)
    FxCalls = FxCalls + 1
    Call Writer.WriteInt16(1)
End Function
Public Function PrepareLocalizedChatOverHead(ByVal Msg As Integer, ByVal Character As Integer, ByVal Color As Long)
    Call Writer.WriteInt16(1)
End Function
Public Function RandomNumber(ByVal Minimum As Integer, ByVal Maximum As Integer) As Integer
    RandomNumber = Minimum
End Function
Public Sub WriteLocaleMsg(ByVal UserIndex As Integer, ByVal Msg As Integer, ByVal Channel As Integer, ByVal Font As Integer, Optional ByVal Text As String)
End Sub
Public Function GetName(ByRef Ref As t_AnyReference) As String
    GetName = "fixture"
End Function
Public Sub RegisterNewAttack(ByVal Target As Integer, ByVal Source As Integer)
    Err.Raise 5, , "PvP must remain unreachable"
End Sub
Public Function IsValidCastlePosition(ByVal UserIndex As Integer) As Boolean
End Function
Public Sub CreateNewEmperorCastle(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)
    Err.Raise 5, , "Unexpected castle dispatch"
End Sub
Public Sub LogInfoServidor(ByVal Text As String)
End Sub
Public Sub ResurrectWithItem(ByVal UserIndex As Integer)
    Err.Raise 5, , "Unexpected resurrection dispatch"
End Sub
Public Sub UseArpon(ByVal UserIndex As Integer)
    Err.Raise 5, , "Unexpected arpon dispatch"
End Sub
Public Sub UseHandCannon(ByVal UserIndex As Integer, ByVal x As Integer, ByVal y As Integer)
    Err.Raise 5, , "Unexpected cannon dispatch"
End Sub
Public Sub TraceError(ByVal Number As Long, ByVal Description As String, ByVal Source As String, ByVal Line As Long)
    Err.Raise Number, Source, Description
End Sub
