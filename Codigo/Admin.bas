Attribute VB_Name = "Admin"
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
Public AdministratorAccounts As Dictionary

Public Type t_Motd
    texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD()   As t_Motd

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type

Public Apuestas                      As tAPuestas
Public NPCs                          As Long
Public DebugSocket                   As Boolean
Public horas                         As Long
Public dias                          As Long
Public MinsRunning                   As Long
Public ReiniciarServer               As Long
Public tInicioServer                 As Long
'INTERVALOS
Public SanaIntervaloSinDescansar     As Integer
Public StaminaIntervaloSinDescansar  As Integer
Public SanaIntervaloDescansar        As Integer
Public StaminaIntervaloDescansar     As Integer
Public IntervaloPerderStamina        As Integer
Public IntervaloSed                  As Integer
Public IntervaloHambre               As Integer
Public IntervaloVeneno               As Integer
'Ladder
Public IntervaloIncineracion         As Integer
Public IntervaloInmovilizado         As Integer
Public IntervaloMaldicion            As Integer
'Ladder
Public IntervaloParalizado           As Integer
Public IntervaloInvisible            As Integer
Public IntervaloFrio                 As Integer
Public IntervaloWavFx                As Integer
Public IntervaloNPCPuedeAtacar       As Integer
Public IntervaloNPCAI                As Integer
Public IntervaloInvocacion           As Integer
Public IntervaloOculto               As Integer '[Nacho]
Public IntervaloUserPuedeAtacar      As Long
Public IntervaloMagiaGolpe           As Long
Public IntervaloGolpeMagia           As Long
Public IntervaloUserPuedeCastear     As Long
Public IntervaloTrabajarExtraer      As Long
Public IntervaloNpcOwner             As Long
Public IntervaloTrabajarConstruir    As Long
Public IntervaloCerrarConexion       As Long '[Gonzalo]
Public IntervaloUserPuedeUsarU       As Long
Public IntervaloUserPuedeUsarClic    As Long
Public IntervaloGolpeUsar            As Long
Public IntervaloFlechasCazadores     As Long
Public TimeoutPrimerPaquete          As Long
Public TimeoutEsperandoLoggear       As Long
Public IntervaloTirar                As Long
Public IntervaloMeditar              As Long
Public IntervaloCaminar              As Long
Public IntervaloEnCombate            As Long
Public IntervaloPuedeSerAtacado      As Long
Public IntervaloGuardarUsuarios      As Long
Public LimiteGuardarUsuarios         As Integer
Public IntervaloTimerGuardarUsuarios As Long
Public IntervaloMensajeGlobal        As Long
Public Const IntervaloConsultaGM     As Long = 300000
'BALANCE
Public PorcentajeRecuperoMana        As Integer
Public RecoveryMana                  As Integer
Public MultiplierManaxSkills         As Currency
Public ManaCommonLute                As Integer
Public ManaMagicLute                 As Integer
Public ManaElvenLute                 As Integer
Public DificultadSubirSkill          As Integer
Public RequiredSpellDisplayTime      As Integer
Public MaxInvisibleSpellDisplayTime  As Integer
Public DesbalancePromedioVidas       As Single
Public RangoVidas                    As Single
Public CapVidaMax                    As Single
Public CapVidaMin                    As Single
Public ExpLevelUp(1 To STAT_MAXELV)  As Long
Public InfluenciaPromedioVidas       As Single
Public ModDañoGolpeCritico          As Single
Public MinutosWs                         As Long
Public PlayerStunTime                    As Long
Public NpcStunTime                       As Long
Public PlayerInmuneTime                  As Long
Public MultiShotReduction                As Single
Public HomeTimer                         As Integer
Public MagicSkillBonusDamageModifier     As Single
Public MRSkillProtectionModifier         As Single
Public MRSkillNpcProtectionModifier      As Single
Public AssistDamageValidTime             As Long 'valid time for damage to count as assit
Public AssistHelpValidTime               As Long 'valid time for helpful spell to count as assist
Public HideAfterHitTime                  As Long 'required time to hide again after a hit remove us from this state
Public FactionReKillTime                 As Long 'required time between killing the same user to get factions points
Public AirHitReductParalisisTime         As Integer 'you can hit to the air to reduce inmo/paralisis time
Public PorcentajePescaSegura             As Integer 'Porcentaje de reducción a la pesca en zona segura
Public DivineBloodHealingMultiplierBonus As Single
Public DivineBloodManaCostMultiplier     As Single
Public WarriorLifeStealOnHitMultiplier   As Single
Public Puerto                            As Long
Public ListenIp                          As String
Public MAXPASOS                          As Long
Public BootDelBackUp                     As Byte
Public Lloviendo                         As Boolean
Public Nebando                           As Boolean
Public Nieblando                         As Boolean
Public IpList                            As New Collection
Public Baneos                            As New Collection

Sub ReSpawnOrigPosNpcs()
    On Error GoTo Handler
    Dim i     As Integer
    Dim MiNPC As t_Npc
    For i = 1 To LastNPC
        'OJO
        If NpcList(i).flags.NPCActive Then
            If InMapBounds(NpcList(i).Orig.Map, NpcList(i).Orig.x, NpcList(i).Orig.y) And NpcList(i).Numero = Guardias Then
                MiNPC = NpcList(i)
                Call QuitarNPC(i, eResetPos)
                Call ReSpawnNpc(MiNPC)
            End If
        End If
    Next i
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Admin.ReSpawnOrigPosNpcs", Erl)
End Sub

Sub WorldSave()
    On Error GoTo Handler
    Dim LoopX As Integer
    Dim Porc  As Long
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg("1732", vbNullString, e_FontTypeNames.FONTTYPE_SERVER)) 'Msg1732=Servidor » Iniciando WorldSave
    Call ReSpawnOrigPosNpcs 'respawn de los guardias en las pos originales
    Dim j As Integer, K As Integer
    For j = 1 To NumMaps
        If MapInfo(j).backup_mode = 1 Then K = K + 1
    Next j
    FrmStat.ProgressBar1.Min = 0
    FrmStat.ProgressBar1.max = K
    FrmStat.ProgressBar1.value = 0
    For LoopX = 1 To NumMaps
        'DoEvents
        If MapInfo(LoopX).backup_mode = 1 Then
            FrmStat.ProgressBar1.value = FrmStat.ProgressBar1.value + 1
        End If
    Next LoopX
    FrmStat.Visible = False
    Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1733, vbNullString, e_FontTypeNames.FONTTYPE_SERVER))
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "Admin.WorldSave", Erl)
End Sub

Public Sub PurgarPenas()
    On Error GoTo PurgarPenas_Err
    Dim i As Long
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.Map, Libertad.x, Libertad.y, True)
                    'Msg1103= Has sido liberado.
                    Call WriteLocaleMsg(i, "1103", e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    Next i
    Exit Sub
PurgarPenas_Err:
    Call TraceError(Err.Number, Err.Description, "Admin.PurgarPenas", Erl)
End Sub

Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal minutos As Long, Optional ByVal GmName As String = vbNullString)
    On Error GoTo Encarcelar_Err
    If EsGM(UserIndex) Then Exit Sub
    UserList(UserIndex).Counters.Pena = minutos
    Call WarpUserChar(UserIndex, Prision.Map, Prision.x, Prision.y, True)
    If LenB(GmName) = 0 Then
        'Msg1107= Has sido encarcelado, deberas permanecer en la carcel  ¬1 minutos.
        Call WriteLocaleMsg(UserIndex, 1107, e_FontTypeNames.FONTTYPE_INFO, minutos)
    Else
        Call WriteLocaleMsg(UserIndex, 1617, e_FontTypeNames.FONTTYPE_INFO, GmName & "¬" & minutos) 'Msg1617=¬1 te ha encarcelado, deberás permanecer en la cárcel ¬2 minutos.
    End If
    Exit Sub
Encarcelar_Err:
    Call TraceError(Err.Number, Err.Description, "Admin.Encarcelar", Erl)
End Sub

Public Function BANCheck(ByVal name As String) As Boolean
    On Error GoTo BANCheck_Err
    BANCheck = BANCheckDatabase(name)
    Exit Function
BANCheck_Err:
    Call TraceError(Err.Number, Err.Description, "Admin.BANCheck", Erl)
End Function

Public Function PersonajeExiste(ByVal name As String) As Boolean
    On Error GoTo PersonajeExiste_Err
    PersonajeExiste = GetUserValue(LCase$(name), "COUNT(*)") > 0
    Exit Function
PersonajeExiste_Err:
    Call TraceError(Err.Number, Err.Description, "Admin.PersonajeExiste", Erl)
End Function

Public Function IsValidUserId(ByVal UserId As Long) As Boolean
    IsValidUserId = GetUserValueById(UserId, "COUNT(*)") > 0
End Function

Public Function UnBan(ByVal name As String) As Boolean
    On Error GoTo UnBan_Err
    Call UnBanDatabase(name)
    'Remove it from the banned people database
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", name, "BannedBy", "")
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", name, "Reason", "")
    Exit Function
UnBan_Err:
    Call TraceError(Err.Number, Err.Description, "Admin.UnBan", Erl)
End Function

Public Function UserDarPrivilegioLevel(ByVal name As String) As e_PlayerType
    On Error GoTo UserDarPrivilegioLevel_Err
    '***************************************************
    'Author: Unknown
    'Last Modification: 03/02/07
    'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
    '***************************************************
    If EsAdmin(name) Then
        UserDarPrivilegioLevel = e_PlayerType.Admin
    ElseIf EsDios(name) Then
        UserDarPrivilegioLevel = e_PlayerType.Dios
    ElseIf EsSemiDios(name) Then
        UserDarPrivilegioLevel = e_PlayerType.SemiDios
    ElseIf EsConsejero(name) Then
        UserDarPrivilegioLevel = e_PlayerType.Consejero
    Else
        UserDarPrivilegioLevel = e_PlayerType.User
    End If
    Exit Function
UserDarPrivilegioLevel_Err:
    Call TraceError(Err.Number, Err.Description, "Admin.UserDarPrivilegioLevel", Erl)
End Function

Public Sub BanTemporal(ByVal nombre As String, ByVal dias As Integer, Causa As String, Baneador As String)
    On Error GoTo BanTemporal_Err
    Dim tBan As tBaneo
    Set tBan = New tBaneo
    tBan.name = UCase$(nombre)
    tBan.FechaLiberacion = (Now + dias)
    tBan.Causa = Causa
    tBan.Baneador = Baneador
    Call Baneos.Add(tBan)
    Call SaveBan(Baneos.count)
    Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageLocaleMsg(1705, nombre & "¬" & Causa & "¬" & dias & "¬" & Baneador, e_FontTypeNames.FONTTYPE_SERVER)) 'Msg1705=¬1 fue baneado por ¬2 durante los próximos ¬3 días. La medida fue tomada por: ¬4.
    Exit Sub
BanTemporal_Err:
    Call TraceError(Err.Number, Err.Description, "Admin.BanTemporal", Erl)
End Sub

Sub SaveBans()
    On Error GoTo SaveBans_Err
    Dim num As Integer
    Call WriteVar(DatPath & "baneos.dat", "INIT", "NumeroBans", Baneos.count)
    For num = 1 To Baneos.count
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "USER", Baneos(num).name)
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "FECHA", Baneos(num).FechaLiberacion)
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "BANEADOR", Baneos(num).Baneador)
        Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "CAUSA", Baneos(num).Causa)
    Next
    Exit Sub
SaveBans_Err:
    Call TraceError(Err.Number, Err.Description, "Admin.SaveBans", Erl)
End Sub

Sub SaveBan(num As Integer)
    On Error GoTo SaveBan_Err
    Call WriteVar(DatPath & "baneos.dat", "INIT", "NumeroBans", Baneos.count)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "USER", Baneos(num).name)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "FECHA", Baneos(num).FechaLiberacion)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "BANEADOR", Baneos(num).Baneador)
    Call WriteVar(DatPath & "baneos.dat", "BANEO" & num, "CAUSA", Baneos(num).Causa)
    Call SaveBanDatabase(Baneos(num).name, Baneos(num).Causa, Baneos(num).Baneador)
    Exit Sub
SaveBan_Err:
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Admin.SaveBan", Erl)
End Sub

Sub LoadBans()
    On Error GoTo LoadBans_Err
    Dim BaneosTemporales As Integer
    Dim tBan             As tBaneo, i As Integer
    If Not FileExist(DatPath & "baneos.dat", vbNormal) Then Exit Sub
    BaneosTemporales = val(GetVar(DatPath & "baneos.dat", "INIT", "NumeroBans"))
    If BaneosTemporales > 0 Then
        For i = 1 To BaneosTemporales
            Set tBan = New tBaneo
            With tBan
                .name = GetVar(DatPath & "baneos.dat", "BANEO" & i, "USER")
                .FechaLiberacion = GetVar(DatPath & "baneos.dat", "BANEO" & i, "FECHA")
                .Causa = GetVar(DatPath & "baneos.dat", "BANEO" & i, "CAUSA")
                .Baneador = GetVar(DatPath & "baneos.dat", "BANEO" & i, "BANEADOR")
                Call Baneos.Add(tBan)
            End With
        Next
    End If
    Exit Sub
LoadBans_Err:
    Call TraceError(Err.Number, Err.Description, "Argentum20Server.Admin.LoadBans", Erl)
End Sub

Public Function CompararUserPrivilegios(ByVal Personaje_1 As Integer, ByVal Personaje_2 As Integer) As Integer
    CompararUserPrivilegios = CompararPrivilegios(UserList(Personaje_1).flags.Privilegios, UserList(Personaje_2).flags.Privilegios)
End Function

Public Function CompararPrivilegiosUser(ByVal Personaje_1 As Integer, ByVal Personaje_2 As Integer) As Integer
    On Error GoTo CompararPrivilegiosUser_Err
    CompararPrivilegiosUser = CompararPrivilegios(UserList(Personaje_1).flags.Privilegios, UserList(Personaje_2).flags.Privilegios)
    Exit Function
CompararPrivilegiosUser_Err:
    Call TraceError(Err.Number, Err.Description, "Admin.CompararPrivilegiosUser", Erl)
End Function

Public Function CompararPrivilegios(ByVal Izquierda As e_PlayerType, ByVal Derecha As e_PlayerType) As Integer
    '**************************************************************************************************************************
    'Author: Jopi
    'Last Modification: 05/07/2020
    '   Funcion encargada de comparar los privilegios entre 2 Game Masters.
    '   Funciona de la misma forma que el operador spaceship de PHP.
    '       - Si los privilegios de el de la izquierda son MAYORES que el de la derecha, devuelve 1
    '       - Si los privilegios de el de la izquierda son IGUALES que el de la derecha, devuelve 0
    '       - Si los privilegios de el de la izquierda son MENORES que el de la derecha, devuelve -1
    '**************************************************************************************************************************
    On Error GoTo CompararPrivilegios_Err
    Dim PrivilegiosGM As e_PlayerType
    PrivilegiosGM = e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster
    ' Obtenemos el rango de los 2 personajes.
    Izquierda = (Izquierda And PrivilegiosGM)
    Derecha = (Derecha And PrivilegiosGM)
    Select Case Izquierda
        Case Is > Derecha
            CompararPrivilegios = 1
        Case Is = Derecha
            CompararPrivilegios = 0
        Case Is < Derecha
            CompararPrivilegios = -1
    End Select
    Exit Function
CompararPrivilegios_Err:
    Call TraceError(Err.Number, Err.Description, "Admin.CompararPrivilegios", Erl)
End Function
