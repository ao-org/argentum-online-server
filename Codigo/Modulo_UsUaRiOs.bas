Attribute VB_Name = "UserMod"
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
Private UserNameCache     As New Dictionary
Private AvailableUserSlot As t_IndexHeap

Public Sub InitializeUserIndexHeap(Optional ByVal Size As Integer = NpcIndexHeapSize)
    On Error GoTo ErrHandler_InitializeUserIndexHeap
    ReDim AvailableUserSlot.IndexInfo(Size)
    Dim i As Integer
    For i = 1 To Size
        AvailableUserSlot.IndexInfo(i) = Size - (i - 1)
        UserList(AvailableUserSlot.IndexInfo(i)).flags.IsSlotFree = True
    Next i
    AvailableUserSlot.currentIndex = Size
    Exit Sub
ErrHandler_InitializeUserIndexHeap:
    Call TraceError(Err.Number, Err.Description, "UserMod.InitializeUserIndexHeap", Erl)
End Sub

Public Function ReleaseUser(ByVal UserIndex As Integer) As Boolean
    On Error GoTo ErrHandler
    If UserList(UserIndex).flags.IsSlotFree Then
        ReleaseUser = False
        Exit Function
    End If
    If IsFeatureEnabled("debug_id_assign") Then
        Call LogError("Releasing usedid: " & UserIndex)
    End If
    AvailableUserSlot.currentIndex = AvailableUserSlot.currentIndex + 1
    Debug.Assert AvailableUserSlot.currentIndex <= UBound(AvailableUserSlot.IndexInfo)
    AvailableUserSlot.IndexInfo(AvailableUserSlot.currentIndex) = UserIndex
    UserList(UserIndex).flags.IsSlotFree = True
    ReleaseUser = True
    Exit Function
ErrHandler:
    ReleaseUser = False
    Call TraceError(Err.Number, Err.Description, "UserMod.ReleaseUser", Erl)
End Function

Public Function GetAvailableUserSlot() As Integer
    GetAvailableUserSlot = AvailableUserSlot.currentIndex
End Function

Public Function IsPatreon(ByVal UserIndex As Integer) As Boolean
    
   On Error GoTo IsPatreon_Error

    With UserList(UserIndex).Stats
        IsPatreon = .tipoUsuario = e_TipoUsuario.tAventurero Or .tipoUsuario = e_TipoUsuario.tHeroe Or .tipoUsuario = e_TipoUsuario.tLeyenda
    End With

   On Error GoTo 0
   Exit Function

IsPatreon_Error:
    Call Logging.TraceError(Err.Number, Err.Description, "UserMod.IsPatreon nick: " & UserList(UserIndex).name, Erl())
    
End Function


Public Function GetNextAvailableUserSlot() As Integer
    On Error GoTo ErrHandler
    If (AvailableUserSlot.currentIndex = 0) Then
        GetNextAvailableUserSlot = -1
        Return
    End If
    GetNextAvailableUserSlot = AvailableUserSlot.IndexInfo(AvailableUserSlot.currentIndex)
    AvailableUserSlot.currentIndex = AvailableUserSlot.currentIndex - 1
    If Not UserList(GetNextAvailableUserSlot).flags.IsSlotFree Then
        Call TraceError(Err.Number, "Trying to active the same user slot twice", "UserMod.GetNextAvailableUserSlot", Erl)
        GetNextAvailableUserSlot = -1
    End If
    UserList(GetNextAvailableUserSlot).flags.IsSlotFree = False
    Exit Function
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "UserMod.GetNextAvailableUserSlot", Erl)
End Function

Public Function GetUserName(ByVal UserId As Long) As String
    On Error GoTo GetUserName_Err
    If UserId <= 0 Then
        GetUserName = ""
        Exit Function
    End If
    If UserNameCache.Exists(UserId) Then
        GetUserName = UserNameCache.Item(UserId)
        Exit Function
    End If
    Dim username As String
    username = GetCharacterName(UserId)
    Call RegisterUserName(UserId, username)
    GetUserName = username
    Exit Function
GetUserName_Err:
    Call TraceError(Err.Number, Err.Description, "UserMod.GetUserName", Erl)
End Function

Public Sub RegisterUserName(ByVal UserId As Long, ByVal username As String)
    If UserNameCache.Exists(UserId) Then
        UserNameCache.Item(UserId) = username
    Else
        UserNameCache.Add UserId, username
    End If
End Sub

Public Function IsValidUserRef(ByRef UserRef As t_UserReference) As Boolean
    IsValidUserRef = False
    If UserRef.ArrayIndex <= 0 Or UserRef.ArrayIndex > UBound(UserList) Then
        Exit Function
    End If
    If UserList(UserRef.ArrayIndex).VersionId <> UserRef.VersionId Then
        Exit Function
    End If
    IsValidUserRef = True
End Function

Public Function SetUserRef(ByRef UserRef As t_UserReference, ByVal Index As Integer) As Boolean
    SetUserRef = False
    UserRef.ArrayIndex = Index
    If Index <= 0 Or UserRef.ArrayIndex > UBound(UserList) Then
        Exit Function
    End If
    UserRef.VersionId = UserList(Index).VersionId
    SetUserRef = True
End Function

Public Sub ClearUserRef(ByRef UserRef As t_UserReference)
    UserRef.ArrayIndex = 0
    UserRef.VersionId = -1
End Sub

Public Sub IncreaseVersionId(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .VersionId > 32760 Then
            .VersionId = 0
        Else
            .VersionId = .VersionId + 1
        End If
    End With
End Sub

Public Sub LogUserRefError(ByRef UserRef As t_UserReference, ByRef Text As String)
    Call LogError("Failed to validate UserRef index(" & UserRef.ArrayIndex & ") version(" & UserRef.VersionId & ") got versionId: " & UserList(UserRef.ArrayIndex).VersionId & _
            " At: " & Text)
End Sub

Public Function ConnectUser_Check(ByVal UserIndex As Integer, ByVal name As String) As Boolean
    On Error GoTo Check_ConnectUser_Err
    ConnectUser_Check = False
    'Controlamos no pasar el maximo de usuarios
    If NumUsers >= MaxUsers Then
        Call WriteShowMessageBox(UserIndex, 1759, vbNullString) 'Msg1759=El servidor ha alcanzado el máximo de usuarios soportado, por favor vuelva a intentarlo más tarde.
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    If EnPausa Then
        Call WritePauseToggle(UserIndex)
        ' Msg520=Servidor » Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.
        Call WriteLocaleMsg(UserIndex, 520, e_FontTypeNames.FONTTYPE_SERVER)
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    If Not EsGM(UserIndex) And ServerSoloGMs > 0 Then
        Call WriteShowMessageBox(UserIndex, 1760, vbNullString) 'Msg1760=Servidor restringido a administradores. Por favor reintente en unos momentos.
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    With UserList(UserIndex)
        If .flags.UserLogged Then
            Call LogSecurity("User " & .name & " trying to log and already an already logged character from IP: " & .ConnectionDetails.IP)
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)
            Exit Function
        End If
        '¿Ya esta conectado el personaje?
        Dim tIndex As t_UserReference: tIndex = NameIndex(name)
        If tIndex.ArrayIndex > 0 Then
            If Not IsValidUserRef(tIndex) Then
                Call CloseSocket(tIndex.ArrayIndex)
            ElseIf IsFeatureEnabled("override_same_ip_connection") And .ConnectionDetails.IP = UserList(tIndex.ArrayIndex).ConnectionDetails.IP Then
                Call WriteShowMessageBox(tIndex.ArrayIndex, 1761, vbNullString) 'Msg1761=Alguien está ingresando con tu personaje. Si no has sido tú, por favor cambia la contraseña de tu cuenta.
                Call CloseSocket(tIndex.ArrayIndex)
            Else
                If UserList(tIndex.ArrayIndex).Counters.Saliendo Then
                    Call WriteShowMessageBox(UserIndex, 1762, vbNullString) 'Msg1762=El personaje está saliendo.
                Else
                    Call WriteShowMessageBox(UserIndex, 1763, vbNullString) 'Msg1763=El personaje ya está conectado. Espere mientras es desconectado.
                    ' Le avisamos al usuario que está jugando, en caso de que haya uno
                    Call WriteShowMessageBox(tIndex.ArrayIndex, 1761, vbNullString) 'Msg1761=Alguien está ingresando con tu personaje. Si no has sido tú, por favor cambia la contraseña de tu cuenta.
                End If
                Call CloseSocket(UserIndex)
                Exit Function
            End If
        End If
        
#If LOGIN_STRESS_TEST = 0 Then
        '¿Supera el máximo de usuarios por cuenta?
        If MaxUsersPorCuenta > 0 Then
            If ContarUsuariosMismaCuenta(.AccountID) >= MaxUsersPorCuenta Then
                If MaxUsersPorCuenta = 1 Then
                    Call WriteShowMessageBox(UserIndex, 1764, vbNullString) 'Msg1764=Ya hay un usuario conectado con esta cuenta.
                Else
                    Call WriteShowMessageBox(UserIndex, 1765, MaxUsersPorCuenta) 'Msg1765=La cuenta ya alcanzó el máximo de ¬1 usuarios conectados.
                End If
                Call CloseSocket(UserIndex)
                Exit Function
            End If
        End If
#End If
        .flags.Privilegios = UserDarPrivilegioLevel(name)
        If EsRolesMaster(name) Then
            .flags.Privilegios = .flags.Privilegios Or e_PlayerType.RoleMaster
        End If
        If EsGM(UserIndex) Then
            Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageLocaleMsg(1706, name, e_FontTypeNames.FONTTYPE_INFOBOLD)) 'Msg1706=Servidor » ¬1 se conecto al juego.
            Call LogGM(name, "Se conectó con IP: " & .ConnectionDetails.IP)
        End If
    End With
    ConnectUser_Check = True
    Exit Function
Check_ConnectUser_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ConnectUser_Check", Erl)
End Function

Public Sub ConnectUser_Prepare(ByVal UserIndex As Integer, ByVal name As String)
    On Error GoTo Prepare_ConnectUser_Err
    With UserList(UserIndex)
        .flags.Escondido = 0
        Call ClearNpcRef(.flags.TargetNPC)
        .flags.TargetNpcTipo = e_NPCType.Comun
        .flags.TargetObj = 0
        Call SetUserRef(.flags.TargetUser, 0)
        .Char.FX = 0
        .Counters.CuentaRegresiva = -1
        .name = name
        Dim UserRef As New clsUserRefWrapper
        UserRef.SetFromIndex (UserIndex)
        Set m_NameIndex(UCase$(name)) = UserRef
        .showName = True
        .NroMascotas = 0
    End With
    Exit Sub
Prepare_ConnectUser_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ConnectUser_Prepare", Erl)
End Sub

Public Function ConnectUser_Complete(ByVal UserIndex As Integer, ByRef name As String, Optional ByVal newUser As Boolean = False)

Dim n                           As Integer
Dim tStr                        As String

    On Error GoTo Complete_ConnectUser_Err
    
    ConnectUser_Complete = False
    Call SendData(SendTarget.ToIndex, UserIndex, PrepareActiveToggles)
    
    With UserList(UserIndex)
#If LOGIN_STRESS_TEST = 1 Then
        .pos.Map = 1 'Ullathorpe
#End If
        If .flags.Paralizado = 1 Then
            .Counters.Paralisis = IntervaloParalizado
        End If
        If .flags.Muerto = 0 Then
            .Char = .OrigChar
            If .Char.body = 0 Then
                Call SetNakedBody(UserList(UserIndex))
            End If
            If .Char.head = 0 Then
                .Char.head = 1
            End If
        Else
            .Char.body = iCuerpoMuerto
            .Char.head = iCabezaMuerto
            .Char.WeaponAnim = NingunArma
            .Char.ShieldAnim = NingunEscudo
            .Char.CascoAnim = NingunCasco
            .Char.CartAnim = NoCart
            .Char.Heading = e_Heading.SOUTH
        End If
        .Stats.UserAtributos(e_Atributos.Fuerza) = 18 + ModRaza(.raza).Fuerza
        .Stats.UserAtributos(e_Atributos.Agilidad) = 18 + ModRaza(.raza).Agilidad
        .Stats.UserAtributos(e_Atributos.Inteligencia) = 18 + ModRaza(.raza).Inteligencia
        .Stats.UserAtributos(e_Atributos.Constitucion) = 18 + ModRaza(.raza).Constitucion
        .Stats.UserAtributos(e_Atributos.Carisma) = 18 + ModRaza(.raza).Carisma
        .Stats.UserAtributosBackUP(e_Atributos.Fuerza) = .Stats.UserAtributos(e_Atributos.Fuerza)
        .Stats.UserAtributosBackUP(e_Atributos.Agilidad) = .Stats.UserAtributos(e_Atributos.Agilidad)
        .Stats.UserAtributosBackUP(e_Atributos.Inteligencia) = .Stats.UserAtributos(e_Atributos.Inteligencia)
        .Stats.UserAtributosBackUP(e_Atributos.Constitucion) = .Stats.UserAtributos(e_Atributos.Constitucion)
        .Stats.UserAtributosBackUP(e_Atributos.Carisma) = .Stats.UserAtributos(e_Atributos.Carisma)
        .Stats.MaxMAN = UserMod.GetMaxMana(UserIndex)
        .Stats.MaxSta = UserMod.GetMaxStamina(UserIndex)
        .Stats.MinHIT = UserMod.GetHitModifier(UserIndex) + 1
        .Stats.MaxHit = UserMod.GetHitModifier(UserIndex) + 2
        .Stats.MinHp = Min(.Stats.MinHp, .Stats.MaxHp)
        .Stats.MinMAN = Min(.Stats.MinMAN, UserMod.GetMaxMana(UserIndex))
        'Obtiene el indice-objeto del arma
        If .invent.EquippedWeaponSlot > 0 Then
            If .invent.Object(.invent.EquippedWeaponSlot).ObjIndex > 0 Then
                .invent.EquippedWeaponObjIndex = .invent.Object(.invent.EquippedWeaponSlot).ObjIndex
                If .flags.Muerto = 0 Then
                    .Char.Arma_Aura = ObjData(.invent.EquippedWeaponObjIndex).CreaGRH
                End If
            Else
                .invent.EquippedWeaponSlot = 0
            End If
        End If
        ' clear hotkey settings, the client should set this
        For n = 0 To HotKeyCount - 1
            .HotkeyList(n).Index = -1
            .HotkeyList(n).LastKnownSlot = -1
            .HotkeyList(n).Type = Unknown
        Next n
        'Obtiene el indice-objeto del armadura
        If .invent.EquippedArmorSlot > 0 Then
            If .invent.Object(.invent.EquippedArmorSlot).ObjIndex > 0 Then
                .invent.EquippedArmorObjIndex = .invent.Object(.invent.EquippedArmorSlot).ObjIndex
                If .flags.Muerto = 0 Then
                    .Char.Body_Aura = ObjData(.invent.EquippedArmorObjIndex).CreaGRH
                End If
            Else
                .invent.EquippedArmorSlot = 0
            End If
            .flags.Desnudo = 0
        Else
            .flags.Desnudo = 1
        End If
        
        If .Invent_Skins.SlotBackpackEquipped > 0 Then
            If .Invent_Skins.Object(.Invent_Skins.SlotBackpackEquipped).ObjIndex = .Invent_Skins.ObjIndexBackpackEquipped And .Invent_Skins.ObjIndexBackpackEquipped > 0 Then
                If CanEquipSkin(UserIndex, .Invent_Skins.SlotBackpackEquipped, False) Then
                    Call SkinEquip(UserIndex, .Invent_Skins.SlotBackpackEquipped, .Invent_Skins.Object(.Invent_Skins.SlotBackpackEquipped).ObjIndex)
                End If
            End If
        End If
        
        'Obtiene el indice-objeto del escudo
        If .invent.EquippedShieldSlot > 0 Then
            If .invent.Object(.invent.EquippedShieldSlot).ObjIndex > 0 Then
                .invent.EquippedShieldObjIndex = .invent.Object(.invent.EquippedShieldSlot).ObjIndex
                If .flags.Muerto = 0 Then
                    .Char.Escudo_Aura = ObjData(.invent.EquippedShieldObjIndex).CreaGRH
                End If
            Else
                .invent.EquippedShieldSlot = 0
            End If
        End If
        'Obtiene el indice-objeto del casco
        If .invent.EquippedHelmetSlot > 0 Then
            If .invent.Object(.invent.EquippedHelmetSlot).ObjIndex > 0 Then
                .invent.EquippedHelmetObjIndex = .invent.Object(.invent.EquippedHelmetSlot).ObjIndex
                If .flags.Muerto = 0 Then
                    .Char.Head_Aura = ObjData(.invent.EquippedHelmetObjIndex).CreaGRH
                End If
            Else
                .invent.EquippedHelmetSlot = 0
            End If
        End If
        'Obtiene el indice-objeto barco
        If .invent.EquippedShipSlot > 0 Then
            If .invent.Object(.invent.EquippedShipSlot).ObjIndex > 0 Then
                .invent.EquippedShipObjIndex = .invent.Object(.invent.EquippedShipSlot).ObjIndex
            Else
                .invent.EquippedShipSlot = 0
            End If
        End If
        'Obtiene el indice-objeto municion
        If .invent.EquippedMunitionSlot > 0 Then
            If .invent.Object(.invent.EquippedMunitionSlot).ObjIndex > 0 Then
                .invent.EquippedMunitionObjIndex = .invent.Object(.invent.EquippedMunitionSlot).ObjIndex
            Else
                .invent.EquippedMunitionSlot = 0
            End If
        End If
        ' DM
        If .invent.EquippedRingAccesorySlot > 0 Then
            If .invent.Object(.invent.EquippedRingAccesorySlot).ObjIndex > 0 Then
                .invent.EquippedRingAccesoryObjIndex = .invent.Object(.invent.EquippedRingAccesorySlot).ObjIndex
                If .flags.Muerto = 0 Then
                    .Char.DM_Aura = ObjData(.invent.EquippedRingAccesoryObjIndex).CreaGRH
                End If
            Else
                .invent.EquippedRingAccesorySlot = 0
            End If
        End If
        If .invent.EquippedAmuletAccesorySlot > 0 Then
            .invent.EquippedAmuletAccesoryObjIndex = .invent.Object(.invent.EquippedAmuletAccesorySlot).ObjIndex
            If ObjData(.invent.EquippedAmuletAccesoryObjIndex).CreaGRH <> "" Then
                .Char.Otra_Aura = ObjData(.invent.EquippedAmuletAccesoryObjIndex).CreaGRH
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Otra_Aura, False, 5))
            End If
            If ObjData(.invent.EquippedAmuletAccesoryObjIndex).Ropaje > 0 Then
                .Char.CartAnim = ObjData(.invent.EquippedAmuletAccesoryObjIndex).Ropaje
            End If
        End If
        ' RM
        If .invent.EquippedRingAccesorySlot > 0 Then
            If .invent.Object(.invent.EquippedRingAccesorySlot).ObjIndex > 0 Then
                .invent.EquippedRingAccesoryObjIndex = .invent.Object(.invent.EquippedRingAccesorySlot).ObjIndex
                If .flags.Muerto = 0 Then
                    .Char.RM_Aura = ObjData(.invent.EquippedRingAccesoryObjIndex).CreaGRH
                End If
            Else
                .invent.EquippedRingAccesorySlot = 0
            End If
        End If
        If .invent.EquippedSaddleSlot > 0 Then
            If .invent.Object(.invent.EquippedSaddleSlot).ObjIndex > 0 Then
                .invent.EquippedSaddleObjIndex = .invent.Object(.invent.EquippedSaddleSlot).ObjIndex
            Else
                .invent.EquippedSaddleSlot = 0
            End If
        End If
        If .invent.EquippedWorkingToolSlot > 0 Then
            If .invent.Object(.invent.EquippedWorkingToolSlot).ObjIndex Then
                .invent.EquippedWorkingToolObjIndex = .invent.Object(.invent.EquippedWorkingToolSlot).ObjIndex
            Else
                .invent.EquippedWorkingToolSlot = 0
            End If
        End If
        If .invent.EquippedAmuletAccesorySlot > 0 Then
            If .invent.Object(.invent.EquippedAmuletAccesorySlot).ObjIndex Then
                .invent.EquippedAmuletAccesoryObjIndex = .invent.Object(.invent.EquippedAmuletAccesorySlot).ObjIndex
                If .flags.Muerto = 0 Then
                    .Char.Otra_Aura = ObjData(.invent.EquippedAmuletAccesoryObjIndex).CreaGRH
                End If
            Else
                .invent.EquippedAmuletAccesorySlot = 0
            End If
        End If
        If .invent.EquippedShieldSlot = 0 Then .Char.ShieldAnim = NingunEscudo
        If .invent.EquippedHelmetSlot = 0 Then .Char.CascoAnim = NingunCasco
        If .invent.EquippedWeaponSlot = 0 And .invent.EquippedWorkingToolSlot = 0 Then .Char.WeaponAnim = NingunArma
        If .invent.EquippedAmuletAccesorySlot = 0 Then .Char.CartAnim = NoCart
        ' -----------------------------------------------------------------------
        '   FIN - INFORMACION INICIAL DEL PERSONAJE
        ' -----------------------------------------------------------------------
        If Not ValidateChr(UserIndex) Then
            Call WriteShowMessageBox(UserIndex, 1766, vbNullString) 'Msg1766=Error en el personaje. Comuniquese con el staff.
            Call CloseSocket(UserIndex)
            Exit Function
        End If
        .flags.SeguroParty = True
        .flags.SeguroClan = True
        .flags.SeguroResu = True
        .flags.LegionarySecure = True
        .CurrentInventorySlots = getMaxInventorySlots(UserIndex)
        Call WriteInventoryUnlockSlots(UserIndex)
        Call LoadUserIntervals(UserIndex)
        Call WriteIntervals(UserIndex)
        Call UpdateUserInv(True, UserIndex, 0)
        Call UpdateUserHechizos(True, UserIndex, 0)
        Call EnviarLlaves(UserIndex)
        If .flags.Paralizado Then Call WriteParalizeOK(UserIndex)
        If .flags.Inmovilizado Then Call WriteInmovilizaOK(UserIndex)
        .flags.Inmunidad = 1
        .Counters.TiempoDeInmunidad = IntervaloPuedeSerAtacado
        .Counters.TiempoDeInmunidadParalisisNoMagicas = 0
  
        If MapInfo(.pos.Map).MapResource = 0 Then
            .pos.Map = Ciudades(.Hogar).Map
            .pos.x = Ciudades(.Hogar).x
            .pos.y = Ciudades(.Hogar).y
        End If
        'Mapa válido
        If Not MapaValido(.pos.Map) Then
            Call WriteErrorMsg(UserIndex, "The character is in an invalid postion/map. Please ask for support on Discord.")
            Call CloseSocket(UserIndex)
            Exit Function
        End If
        
        
        If MapData(.pos.Map, .pos.x, .pos.y).UserIndex <> 0 Or MapData(.pos.Map, .pos.x, .pos.y).NpcIndex <> 0 Then
            Dim FoundPlace As Boolean
            Dim esAgua     As Boolean
            Dim nX         As Long
            Dim nY         As Long
        
            esAgua = (MapData(.pos.Map, .pos.x, .pos.y).Blocked And FLAG_AGUA) <> 0
        
            ' Busca el tile libre más cercano (espiral/radial) respetando agua/tierra
            FoundPlace = FindNearestFreeTile(.pos.Map, .pos.x, .pos.y, esAgua, SPAWN_SEARCH_MAX_RADIUS, nX, nY)
        
            If FoundPlace Then
                .pos.x = nX
                .pos.y = nY
            Else
                ' Sin lugar libre: si hay un usuario debajo, avisamos/cerramos comercio y lo desconectamos.
                Dim uidBelow As Integer
                uidBelow = MapData(.pos.Map, .pos.x, .pos.y).UserIndex
        
                If uidBelow <> 0 Then
                    ' Notificar al compañero de comercio (si corresponde)
                    If IsValidUserRef(UserList(uidBelow).ComUsu.DestUsu) Then
                        Dim destIdx As Integer
                        destIdx = UserList(uidBelow).ComUsu.DestUsu.ArrayIndex
                        If destIdx > 0 And UserList(destIdx).flags.UserLogged Then
                            Call FinComerciarUsu(destIdx)
                            Call WriteConsoleMsg(destIdx, _
                                PrepareMessageLocaleMsg(1925, vbNullString, e_FontTypeNames.FONTTYPE_WARNING)) ' "Comercio cancelado..."
                        End If
                    End If
        
                    ' Cerrar comercio del usuario pisado y avisarle
                    If UserList(uidBelow).flags.UserLogged Then
                        Call FinComerciarUsu(uidBelow)
                        Call WriteErrorMsg(uidBelow, "Somebody has connected to the game in the same position you were, please reconnect...")
                    End If
        
                    ' Desconectar al usuario debajo
                    Call CloseSocket(uidBelow)
                End If
                ' Si hay un NPC debajo, se pisa (comportamiento original).
            End If
        End If



        'If in the water, and has a boat, equip it!
        Dim trigger     As Integer
        Dim slotBarco   As Integer
        Dim itemBuscado As Integer
        trigger = MapData(.pos.Map, .pos.x, .pos.y).trigger
        If trigger = e_Trigger.DETALLEAGUA Then 'Esta en zona de caucho obj 199, 200
            If .raza = e_Raza.Enano Or .raza = e_Raza.Gnomo Then
                itemBuscado = iObjTrajeBajoNw
            Else
                itemBuscado = iObjTrajeAltoNw
            End If
            slotBarco = GetSlotInInventory(UserIndex, itemBuscado)
            If slotBarco > -1 Then
                .invent.EquippedShipObjIndex = itemBuscado
                .invent.EquippedShipSlot = slotBarco
            End If
        ElseIf trigger = e_Trigger.VALIDONADO Or trigger = e_Trigger.NADOCOMBINADO Then  'Esta en zona de nado comun obj 197
            itemBuscado = iObjTraje
            slotBarco = GetSlotInInventory(UserIndex, itemBuscado)
            If slotBarco > -1 Then
                .invent.EquippedShipObjIndex = itemBuscado
                .invent.EquippedShipSlot = slotBarco
            End If
        End If
        If .invent.EquippedShipObjIndex > 0 And (MapData(.pos.Map, .pos.x, .pos.y).Blocked And FLAG_AGUA) <> 0 Then
            .flags.Navegando = 1
            Call EquiparBarco(UserIndex)
        ElseIf .flags.Navegando = 1 And (MapData(.pos.Map, .pos.x, .pos.y).Blocked And FLAG_AGUA) <> 0 Then
            Dim iSlot As Integer
            For iSlot = 1 To UBound(.invent.Object)
                If .invent.Object(iSlot).ObjIndex > 0 Then
                    If ObjData(.invent.Object(iSlot).ObjIndex).OBJType = otShips And ObjData(.invent.Object(iSlot).ObjIndex).Subtipo > 0 Then
                        .invent.EquippedShipObjIndex = .invent.Object(iSlot).ObjIndex
                        .invent.EquippedShipSlot = iSlot
                        Exit For
                    End If
                End If
            Next
        End If
        If .invent.EquippedAmuletAccesoryObjIndex <> 0 Then
            If ObjData(.invent.EquippedAmuletAccesoryObjIndex).EfectoMagico = 11 Then .flags.Paraliza = 1
        End If
        Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
        Call WriteHora(UserIndex)
        Call WriteChangeMap(UserIndex, .pos.Map) 'Carga el mapa
        Call UpdateCharWithEquipedItems(UserIndex)
        Select Case .flags.Privilegios
            Case e_PlayerType.Admin
                .flags.ChatColor = RGB(252, 195, 0)
            Case e_PlayerType.Dios
                .flags.ChatColor = RGB(26, 209, 107)
            Case e_PlayerType.SemiDios
                .flags.ChatColor = RGB(60, 150, 60)
            Case e_PlayerType.Consejero
                .flags.ChatColor = RGB(170, 170, 170)
            Case Else
                .flags.ChatColor = vbWhite
        End Select
        Select Case .Faccion.Status
            Case e_Facciones.Ciudadano
                .flags.ChatColor = vbWhite
            Case e_Facciones.Armada
                .flags.ChatColor = vbWhite
            Case e_Facciones.consejo
                .flags.ChatColor = RGB(66, 201, 255)
            Case e_Facciones.Criminal
                .flags.ChatColor = vbWhite
            Case e_Facciones.Caos
                .flags.ChatColor = vbWhite
            Case e_Facciones.concilio
                .flags.ChatColor = RGB(255, 102, 102)
        End Select
        ' Jopi: Te saco de los mapas de retos (si logueas ahi) 324 372 389 390
        If Not EsGM(UserIndex) And (.pos.Map = 324 Or .pos.Map = 372 Or .pos.Map = 389 Or .pos.Map = 390) Then
            ' Si tiene una posicion a la que volver, lo mando ahi
            If MapaValido(.flags.ReturnPos.Map) And .flags.ReturnPos.x > 0 And .flags.ReturnPos.x <= XMaxMapSize And .flags.ReturnPos.y > 0 And .flags.ReturnPos.y <= YMaxMapSize _
                    Then
                Call WarpToLegalPos(UserIndex, .flags.ReturnPos.Map, .flags.ReturnPos.x, .flags.ReturnPos.y, True)
            Else ' Lo mando a su hogar
                Call WarpToLegalPos(UserIndex, Ciudades(.Hogar).Map, Ciudades(.Hogar).x, Ciudades(.Hogar).y, True)
            End If
        End If
        ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
        #If ConUpTime Then
            .LogOnTime = Now
        #End If
        'Crea  el personaje del usuario
        Call MakeUserChar(True, .pos.Map, UserIndex, .pos.Map, .pos.x, .pos.y, 1)
        Call WriteUserCharIndexInServer(UserIndex)
        Call ActualizarVelocidadDeUsuario(UserIndex)
        If .flags.Privilegios And (e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.Admin) Then
            Call DoAdminInvisible(UserIndex)
        End If
        Call WriteUpdateUserStats(UserIndex)
        Call WriteUpdateHungerAndThirst(UserIndex)
        Call WriteUpdateDM(UserIndex)
        Call WriteUpdateRM(UserIndex)
        Call SendMOTD(UserIndex)
        'Actualiza el Num de usuarios
        NumUsers = NumUsers + 1
        .flags.UserLogged = True
        Call Execute("Update user set is_logged = true where id = ?", UserList(UserIndex).Id)
        .Counters.LastSave = GetTickCountRaw()
        MapInfo(.pos.Map).NumUsers = MapInfo(.pos.Map).NumUsers + 1
        If .Stats.SkillPts > 0 Then
            Call WriteSendSkills(UserIndex)
            Call WriteLevelUp(UserIndex, .Stats.SkillPts)
        End If
        If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers
        If NumUsers > RecordUsuarios Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg("1550", NumUsers, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1550=Record de usuarios conectados simultáneamente: ¬1 usuarios.
            RecordUsuarios = NumUsers
        End If
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageOnlineUser(NumUsers))
        Call WriteFYA(UserIndex)
        Call WriteBindKeys(UserIndex)
        If .NroMascotas > 0 And MapInfo(.pos.Map).NoMascotas = 0 And .flags.MascotasGuardadas = 0 Then
            Dim i As Integer
            For i = 1 To MAXMASCOTAS
                If .MascotasType(i) > 0 Then
                    Call SetNpcRef(.MascotasIndex(i), SpawnNpc(.MascotasType(i), .pos, False, False, False, UserIndex))
                    If .MascotasIndex(i).ArrayIndex > 0 Then
                        Call SetUserRef(NpcList(.MascotasIndex(i).ArrayIndex).MaestroUser, UserIndex)
                        Call FollowAmo(.MascotasIndex(i).ArrayIndex)
                    End If
                End If
            Next i
        End If
        If .flags.Montado = 1 Then
            Call WriteEquiteToggle(UserIndex)
        End If
        Call ActualizarVelocidadDeUsuario(UserIndex)
        If .GuildIndex > 0 Then
            'welcome to the show baby...
            If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
                ' Msg521=Tu estado no te permite entrar al clan.
                Call WriteLocaleMsg(UserIndex, 521, e_FontTypeNames.FONTTYPE_GUILD)
            End If
        End If
        If LenB(.LastGuildRejection) <> 0 Then
            Call WriteShowMessageBox(UserIndex, 1767, .LastGuildRejection) 'Msg1767=Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: ¬1
            .LastGuildRejection = vbNullString
            Call SaveUserGuildRejectionReason(.name, vbNullString)
        End If
        If Lloviendo Then Call WriteRainToggle(UserIndex)
        If ServidorNublado Then Call WriteNubesToggle(UserIndex)
        Call WriteLoggedMessage(UserIndex, newUser)
        If .Stats.ELV = 1 Then
            Call WriteLocaleMsg(UserIndex, 522, e_FontTypeNames.FONTTYPE_GUILD, .name) ' Msg522=¡Bienvenido a las tierras de Argentum Online! ¡<nombre> que tengas buen viaje y mucha suerte!
        Else
            Call WriteLocaleMsg(UserIndex, 1439, e_FontTypeNames.FONTTYPE_GUILD, .name & "¬" & .Stats.ELV & "¬" & get_map_name(.pos.Map)) ' Msg1439=¡Bienvenido de nuevo ¬1! Actualmente estas en el nivel ¬2 en ¬3, ¡buen viaje y mucha suerte!
        End If
        If Status(UserIndex) = e_Facciones.Criminal Or Status(UserIndex) = e_Facciones.Caos Or Status(UserIndex) = e_Facciones.concilio Then
            Call WriteSafeModeOff(UserIndex)
            .flags.Seguro = False
            Call WriteLegionarySecure(UserIndex, False)
            .flags.LegionarySecure = False
        Else
            .flags.Seguro = True
            Call WriteSafeModeOn(UserIndex)
        End If
        If LenB(.MENSAJEINFORMACION) > 0 Then
            Dim Lines() As String
            Lines = Split(.MENSAJEINFORMACION, vbNewLine)
            For i = 0 To UBound(Lines)
                If LenB(Lines(i)) > 0 Then
                    Call WriteConsoleMsg(UserIndex, Lines(i), e_FontTypeNames.FONTTYPE_New_DONADOR)
                End If
            Next
            .MENSAJEINFORMACION = vbNullString
        End If
        If EventoActivo Then
            Call WriteLocaleMsg(UserIndex, 1625, e_FontTypeNames.FONTTYPE_New_Eventos, PublicidadEvento & "¬" & TiempoRestanteEvento) 'Msg1625=¬1. Tiempo restante: ¬2 minuto(s).
        End If
        Call WriteContadores(UserIndex)
        Call WritePrivilegios(UserIndex)
        Call RestoreDCUserCache(UserIndex)
        Call CustomScenarios.UserConnected(UserIndex)
        Call AntiCheat.OnNewPlayerConnect(UserIndex)
    End With
    ConnectUser_Complete = True
    Exit Function
Complete_ConnectUser_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ConnectUser_Complete", Erl)
End Function

Sub ActStats(ByVal VictimIndex As Integer, ByVal attackerIndex As Integer)
    On Error GoTo ActStats_Err
    Dim DaExp       As Integer
    Dim EraCriminal As Byte
    DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)
    If UserList(attackerIndex).Stats.ELV < STAT_MAXELV Then
        UserList(attackerIndex).Stats.Exp = UserList(attackerIndex).Stats.Exp + DaExp
        If UserList(attackerIndex).Stats.Exp > MAXEXP Then UserList(attackerIndex).Stats.Exp = MAXEXP
        Call WriteUpdateExp(attackerIndex)
        Call CheckUserLevel(attackerIndex)
    End If
    Call WriteLocaleMsg(attackerIndex, "76", e_FontTypeNames.FONTTYPE_FIGHT, UserList(VictimIndex).name)
    Call WriteLocaleMsg(attackerIndex, "140", e_FontTypeNames.FONTTYPE_EXP, DaExp)
    Call WriteLocaleMsg(VictimIndex, "185", e_FontTypeNames.FONTTYPE_FIGHT, UserList(attackerIndex).name)
    If Not PeleaSegura(VictimIndex, attackerIndex) Then
        EraCriminal = Status(attackerIndex)
        If EraCriminal = 2 And Status(attackerIndex) < 2 Then
            Call RefreshCharStatus(attackerIndex)
        ElseIf EraCriminal < 2 And Status(attackerIndex) = 2 Then
            Call RefreshCharStatus(attackerIndex)
        End If
    End If
    Call UserMod.UserDie(VictimIndex)
    If TriggerZonaPelea(attackerIndex, attackerIndex) <> TRIGGER6_PERMITE Then
        If UserList(attackerIndex).Stats.UsuariosMatados < MAXUSERMATADOS Then
            UserList(attackerIndex).Stats.UsuariosMatados = UserList(attackerIndex).Stats.UsuariosMatados + 1
        End If
    End If
    Exit Sub
ActStats_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ActStats", Erl)
End Sub

Sub RevivirUsuario(ByVal UserIndex As Integer, Optional ByVal MedianteHechizo As Boolean, Optional ByVal CasterUserIndex As Integer = 0)
    On Error GoTo RevivirUsuario_Err
    With UserList(UserIndex)
        .flags.Muerto = 0
        .Stats.MinHp = .Stats.MaxHp
        ' El comportamiento cambia si usamos el hechizo Resucitar
        If MedianteHechizo And CasterUserIndex > 0 Then
            If IsFeatureEnabled("healers_and_tanks") And UserList(CasterUserIndex).flags.DivineBlood > 0 Then
                .Stats.MinHp = .Stats.MaxHp
            Else
                .Stats.MinHp = 1
                .Stats.MinHam = 0
                .Stats.MinAGU = 0
                .Stats.MinMAN = 0
                Call WriteUpdateHungerAndThirst(UserIndex)
            End If
        End If
        Call WriteUpdateHP(UserIndex)
        Call WriteUpdateMana(UserIndex)
        If .flags.Navegando = 1 Then
            Call EquiparBarco(UserIndex)
        Else
            .Char.head = .OrigChar.head
            If .invent.EquippedHelmetObjIndex > 0 Then
                .Char.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
            End If
            If .invent.EquippedShieldObjIndex > 0 Then
                .Char.ShieldAnim = ObjData(.invent.EquippedShieldObjIndex).ShieldAnim
            End If
            If .invent.EquippedWeaponObjIndex > 0 Then
                .Char.WeaponAnim = ObjData(.invent.EquippedWeaponObjIndex).WeaponAnim
                If ObjData(.invent.EquippedWeaponObjIndex).CreaGRH <> "" Then
                    .Char.Arma_Aura = ObjData(.invent.EquippedWeaponObjIndex).CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Arma_Aura, False, 1))
                End If
            End If
            If .invent.EquippedArmorObjIndex > 0 Then
                .Char.body = ObtenerRopaje(UserIndex, ObjData(.invent.EquippedArmorObjIndex))
                If ObjData(.invent.EquippedArmorObjIndex).CreaGRH <> "" Then
                    .Char.Body_Aura = ObjData(.invent.EquippedArmorObjIndex).CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Body_Aura, False, 2))
                End If
            Else
                Call SetNakedBody(UserList(UserIndex))
            End If
            If .invent.EquippedShieldObjIndex > 0 Then
                .Char.ShieldAnim = ObjData(.invent.EquippedShieldObjIndex).ShieldAnim
                If ObjData(.invent.EquippedShieldObjIndex).CreaGRH <> "" Then
                    .Char.Escudo_Aura = ObjData(.invent.EquippedShieldObjIndex).CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Escudo_Aura, False, 3))
                End If
            End If
            If .invent.EquippedHelmetObjIndex > 0 Then
                .Char.CascoAnim = ObjData(.invent.EquippedHelmetObjIndex).CascoAnim
                If ObjData(.invent.EquippedHelmetObjIndex).CreaGRH <> "" Then
                    .Char.Head_Aura = ObjData(.invent.EquippedHelmetObjIndex).CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Head_Aura, False, 4))
                End If
            End If
            If .invent.EquippedAmuletAccesoryObjIndex > 0 Then
                If ObjData(.invent.EquippedAmuletAccesoryObjIndex).CreaGRH <> "" Then
                    .Char.Otra_Aura = ObjData(.invent.EquippedAmuletAccesoryObjIndex).CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Otra_Aura, False, 5))
                End If
                If ObjData(.invent.EquippedAmuletAccesoryObjIndex).Ropaje > 0 Then
                    .Char.CartAnim = ObjData(.invent.EquippedAmuletAccesoryObjIndex).Ropaje
                End If
            End If
            If .invent.EquippedRingAccesoryObjIndex > 0 Then
                If ObjData(.invent.EquippedRingAccesoryObjIndex).CreaGRH <> "" Then
                    .Char.DM_Aura = ObjData(.invent.EquippedRingAccesoryObjIndex).CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.DM_Aura, False, 6))
                End If
            End If
            If .invent.EquippedRingAccesoryObjIndex > 0 Then
                If ObjData(.invent.EquippedRingAccesoryObjIndex).CreaGRH <> "" Then
                    .Char.RM_Aura = ObjData(.invent.EquippedRingAccesoryObjIndex).CreaGRH
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.RM_Aura, False, 7))
                End If
            End If
        End If
        Call ActualizarVelocidadDeUsuario(UserIndex)
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim, .Char.BackpackAnim)
        Call MakeUserChar(True, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, 0)
    End With
    Exit Sub
RevivirUsuario_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.RevivirUsuario", Erl)
End Sub

Sub ChangeUserChar(ByVal UserIndex As Integer, _
                   ByVal body As Integer, _
                   ByVal head As Integer, _
                   ByVal Heading As Byte, _
                   ByVal Arma As Integer, _
                   ByVal Escudo As Integer, _
                   ByVal Casco As Integer, _
                   ByVal Cart As Integer, _
                   ByVal BackPack As Integer)
    On Error GoTo ChangeUserChar_Err
    
    If IsSet(UserList(UserIndex).flags.StatusMask, e_StatusMask.eTransformed) Then Exit Sub
    With UserList(UserIndex).Char
        .body = body
        .head = head
        .Heading = Heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = Casco
        .CartAnim = Cart
        .BackpackAnim = BackPack
    
        If .charindex > 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(body, head, Heading, .charindex, Arma, Escudo, Cart, BackPack, .FX, .loops, Casco, False, UserList(UserIndex).flags.Navegando))
        End If
    End With
    

    Exit Sub
ChangeUserChar_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ChangeUserChar", Erl)
End Sub

Sub EraseUserChar(ByVal UserIndex As Integer, ByVal Desvanecer As Boolean, Optional ByVal FueWarp As Boolean = False)
    On Error GoTo ErrorHandler
    Dim Error As String
    Error = "1"
    If UserList(UserIndex).Char.charindex = 0 Then Exit Sub
    CharList(UserList(UserIndex).Char.charindex) = 0
    If UserList(UserIndex).Char.charindex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If
    Error = "2"
    #If UNIT_TEST = 0 Then
        'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(4, UserList(UserIndex).Char.charindex, Desvanecer, FueWarp))
        Error = "3"
        Call QuitarUser(UserIndex, UserList(UserIndex).pos.Map)
        Error = "4"
    #End If
    MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).UserIndex = 0
    Error = "5"
    UserList(UserIndex).Char.charindex = 0
    NumChars = NumChars - 1
    Error = "6"
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.EraseUserChar", Erl)
End Sub

Sub RefreshCharStatus(ByVal UserIndex As Integer)
    On Error GoTo RefreshCharStatus_Err
    '*************************************************
    'Author: Tararira
    'Last modified: 6/04/2007
    'Refreshes the status and tag of UserIndex.
    '*************************************************
    Dim klan As String, name As String
    If UserList(UserIndex).showName Then
        If UserList(UserIndex).flags.Mimetizado = e_EstadoMimetismo.Desactivado Then
            If UserList(UserIndex).GuildIndex > 0 Then
                klan = modGuilds.GuildName(UserList(UserIndex).GuildIndex)
                klan = " <" & klan & ">"
            End If
            name = UserList(UserIndex).name & klan
        Else
            name = UserList(UserIndex).NameMimetizado
        End If
        If UserList(UserIndex).clase = e_Class.Pirat Then
            If UserList(UserIndex).flags.Oculto = 1 Then
                name = vbNullString
            End If
        End If
    End If
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, UserList(UserIndex).Faccion.Status, name))
    Exit Sub
RefreshCharStatus_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.RefreshCharStatus", Erl)
End Sub

Sub MakeUserChar(ByVal toMap As Boolean, _
                 ByVal sndIndex As Integer, _
                 ByVal UserIndex As Integer, _
                 ByVal Map As Integer, _
                 ByVal x As Integer, _
                 ByVal y As Integer, _
                 Optional ByVal appear As Byte = 0)
    On Error GoTo HayError
    Dim charindex As Integer
    Dim TempName  As String
    If InMapBounds(Map, x, y) Then
        With UserList(UserIndex)
            'If needed make a new character in list
            If .Char.charindex = 0 Then
                charindex = NextOpenCharIndex
                .Char.charindex = charindex
                CharList(charindex) = UserIndex
                If .Grupo.EnGrupo Then
                    Call modSendData.SendData(ToGroup, UserIndex, PrepareUpdateGroupInfo(UserIndex))
                End If
            End If
            'Place character on map if needed
            If toMap Then MapData(Map, x, y).UserIndex = UserIndex
            'Send make character command to clients
            Dim klan       As String
            Dim clan_nivel As Byte
            If Not toMap Then
                If .showName Then
                    If .flags.Mimetizado = e_EstadoMimetismo.Desactivado Then
                        If .GuildIndex > 0 Then
                            klan = modGuilds.GuildName(.GuildIndex)
                            clan_nivel = modGuilds.NivelDeClan(.GuildIndex)
                            TempName = .name & " <" & klan & ">"
                        Else
                            klan = vbNullString
                            clan_nivel = 0
                            If .flags.EnConsulta Then
                                TempName = .name & " [CONSULTA]"
                            Else
                                TempName = .name
                            End If
                        End If
                    Else
                        TempName = .NameMimetizado
                    End If
                End If
                Call WriteCharacterCreate(sndIndex, .Char.body, .Char.head, .Char.Heading, .Char.charindex, x, y, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CartAnim, _
                        .Char.BackpackAnim, .Char.FX, 999, .Char.CascoAnim, TempName, .Faccion.Status, .flags.Privilegios, .Char.ParticulaFx, .Char.Head_Aura, .Char.Arma_Aura, _
                        .Char.Body_Aura, .Char.DM_Aura, .Char.RM_Aura, .Char.Otra_Aura, .Char.Escudo_Aura, .Char.speeding, 0, appear, .Grupo.Lider.ArrayIndex, .GuildIndex, _
                        clan_nivel, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMAN, .Stats.MaxMAN, 0, False, .flags.Navegando, .Stats.tipoUsuario, .flags.CurrentTeam, _
                        .flags.tiene_bandera)
            Else
                'Hide the name and clan - set privs as normal user
                Call AgregarUser(UserIndex, .pos.Map, appear)
            End If
        End With
    End If
    Exit Sub
HayError:
    Dim Desc As String
    Desc = Err.Description & vbNewLine & " Usuario: " & UserList(UserIndex).name & vbNewLine & "Pos: " & Map & "-" & x & "-" & y
    Call TraceError(Err.Number, Err.Description, "Usuarios.MakeUserChar", Erl())
    Call CloseSocket(UserIndex)
End Sub

Sub CheckUserLevel(ByVal UserIndex As Integer)
    '*************************************************
    'Author: Unknown
    'Last modified: 01/10/2007
    'Chequea que el usuario no halla alcanzado el siguiente nivel,
    'de lo contrario le da la vida, mana, etc, correspodiente.
    '07/08/2006 Integer - Modificacion de los valores
    '01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV
    '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
    '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
    '13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitución.
    '09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consitución se controla desde Balance.dat
    '17/12/2020 WyroX - Distribución normal de las vidas
    '15/07/2024 Shugar - Vuelvo a implementar vidas variables y les agrego un capeo min/max
    '*************************************************
    On Error GoTo ErrHandler
    Dim Pts                 As Integer
    Dim AumentoHIT          As Integer
    Dim AumentoMANA         As Integer
    Dim AumentoSta          As Integer
    Dim AumentoHP           As Integer
    Dim WasNewbie           As Boolean
    Dim PromBias            As Double
    Dim PromClaseRaza       As Double
    Dim PromPersonaje       As Double
    Dim aux                 As Integer
    Dim PasoDeNivel         As Boolean
    Dim experienceToLevelUp As Long
    ' Randomizo las vidas
    Randomize Time
    With UserList(UserIndex)
        WasNewbie = EsNewbie(UserIndex)
        experienceToLevelUp = ExpLevelUp(.Stats.ELV)
        Do While .Stats.Exp >= experienceToLevelUp And .Stats.ELV < STAT_MAXELV
            'Store it!
            'Call Statistics.UserLevelUp(UserIndex)
            UserList(UserIndex).Counters.timeFx = 3
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 106, 0, .pos.x, .pos.y))
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .pos.x, .pos.y))
            .Stats.Exp = .Stats.Exp - experienceToLevelUp
            Pts = Pts + ModClase(.clase).LevelSkillPoints
            .Stats.ELV = .Stats.ELV + 1
            experienceToLevelUp = ExpLevelUp(.Stats.ELV)
            AumentoSta = .Stats.MaxSta
            AumentoMANA = .Stats.MaxMAN
            AumentoHIT = .Stats.MaxHit
            .Stats.MaxMAN = UserMod.GetMaxMana(UserIndex)
            .Stats.MaxSta = UserMod.GetMaxStamina(UserIndex)
            .Stats.MinHIT = UserMod.GetHitModifier(UserIndex) + 1
            .Stats.MaxHit = UserMod.GetHitModifier(UserIndex) + 2
            AumentoSta = .Stats.MaxSta - AumentoSta
            AumentoMANA = .Stats.MaxMAN - AumentoMANA
            AumentoHIT = .Stats.MaxHit - AumentoHIT
            ' Shugar 15/7/2024
            ' Devuelvo el aumento de vida variable pero con capeo min/max
            ' Promedio sin vida variable
            PromClaseRaza = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(e_Atributos.Constitucion)) * 0.5
            ' Promedio real del personaje
            PromPersonaje = CalcularPromedioVida(UserIndex)
            ' Sesgo a favor del promedio sin vida variable
            ' Si DesbalancePromedioVidas = 0, el PromBias es el PromClaseRaza del manual
            PromBias = PromClaseRaza + (PromClaseRaza - PromPersonaje) * DesbalancePromedioVidas
            ' Aumenta la vida un número entero al azar en un rango dado
            ' Min: PromClaseRaza - RangoVidas
            ' Max: PromClaseRaza + RangoVidas
            ' Media: PromBias
            ' Desviación: InfluenciaPromedioVidas
            AumentoHP = RandomIntBiased(PromClaseRaza - RangoVidas, PromClaseRaza + RangoVidas, PromBias, InfluenciaPromedioVidas)
            ' Capeo de vida máxima a +10
            If .Stats.MaxHp + AumentoHP > UserMod.GetMaxHp(UserIndex) + CapVidaMax Then
                AumentoHP = (UserMod.GetMaxHp(UserIndex) + CapVidaMax) - .Stats.MaxHp
            End If
            ' Capeo de vida mínima a -10
            If .Stats.MaxHp + AumentoHP < UserMod.GetMaxHp(UserIndex) + CapVidaMin Then
                AumentoHP = (UserMod.GetMaxHp(UserIndex) + CapVidaMin) - .Stats.MaxHp
            End If
            ' Aumento la vida máxima del personaje
            .Stats.MaxHp = .Stats.MaxHp + AumentoHP
            'Notificamos al user
            'Msg186=¡Has subido al nivel ¬1!
            Call WriteLocaleMsg(UserIndex, 186, e_FontTypeNames.FONTTYPE_INFO, .Stats.ELV)
            If AumentoHP > 0 Then
                'Msg197=Has ganado ¬1 puntos de vida. Tu vida actual es: ¬2
                Call WriteLocaleMsg(UserIndex, 197, e_FontTypeNames.FONTTYPE_INFO, AumentoHP & "¬" & .Stats.MaxHp)
            End If
            If AumentoSta > 0 Then
                'Msg198=Has ganado ¬1 puntos de energía. Tu energía actual es: ¬2
                Call WriteLocaleMsg(UserIndex, 198, e_FontTypeNames.FONTTYPE_INFO, AumentoSta & "¬" & .Stats.MaxSta)
            End If
            If AumentoMANA > 0 Then
                'Msg199=Has ganado ¬1 puntos de maná. Tu maná actual es: ¬2
                Call WriteLocaleMsg(UserIndex, 199, e_FontTypeNames.FONTTYPE_INFO, AumentoMANA & "¬" & .Stats.MaxMAN)
            End If
            If AumentoHIT > 0 Then
                'Msg200=Tu golpe mínimo y máximo aumentaron en ¬1 puntos. Tus daños actuales son ¬2 / ¬3
                Call WriteLocaleMsg(UserIndex, 200, e_FontTypeNames.FONTTYPE_INFO, AumentoHIT & "¬" & .Stats.MinHIT & "¬" & .Stats.MaxHit)
            End If
            PasoDeNivel = True
            .Stats.MinHp = .Stats.MaxHp
            ' Call UpdateUserInv(True, UserIndex, 0)
            If SvrConfig.GetValue("OroPorNivel") > 0 Then
                If EsNewbie(UserIndex) Then
                    Dim OroRecompenza As Long
                    OroRecompenza = SvrConfig.GetValue("OroPorNivel") * .Stats.ELV * SvrConfig.GetValue("GoldMult")
                    .Stats.GLD = .Stats.GLD + OroRecompenza
                    'Msg1293= Has ganado ¬1 monedas de oro.
                    Call WriteLocaleMsg(UserIndex, 1293, e_FontTypeNames.FONTTYPE_INFO, OroRecompenza)
                End If
            End If
        Loop
        If PasoDeNivel Then
            If .Stats.ELV >= STAT_MAXELV Then .Stats.Exp = 0
            Call UpdateUserInv(True, UserIndex, 0)
            'Call CheckearRecompesas(UserIndex, 3)
            Call WriteUpdateUserStats(UserIndex)
            If Pts > 0 Then
                .Stats.SkillPts = .Stats.SkillPts + Pts
                Call WriteLevelUp(UserIndex, .Stats.SkillPts)
                'Msg187=¡Has ganado ¬1 skillpoints! Dispones de ¬2 puntos libres, se cuidadoso al momento de usarlos.
                Call WriteLocaleMsg(UserIndex, 187, e_FontTypeNames.FONTTYPE_INFO, Pts & "¬" & .Stats.SkillPts)
            End If
            If Not EsNewbie(UserIndex) And WasNewbie Then
                Call QuitarNewbieObj(UserIndex)
            ElseIf .Stats.ELV >= MapInfo(.pos.Map).MaxLevel And Not EsGM(UserIndex) Then
                If MapInfo(.pos.Map).Salida.Map <> 0 Then
                    ' Msg523=Tu nivel no te permite seguir en el mapa.
                    Call WriteLocaleMsg(UserIndex, 523, e_FontTypeNames.FONTTYPE_INFO)
                    Call WarpUserChar(UserIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
                End If
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.Description)
End Sub

Public Sub SwapTargetUserPos(ByVal TargetUser As Integer, ByRef NewTargetPos As t_WorldPos)
    Dim Heading As e_Heading
    Heading = UserList(TargetUser).Char.Heading
    UserList(TargetUser).pos = NewTargetPos
    Call WritePosUpdate(TargetUser)
    If UserList(TargetUser).flags.AdminInvisible = 0 Then
        Call SendData(SendTarget.ToPCAreaButIndex, TargetUser, PrepareMessageCharacterMove(UserList(TargetUser).Char.charindex, UserList(TargetUser).pos.x, UserList( _
                TargetUser).pos.y), True)
    Else
        Call SendData(SendTarget.ToAdminAreaButIndex, TargetUser, PrepareMessageCharacterMove(UserList(TargetUser).Char.charindex, UserList(TargetUser).pos.x, UserList( _
                TargetUser).pos.y))
    End If
    If IsValidUserRef(UserList(TargetUser).flags.GMMeSigue) Then
        Call WriteForceCharMoveSiguiendo(UserList(TargetUser).flags.GMMeSigue.ArrayIndex, Heading)
    End If
    Call WriteForceCharMove(TargetUser, Heading)
    'Update map and char
    UserList(TargetUser).Char.Heading = Heading
    MapData(UserList(TargetUser).pos.Map, UserList(TargetUser).pos.x, UserList(TargetUser).pos.y).UserIndex = TargetUser
    'Actualizamos las areas de ser necesario
    Call ModAreas.CheckUpdateNeededUser(TargetUser, Heading, 0)
    Call CancelarComercioUsuario(TargetUser, 1844)
End Sub

Function TranslateUserPos(ByVal UserIndex As Integer, ByRef NewPos As t_WorldPos, ByVal Speed As Long)
    On Error GoTo TranslateUserPos_Err
    Dim OriginalPos As t_WorldPos
    With UserList(UserIndex)
        OriginalPos = .pos
        If MapInfo(.pos.Map).NumUsers > 1 Or IsValidUserRef(.flags.GMMeSigue) Then
            If MapData(NewPos.Map, NewPos.x, NewPos.y).UserIndex > 0 Then
                Call SwapTargetUserPos(MapData(NewPos.Map, NewPos.x, NewPos.y).UserIndex, .pos)
            End If
        End If
        If .flags.AdminInvisible = 0 Then
            If IsValidUserRef(.flags.GMMeSigue) Then
                Call SendData(SendTarget.ToPCAreaButFollowerAndIndex, UserIndex, PrepareCharacterTranslate(.Char.charindex, NewPos.x, NewPos.y, Speed))
                Call WriteForceCharMoveSiguiendo(.flags.GMMeSigue.ArrayIndex, .Char.Heading)
            Else
                'Mando a todos menos a mi donde estoy
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCharacterTranslate(.Char.charindex, NewPos.x, NewPos.y, Speed), True)
            End If
        Else
            Call SendData(SendTarget.ToAdminAreaButIndex, UserIndex, PrepareCharacterTranslate(.Char.charindex, NewPos.x, NewPos.y, Speed))
        End If
        'Update map and user pos
        If MapData(.pos.Map, .pos.x, .pos.y).UserIndex = UserIndex Then
            MapData(.pos.Map, .pos.x, .pos.y).UserIndex = 0
        End If
        .pos = NewPos
        MapData(.pos.Map, .pos.x, .pos.y).UserIndex = UserIndex
        Call WritePosUpdate(UserIndex)
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(UserIndex, .Char.Heading, 0)
        If .Counters.Trabajando Then
            Call WriteMacroTrabajoToggle(UserIndex, False)
        End If
    End With
    Call CancelarComercioUsuario(UserIndex, 1844)
    Exit Function
TranslateUserPos_Err:
    Call LogError("Error en la subrutina TranslateUserPos - Error : " & Err.Number & " - Description : " & Err.Description)
End Function

Public Sub SwapNpcPos(ByVal UserIndex As Integer, ByRef TargetPos As t_WorldPos, ByVal nHeading As e_Heading)
    Dim NpcIndex         As Integer
    Dim Opposite_Heading As e_Heading
    NpcIndex = MapData(TargetPos.Map, TargetPos.x, TargetPos.y).NpcIndex
    If NpcIndex <= 0 Then Exit Sub
    Opposite_Heading = InvertHeading(nHeading)
    Call HeadtoPos(Opposite_Heading, NpcList(NpcIndex).pos)
    Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageCharacterMove(NpcList(NpcIndex).Char.charindex, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y), False)
    MapData(NpcList(NpcIndex).pos.Map, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y).NpcIndex = NpcIndex
    MapData(TargetPos.Map, TargetPos.x, TargetPos.y).NpcIndex = 0
    Call CheckUpdateNeededNpc(NpcIndex, Opposite_Heading)
End Sub

Function MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As e_Heading) As Boolean
    ' Lo convierto a función y saco los WritePosUpdate, ahora están en el paquete
    On Error GoTo MoveUserChar_Err
    Dim nPos             As t_WorldPos
    Dim nPosOriginal     As t_WorldPos
    Dim nPosMuerto       As t_WorldPos
    Dim IndexMover       As Integer
    Dim Opposite_Heading As e_Heading
    With UserList(UserIndex)
        nPos = .pos
        Call HeadtoPos(nHeading, nPos)
        If Not LegalWalk(.pos.Map, nPos.x, nPos.y, nHeading, .flags.Navegando = 1, .flags.Navegando = 0, .flags.Montado, , UserIndex) Then
            Exit Function
        End If
        If .flags.Navegando And .invent.EquippedShipObjIndex = iObjTraje And Not (MapData(.pos.Map, nPos.x, nPos.y).trigger = e_Trigger.DETALLEAGUA Or MapData(.pos.Map, nPos.x, _
                nPos.y).trigger = e_Trigger.NADOCOMBINADO Or MapData(.pos.Map, nPos.x, nPos.y).trigger = e_Trigger.VALIDONADO Or MapData(.pos.Map, nPos.x, nPos.y).trigger = _
                e_Trigger.NADOBAJOTECHO) Then
            Exit Function
        End If
        If .Accion.AccionPendiente = True Then
            .Counters.TimerBarra = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, .Accion.Particula, .Counters.TimerBarra, True))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.charindex, .Counters.TimerBarra, e_AccionBarra.CancelarAccion))
            .Accion.AccionPendiente = False
            .Accion.Particula = 0
            .Accion.TipoAccion = e_AccionBarra.CancelarAccion
            .Accion.HechizoPendiente = 0
            .Accion.RunaObj = 0
            .Accion.ObjSlot = 0
            .Accion.AccionPendiente = False
        End If
        Call SwapNpcPos(UserIndex, nPos, nHeading)
        'Si no estoy solo en el mapa...
        If MapInfo(.pos.Map).NumUsers > 1 Or IsValidUserRef(.flags.GMMeSigue) Then
            ' Intercambia posición si hay un casper o gm invisible
            IndexMover = MapData(nPos.Map, nPos.x, nPos.y).UserIndex
            If IndexMover <> 0 Then
                ' Sólo puedo patear caspers/gms invisibles si no es él un gm invisible
                ' If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
                Call WritePosUpdate(IndexMover)
                Opposite_Heading = InvertHeading(nHeading)
                Call HeadtoPos(Opposite_Heading, UserList(IndexMover).pos)
                ' Si es un admin invisible, no se avisa a los demas clientes
                If UserList(IndexMover).flags.AdminInvisible = 0 Then
                    Call SendData(SendTarget.ToPCAreaButIndex, IndexMover, PrepareMessageCharacterMove(UserList(IndexMover).Char.charindex, UserList(IndexMover).pos.x, UserList( _
                            IndexMover).pos.y), True)
                Else
                    Call SendData(SendTarget.ToAdminAreaButIndex, IndexMover, PrepareMessageCharacterMove(UserList(IndexMover).Char.charindex, UserList(IndexMover).pos.x, _
                            UserList(IndexMover).pos.y))
                End If
                If IsValidUserRef(UserList(IndexMover).flags.GMMeSigue) Then
                    Call WriteForceCharMoveSiguiendo(UserList(IndexMover).flags.GMMeSigue.ArrayIndex, Opposite_Heading)
                End If
                Call WriteForceCharMove(IndexMover, Opposite_Heading)
                'Update map and char
                UserList(IndexMover).Char.Heading = Opposite_Heading
                MapData(UserList(IndexMover).pos.Map, UserList(IndexMover).pos.x, UserList(IndexMover).pos.y).UserIndex = IndexMover
                'Actualizamos las areas de ser necesario
                Call ModAreas.CheckUpdateNeededUser(IndexMover, Opposite_Heading, 0)
            End If
            If .flags.AdminInvisible = 0 Then
                If IsValidUserRef(.flags.GMMeSigue) Then
                    Call SendData(SendTarget.ToPCAreaButFollowerAndIndex, UserIndex, PrepareMessageCharacterMove(.Char.charindex, nPos.x, nPos.y))
                    Call WriteForceCharMoveSiguiendo(.flags.GMMeSigue.ArrayIndex, nHeading)
                Else
                    'Mando a todos menos a mi donde estoy
                    Call SendData(SendTarget.ToPCAliveAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.charindex, nPos.x, nPos.y), True)
                    Dim LoopC     As Integer
                    Dim tempIndex As Integer
                    'Togle para alternar el paso para los invis
                    .flags.stepToggle = Not .flags.stepToggle
                    If Not EsGM(UserIndex) Then
                        If .flags.invisible + .flags.Oculto > 0 And .flags.Navegando = 0 Then
                            For LoopC = 1 To ConnGroups(UserList(UserIndex).pos.Map).CountEntrys
                                tempIndex = ConnGroups(UserList(UserIndex).pos.Map).UserEntrys(LoopC)
                                If tempIndex <> UserIndex And Not EsGM(tempIndex) Then
                                    If Abs(nPos.x - UserList(tempIndex).pos.x) <= RANGO_VISION_X And Abs(nPos.y - UserList(tempIndex).pos.y) <= RANGO_VISION_Y Then
                                        If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                                            If UserList(tempIndex).flags.Muerto = 0 Or MapInfo(UserList(tempIndex).pos.Map).Seguro = 1 Then
                                                If Not CheckGuildSend(UserList(UserIndex), UserList(tempIndex)) Then
                                                    If .Counters.timeFx + .Counters.timeChat = 0 Then
                                                        If Distancia(nPos, UserList(tempIndex).pos) > DISTANCIA_ENVIO_DATOS Then
                                                            'Mandamos los pasos para los pjs q estan lejos para que simule que caminen.
                                                            'Mando tambien el char para q lo borre
                                                            Call WritePlayWaveStep(tempIndex, .Char.charindex, MapData(nPos.Map, nPos.x, nPos.y).Graphic(1), MapData(nPos.Map, _
                                                                    nPos.x, nPos.y).Graphic(2), Distance(nPos.x, nPos.y, UserList(tempIndex).pos.x, UserList(tempIndex).pos.y), _
                                                                    Sgn(nPos.x - UserList(tempIndex).pos.x), .flags.stepToggle)
                                                        Else
                                                            Call WritePosUpdateChar(tempIndex, nPos.x, nPos.y, .Char.charindex)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next LoopC
                        End If
                        Dim x As Byte, y As Byte
                        'Esto es para q si me acerco a un usuario que esta invisible y no se mueve me notifique su posicion
                        For x = nPos.x - DISTANCIA_ENVIO_DATOS To nPos.x + DISTANCIA_ENVIO_DATOS
                            For y = nPos.y - DISTANCIA_ENVIO_DATOS To nPos.y + DISTANCIA_ENVIO_DATOS
                                tempIndex = MapData(.pos.Map, x, y).UserIndex
                                If tempIndex > 0 And tempIndex <> UserIndex And Not EsGM(tempIndex) Then
                                    If UserList(tempIndex).flags.invisible + UserList(tempIndex).flags.Oculto > 0 And UserList(tempIndex).flags.Navegando = 0 And (.GuildIndex = _
                                            0 Or .GuildIndex <> UserList(tempIndex).GuildIndex Or modGuilds.NivelDeClan(.GuildIndex) < 6) Then
                                        Call WritePosUpdateChar(UserIndex, x, y, UserList(tempIndex).Char.charindex)
                                    End If
                                End If
                            Next y
                        Next x
                    End If
                End If
            Else
                Call SendData(SendTarget.ToAdminAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.charindex, nPos.x, nPos.y))
            End If
        End If
        'Update map and user pos
        If MapData(.pos.Map, .pos.x, .pos.y).UserIndex = UserIndex Then
            MapData(.pos.Map, .pos.x, .pos.y).UserIndex = 0
        End If
        .pos = nPos
        .Char.Heading = nHeading
        MapData(.pos.Map, .pos.x, .pos.y).UserIndex = UserIndex
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading, 0)
        If .Counters.Trabajando Then
            Call WriteMacroTrabajoToggle(UserIndex, False)
        End If
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
    End With
    Call CancelarComercioUsuario(UserIndex, 1844)
    MoveUserChar = True
    Exit Function
MoveUserChar_Err:
    Call TraceError(Err.Number, Err.Description + " UI:" + UserIndex, "UsUaRiOs.MoveUserChar", Erl)
End Function

Public Function InvertHeading(ByVal nHeading As e_Heading) As e_Heading
    On Error GoTo InvertHeading_Err
    '*************************************************
    'Author: ZaMa
    'Last modified: 30/03/2009
    'Returns the heading opposite to the one passed by val.
    '*************************************************
    Select Case nHeading
        Case e_Heading.EAST
            InvertHeading = e_Heading.WEST
        Case e_Heading.WEST
            InvertHeading = e_Heading.EAST
        Case e_Heading.SOUTH
            InvertHeading = e_Heading.NORTH
        Case e_Heading.NORTH
            InvertHeading = e_Heading.SOUTH
    End Select
    Exit Function
InvertHeading_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.InvertHeading", Erl)
End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Object As t_UserOBJ)
    On Error GoTo ChangeUserInv_Err
    UserList(UserIndex).invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(UserIndex, Slot)
    Exit Sub
ChangeUserInv_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ChangeUserInv", Erl)
End Sub

Function NextOpenCharIndex() As Integer
    On Error GoTo NextOpenCharIndex_Err
    Dim LoopC As Long
    For LoopC = 1 To MAXCHARS
        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            If LoopC > LastChar Then LastChar = LoopC
            Exit Function
        End If
    Next LoopC
    Exit Function
NextOpenCharIndex_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.NextOpenCharIndex", Erl)
End Function

Function NextOpenUser() As Integer
    On Error GoTo NextOpenUser_Err
    Dim LoopC As Long
    If IsFeatureEnabled("use_old_user_slot_check") Then
        For LoopC = 1 To MaxUsers + 1
            If LoopC > MaxUsers Then Exit For
            If (Not UserList(LoopC).ConnectionDetails.ConnIDValida And UserList(LoopC).flags.UserLogged = False) Then Exit For
        Next LoopC
        NextOpenUser = LoopC
    Else
        NextOpenUser = GetNextAvailableUserSlot
    End If
    Exit Function
NextOpenUser_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.NextOpenUser", Erl)
End Function

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
    On Error GoTo SendUserStatsTxt_Err
    Dim GuildI As Integer
    'Msg1295= Estadisticas de: ¬1
    Call WriteLocaleMsg(sendIndex, "1295", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).name)
    Call WriteConsoleMsg(sendIndex, PrepareMessageLocaleMsg(1857, UserList(UserIndex).Stats.ELV & "¬" & UserList(UserIndex).Stats.Exp & "¬" & ExpLevelUp(UserList( _
            UserIndex).Stats.ELV), e_FontTypeNames.FONTTYPE_INFO)) ' Msg1857=Nivel: ¬1  EXP: ¬2/¬3
    Call WriteConsoleMsg(sendIndex, PrepareMessageLocaleMsg(1858, UserList(UserIndex).Stats.MinHp & "¬" & UserList(UserIndex).Stats.MaxHp & "¬" & UserList( _
            UserIndex).Stats.MinMAN & "¬" & UserList(UserIndex).Stats.MaxMAN & "¬" & UserList(UserIndex).Stats.MinSta & "¬" & UserList(UserIndex).Stats.MaxSta, _
            e_FontTypeNames.FONTTYPE_INFO)) ' Msg1858=Salud: ¬1/¬2  Mana: ¬3/¬4  Vitalidad: ¬5/¬6
    If UserList(UserIndex).invent.EquippedWeaponObjIndex > 0 Then
        Call WriteConsoleMsg(sendIndex, PrepareMessageLocaleMsg(1859, UserList(UserIndex).Stats.MinHIT & "¬" & UserList(UserIndex).Stats.MaxHit & "¬" & ObjData(UserList( _
                UserIndex).invent.EquippedWeaponObjIndex).MinHIT & "¬" & ObjData(UserList(UserIndex).invent.EquippedWeaponObjIndex).MaxHit, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1859=Menor Golpe/Mayor Golpe: ¬1/¬2 (¬3/¬4)
    Else
        Call WriteConsoleMsg(sendIndex, PrepareMessageLocaleMsg(1860, UserList(UserIndex).Stats.MinHIT & "¬" & UserList(UserIndex).Stats.MaxHit, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1860=Menor Golpe/Mayor Golpe: ¬1/¬2
    End If
    If UserList(UserIndex).invent.EquippedArmorObjIndex > 0 Then
        If UserList(UserIndex).invent.EquippedShieldObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, PrepareMessageLocaleMsg(1861, ObjData(UserList(UserIndex).invent.EquippedArmorObjIndex).MinDef + ObjData(UserList( _
                    UserIndex).invent.EquippedShieldObjIndex).MinDef & "¬" & ObjData(UserList(UserIndex).invent.EquippedArmorObjIndex).MaxDef + ObjData(UserList( _
                    UserIndex).invent.EquippedShieldObjIndex).MaxDef, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1861=(CUERPO) Min Def/Max Def: ¬1/¬2
        Else
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).invent.EquippedArmorObjIndex).MinDef & "/" & ObjData(UserList( _
                    UserIndex).invent.EquippedArmorObjIndex).MaxDef, e_FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        'Msg1098= (CUERPO) Min Def/Max Def: 0
        Call WriteLocaleMsg(sendIndex, "1098", e_FontTypeNames.FONTTYPE_INFO)
    End If
    If UserList(UserIndex).invent.EquippedHelmetObjIndex > 0 Then
        Call WriteConsoleMsg(sendIndex, PrepareMessageLocaleMsg(1862, ObjData(UserList(UserIndex).invent.EquippedHelmetObjIndex).MinDef & "¬" & ObjData(UserList( _
                UserIndex).invent.EquippedHelmetObjIndex).MaxDef, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1862=(CABEZA) Min Def/Max Def: ¬1/¬2
    Else
        'Msg1099= (CABEZA) Min Def/Max Def: 0
        Call WriteLocaleMsg(sendIndex, "1099", e_FontTypeNames.FONTTYPE_INFO)
    End If
    GuildI = UserList(UserIndex).GuildIndex
    If GuildI > 0 Then
        'Msg1296= Clan: ¬1
        Call WriteLocaleMsg(sendIndex, "1296", e_FontTypeNames.FONTTYPE_INFO, modGuilds.GuildName(GuildI))
        If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(UserList(sendIndex).name) Then
            'Msg1100= Status: Líder
            Call WriteLocaleMsg(sendIndex, "1100", e_FontTypeNames.FONTTYPE_INFO)
        End If
        'guildpts no tienen objeto
    End If
    #If ConUpTime Then
        Dim TempDate As Date
        Dim TempSecs As Long
        Dim TempStr  As String
        TempDate = Now - UserList(UserIndex).LogOnTime
        TempSecs = (UserList(UserIndex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
        TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod _
                3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, PrepareMessageLocaleMsg(1863, Hour(TempDate) & "¬" & Minute(TempDate) & "¬" & Second(TempDate), e_FontTypeNames.FONTTYPE_INFO)) ' Msg1863=Logeado hace: ¬1:¬2:¬3
        'Msg1297= Total: ¬1
        Call WriteLocaleMsg(sendIndex, "1297", e_FontTypeNames.FONTTYPE_INFO, TempStr)
    #End If
    Call LoadPatronCreditsFromDB(UserIndex)
    'Msg1298= Oro: ¬1
    Call WriteLocaleMsg(sendIndex, "1298", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).Stats.GLD)
    'Msg1299= Veces que Moriste: ¬1
    Call WriteLocaleMsg(sendIndex, "1299", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).flags.VecesQueMoriste)
    Call WriteLocaleMsg(sendIndex, MsgFactionScore, e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).Faccion.FactionScore)
    'Msg1300= Creditos Patreon: ¬1
    Call WriteLocaleMsg(sendIndex, "1300", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).Stats.Creditos)
    'Msg2078 = Nivel de Jinete:¬1
    Call WriteLocaleMsg(sendIndex, MSG_RIDER_LEVEL_REQUIREMENT, e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).Stats.JineteLevel)

' ========================
' Show current home
' ========================
Dim char_home As String
Select Case UserList(UserIndex).Hogar
    Case e_Ciudad.cUllathorpe: char_home = CIUDAD_ULLATHORPE
    Case e_Ciudad.cNix: char_home = CIUDAD_NIX
    Case e_Ciudad.cBanderbill: char_home = CIUDAD_BANDERBILL
    Case e_Ciudad.cLindos: char_home = CIUDAD_LINDOS
    Case e_Ciudad.cArghal: char_home = CIUDAD_ARGHAL
    Case e_Ciudad.cForgat: char_home = CIUDAD_FORGAT
    Case e_Ciudad.cArkhein: char_home = CIUDAD_ARKHEIN
    Case e_Ciudad.cEldoria: char_home = CIUDAD_ELDORIA
    Case e_Ciudad.cPenthar: char_home = CIUDAD_PENTHAR
    Case Else: char_home = CIUDAD_ULLATHORPE
End Select
    Call WriteLocaleMsg(sendIndex, MSG_CHARACTER_HOME, e_FontTypeNames.FONTTYPE_INFO, char_home)



Exit Sub
SendUserStatsTxt_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserStatsTxt", Erl)
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
    On Error GoTo SendUserMiniStatsTxt_Err
    '*************************************************
    'Author: Unknown
    'Last modified: 23/01/2007
    'Shows the users Stats when the user is online.
    '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
    '*************************************************
    With UserList(UserIndex)
        'Msg1301= Pj: ¬1
        Call WriteLocaleMsg(sendIndex, "1301", e_FontTypeNames.FONTTYPE_INFO, .name)
        'Msg1302= Ciudadanos Matados: ¬1
        Call WriteLocaleMsg(sendIndex, "1302", e_FontTypeNames.FONTTYPE_INFO, .Faccion.ciudadanosMatados)
        'Msg1303= Criminales Matados: ¬1
        Call WriteLocaleMsg(sendIndex, "1303", e_FontTypeNames.FONTTYPE_INFO, .Faccion.CriminalesMatados)
        'Msg1304= UsuariosMatados: ¬1
        Call WriteLocaleMsg(sendIndex, "1304", e_FontTypeNames.FONTTYPE_INFO, .Stats.UsuariosMatados)
        'Msg1305= NPCsMuertos: ¬1
        Call WriteLocaleMsg(sendIndex, "1305", e_FontTypeNames.FONTTYPE_INFO, .Stats.NPCsMuertos)
        'Msg1306= Clase: ¬1
        Call WriteLocaleMsg(sendIndex, "1306", e_FontTypeNames.FONTTYPE_INFO, ListaClases(.clase))
        'Msg1307= Pena: ¬1
        Call WriteLocaleMsg(sendIndex, "1307", e_FontTypeNames.FONTTYPE_INFO, .Counters.Pena)
        If .GuildIndex > 0 Then
            'Msg1308= Clan: ¬1
            Call WriteLocaleMsg(sendIndex, "1308", e_FontTypeNames.FONTTYPE_INFO, GuildName(.GuildIndex))
        End If
        'Msg1309= Oro en billetera: ¬1
        Call WriteLocaleMsg(sendIndex, "1309", e_FontTypeNames.FONTTYPE_INFO, .Stats.GLD)
        'Msg1310= Oro en banco: ¬1
        Call WriteLocaleMsg(sendIndex, "1310", e_FontTypeNames.FONTTYPE_INFO, .Stats.Banco)
    End With
    Exit Sub
SendUserMiniStatsTxt_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserMiniStatsTxt", Erl)
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
    On Error GoTo SendUserInvTxt_Err
    Dim j As Long
    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)
    'Msg1311= Tiene ¬1 objetos.
    Call WriteLocaleMsg(sendIndex, "1311", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).invent.NroItems)
    For j = 1 To UserList(UserIndex).CurrentInventorySlots
        If UserList(UserIndex).invent.Object(j).ObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, PrepareMessageLocaleMsg(1865, j & "¬" & ObjData(UserList(UserIndex).invent.Object(j).ObjIndex).name & "¬" & UserList( _
                    UserIndex).invent.Object(j).amount, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1865= Objeto ¬1 ¬2 Cantidad:¬3
        End If
    Next j
    Exit Sub
SendUserInvTxt_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserInvTxt", Erl)
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
    On Error GoTo SendUserSkillsTxt_Err
    Dim j As Integer
    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)
    For j = 1 To NUMSKILLS
        Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), e_FontTypeNames.FONTTYPE_INFO)
    Next
    'Msg1312=  SkillLibres:¬1
    Call WriteLocaleMsg(sendIndex, "1312", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).Stats.SkillPts)
    Exit Sub
SendUserSkillsTxt_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserSkillsTxt", Erl)
End Sub

Function DameUserIndexConNombre(ByVal nombre As String) As Integer
    On Error GoTo DameUserIndexConNombre_Err
    Dim LoopC As Integer
    LoopC = 1
    nombre = UCase$(nombre)
    Do Until UCase$(UserList(LoopC).name) = nombre
        LoopC = LoopC + 1
        If LoopC > MaxUsers Then
            DameUserIndexConNombre = 0
            Exit Function
        End If
    Loop
    DameUserIndexConNombre = LoopC
    Exit Function
DameUserIndexConNombre_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.DameUserIndexConNombre", Erl)
End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal AffectsOwner As Boolean = True)
    On Error GoTo NPCAtacado_Err
    '  El usuario pierde la protección
    UserList(UserIndex).Counters.TiempoDeInmunidad = 0
    UserList(UserIndex).flags.Inmunidad = 0
    'Guardamos el usuario que ataco el npc.
    If Not IsSet(NpcList(NpcIndex).flags.StatusMask, eTaunted) And NpcList(NpcIndex).Movement <> Estatico And NpcList(NpcIndex).flags.AttackedFirstBy = vbNullString Then
        Call SetUserRef(NpcList(NpcIndex).TargetUser, UserIndex)
        NpcList(NpcIndex).Hostile = 1
        If AffectsOwner Then
            NpcList(NpcIndex).flags.AttackedBy = UserList(UserIndex).name
            NpcList(NpcIndex).flags.AttackedTime = GlobalFrameTime
        End If
    End If
    'Guarda el NPC que estas atacando ahora.
    If AffectsOwner Then Call SetNpcRef(UserList(UserIndex).flags.NPCAtacado, NpcIndex)
    If NpcList(NpcIndex).flags.Faccion = Armada And Status(UserIndex) = e_Facciones.Ciudadano Then
        Call VolverCriminal(UserIndex)
    End If
    If IsValidUserRef(NpcList(NpcIndex).MaestroUser) And NpcList(NpcIndex).MaestroUser.ArrayIndex <> UserIndex Then
        Call AllMascotasAtacanUser(UserIndex, NpcList(NpcIndex).MaestroUser.ArrayIndex)
    End If
    Exit Sub
NPCAtacado_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.NPCAtacado", Erl)
End Sub

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)
    On Error GoTo SubirSkill_Err
    Dim Lvl As Integer, maxPermitido As Integer
    Lvl = UserList(UserIndex).Stats.ELV
    If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then
        Exit Sub
    End If
    ' Se suben 5 skills cada dos niveles como máximo.
    If (Lvl Mod 2 = 0) Then ' El level es numero par
        maxPermitido = (Lvl \ 2) * 5
    Else ' El level es numero impar
        ' Esta cuenta signifca, que si el nivel anterior terminaba en 5 ahora
        ' suma dos puntos mas, sino 3. Lo de siempre.
        maxPermitido = (Lvl \ 2) * 5 + 3 - (((((Lvl - 1) \ 2) * 5) Mod 10) \ 5)
    End If
    If UserList(UserIndex).Stats.UserSkills(Skill) >= maxPermitido Then
        Exit Sub
    End If
    If UserList(UserIndex).Stats.MinHam = 0 Or UserList(UserIndex).Stats.MinAGU = 0 Then
        Exit Sub
    End If
    Dim Aumenta As Integer
    Dim Prob    As Integer
    'Cuadratic expression to sumarize old select case lvl bands
    Prob = Int(0.1 * (Lvl ^ 2) + 15)
    Aumenta = RandomNumber(1, Prob * DificultadSubirSkill)
    Dim cutoff As Integer
    If UserList(UserIndex).flags.PendienteDelExperto = 1 Then
        cutoff = EXPERT_SKILL_CUTOFF   ' Expert used to allow 1..14
    Else
        cutoff = NONEXPERT_SKILL_CUTOFF   ' Non-expert used to allow 1..9
    End If
    If Aumenta >= cutoff Then Exit Sub
    UserList(UserIndex).Stats.UserSkills(Skill) = UserList(UserIndex).Stats.UserSkills(Skill) + 1
    Call WriteLocaleMsg(UserIndex, 1626, e_FontTypeNames.FONTTYPE_INFO, SkillsNames(Skill) & "¬" & UserList(UserIndex).Stats.UserSkills(Skill))
    Dim BonusExp As Long
    BonusExp = 5& * SvrConfig.GetValue("ExpMult")
    If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + BonusExp
        If UserList(UserIndex).Stats.Exp > MAXEXP Then
            UserList(UserIndex).Stats.Exp = MAXEXP
        End If
        UserList(UserIndex).flags.ModificoSkills = True
        If UserList(UserIndex).ChatCombate = 1 Then
            Call WriteLocaleMsg(UserIndex, 140, e_FontTypeNames.FONTTYPE_EXP, BonusExp) 'Msg140=Has ganado ¬1 puntos de experiencia.
        End If
        Call WriteUpdateExp(UserIndex)
        Call CheckUserLevel(UserIndex)
    End If
    Exit Sub
SubirSkill_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SubirSkill", Erl)
End Sub

Public Sub SubirSkillDeArmaActual(ByVal UserIndex As Integer)
    On Error GoTo SubirSkillDeArmaActual_Err
    With UserList(UserIndex)
        If .invent.EquippedWeaponObjIndex > 0 Then
            ' Arma con proyectiles, subimos armas a distancia
            If ObjData(.invent.EquippedWeaponObjIndex).Proyectil Then
                Call SubirSkill(UserIndex, e_Skill.Proyectiles)
            ElseIf ObjData(.invent.EquippedWeaponObjIndex).WeaponType = eKnuckle Then
                Call SubirSkill(UserIndex, e_Skill.Wrestling)
                ' Sino, subimos combate con armas
            Else
                Call SubirSkill(UserIndex, e_Skill.Armas)
            End If
            ' Si no está usando un arma, subimos combate sin armas
        Else
            Call SubirSkill(UserIndex, e_Skill.Wrestling)
        End If
    End With
    Exit Sub
SubirSkillDeArmaActual_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SubirSkillDeArmaActual", Erl)
End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'
Sub UserDie(ByVal UserIndex As Integer)
    '************************************************
    'Author: Uknown
    'Last Modified: 04/15/2008 (NicoNZ)
    'Ahora se resetea el counter del invi
    '************************************************
    On Error GoTo ErrorHandler
    Dim i  As Long
    Dim aN As Integer
    With UserList(UserIndex)
        .Counters.Mimetismo = 0
        .flags.Mimetizado = e_EstadoMimetismo.Desactivado
        Call RefreshCharStatus(UserIndex)
        'Sonido
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(IIf(.genero = e_Genero.Hombre, e_SoundIndex.MUERTE_HOMBRE, e_SoundIndex.MUERTE_MUJER), .pos.x, .pos.y))
        'Quitar el dialogo del user muerto
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.charindex))
        .Stats.MinHp = 0
        .Stats.MinSta = 0
        .Stats.shield = 0
        .flags.AtacadoPorUser = 0
        .flags.incinera = 0
        .flags.Paraliza = 0
        .flags.Envenena = 0
        .flags.Estupidiza = 0
        .flags.DivineBlood = 0
        Call ClearEffectList(.EffectOverTime, e_EffectType.eAny, True)
        Call ClearModifiers(.Modifiers)
        .flags.Muerto = 1
        Call WriteUpdateHP(UserIndex)
        Call WriteUpdateSta(UserIndex)
        Call ClearAttackerNpc(UserIndex)
        If MapData(.pos.Map, .pos.x, .pos.y).trigger <> e_Trigger.ZONAPELEA And MapInfo(.pos.Map).DropItems Then
            If (.flags.Privilegios And e_PlayerType.User) <> 0 Then
                If .flags.PendienteDelSacrificio = 0 Then
                    Call TirarTodosLosItems(UserIndex)
                Else
                    Dim MiObj As t_Obj
                    MiObj.amount = 1
                    MiObj.ObjIndex = PENDIENTE
                    Call QuitarObjetos(PENDIENTE, 1, UserIndex)
                End If
            End If
        End If
        Call Desequipar(UserIndex, .invent.EquippedArmorSlot)
        Call Desequipar(UserIndex, .invent.EquippedWeaponSlot)
        Call Desequipar(UserIndex, .invent.EquippedShieldSlot)
        Call Desequipar(UserIndex, .invent.EquippedHelmetSlot)
        Call Desequipar(UserIndex, .invent.EquippedRingAccesorySlot)
        Call Desequipar(UserIndex, .invent.EquippedWorkingToolSlot)
        Call Desequipar(UserIndex, .invent.EquippedSaddleSlot)
        Call Desequipar(UserIndex, .invent.EquippedMunitionSlot)
        Call Desequipar(UserIndex, .invent.EquippedAmuletAccesorySlot)
        Call Desequipar(UserIndex, .invent.EquippedRingAccesorySlot)
        'desequipar montura
        If .flags.Montado > 0 Then
            Call DoMontar(UserIndex, ObjData(.invent.EquippedSaddleObjIndex), .invent.EquippedSaddleSlot)
        End If
        ' << Reseteamos los posibles FX sobre el personaje >>
        If .Char.loops = INFINITE_LOOPS Then
            .Char.FX = 0
            .Char.loops = 0
        End If
        If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
            .flags.VecesQueMoriste = .flags.VecesQueMoriste + 1
        End If
        ' << Restauramos los atributos >>
        If .flags.TomoPocion Then
            For i = 1 To 4
                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
            Next i
            Call WriteFYA(UserIndex)
        End If
        ' << Frenamos el contador de la droga >>
        .flags.DuracionEfecto = 0
        '<< Cambiamos la apariencia del char >>
        If .flags.Navegando = 0 Then
            .Char.body = iCuerpoMuerto
            .Char.head = 0
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            .Char.CartAnim = NoCart
        Else
            Call EquiparBarco(UserIndex)
        End If
        Call ActualizarVelocidadDeUsuario(UserIndex)
        Call LimpiarEstadosAlterados(UserIndex)
        For i = 1 To MAXMASCOTAS
            If .MascotasIndex(i).ArrayIndex > 0 Then
                If IsValidNpcRef(.MascotasIndex(i)) Then
                    Call MuereNpc(.MascotasIndex(i).ArrayIndex, 0)
                Else
                    Call ClearNpcRef(.MascotasIndex(i))
                End If
            End If
        Next i
        If .clase = e_Class.Druid Then
            Dim Params() As Variant
            Dim ParamC   As Long
            ReDim Params(MAXMASCOTAS * 3 - 1)
            ParamC = 0
            For i = 1 To MAXMASCOTAS
                Params(ParamC) = .Id
                ParamC = ParamC + 1
                Params(ParamC) = i
                ParamC = ParamC + 1
                Params(ParamC) = 0
                ParamC = ParamC + 1
            Next i
            Call Execute(QUERY_UPSERT_PETS, Params)
        End If
        If (.flags.MascotasGuardadas = 0) Then
            .NroMascotas = 0
        End If
        '<< Actualizamos clientes >>
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NoCart, NoBackPack)
        If MapInfo(.pos.Map).Seguro = 0 Then
            ' Msg524=Escribe /HOGAR si deseas regresar rápido a tu hogar.
            Call WriteLocaleMsg(UserIndex, 524, e_FontTypeNames.FONTTYPE_New_Naranja)
        End If
        If .flags.EnReto Then
            Call MuereEnReto(UserIndex)
        End If
        If .flags.jugando_captura = 1 Then
            If Not InstanciaCaptura Is Nothing Then
                Call InstanciaCaptura.muereUsuario(UserIndex)
            End If
        End If
        'Borramos todos los personajes del area
        'HarThaoS: Mando un 5 en head para que cuente como muerto el area y no recalcule las posiciones.
        Call CheckUpdateNeededUser(UserIndex, 5, 0, .flags.Muerto)
        Dim LoopC     As Long
        Dim tempIndex As Integer
        Dim Map       As Integer
        Dim AreaX     As Integer
        Dim AreaY     As Integer
        AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
        AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY
        For LoopC = 1 To ConnGroups(UserList(UserIndex).pos.Map).CountEntrys
            tempIndex = ConnGroups(UserList(UserIndex).pos.Map).UserEntrys(LoopC)
            If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
                If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
                    If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
                        'Si no soy el que se murió
                        If UserIndex <> tempIndex And (Not EsGM(UserIndex)) And MapInfo(UserList(UserIndex).pos.Map).Seguro = 0 And UserList(tempIndex).flags.AdminInvisible = 1 _
                                Then
                            If UserList(UserIndex).GuildIndex = 0 Then
                                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageCharacterRemove(3, UserList(tempIndex).Char.charindex, True))
                            Else
                                If UserList(UserIndex).GuildIndex <> UserList(tempIndex).GuildIndex Then
                                    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageCharacterRemove(3, UserList(tempIndex).Char.charindex, True))
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next LoopC
    End With
    Exit Sub
ErrorHandler:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.UserDie", Erl)
End Sub

Public Function AlreadyKilledBy(ByVal TargetIndex As Integer, ByVal killerIndex As Integer) As Boolean
    Dim TargetPos As Integer
    With UserList(TargetIndex)
        TargetPos = Min(.flags.LastKillerIndex, MaxRecentKillToStore)
        Dim i As Integer
        For i = 0 To TargetPos
            If .flags.RecentKillers(i).UserId = UserList(killerIndex).Id And (GlobalFrameTime - .flags.RecentKillers(i).KillTime) < FactionReKillTime Then
                AlreadyKilledBy = True
                Exit Function
            End If
        Next i
    End With
End Function

Public Sub RegisterRecentKiller(ByVal TargetIndex As Integer, ByVal killerIndex As Integer)
    Dim InsertIndex As Integer
    With UserList(TargetIndex)
        InsertIndex = .flags.LastKillerIndex Mod MaxRecentKillToStore
        .flags.RecentKillers(InsertIndex).UserId = UserList(killerIndex).Id
        .flags.RecentKillers(InsertIndex).KillTime = GlobalFrameTime
        .flags.LastKillerIndex = .flags.LastKillerIndex + 1
        If .flags.LastKillerIndex > MaxRecentKillToStore * 10 Then 'prevent overflow
            .flags.LastKillerIndex = .flags.LastKillerIndex \ 10
        End If
    End With
End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
    On Error GoTo ContarMuerte_Err
    If EsNewbie(Muerto) Then Exit Sub
    If PeleaSegura(Atacante, Muerto) Then Exit Sub
    'Si se llevan más de 10 niveles no le cuento la muerte.
    If CInt(UserList(Atacante).Stats.ELV) - CInt(UserList(Muerto).Stats.ELV) > 10 Then Exit Sub
    Dim AttackerStatus As e_Facciones
    AttackerStatus = Status(Atacante)
    If Status(Muerto) = e_Facciones.Criminal Or Status(Muerto) = e_Facciones.Caos Or Status(Muerto) = e_Facciones.concilio Then
        'Si es un enfrentamiento entre Concilio–Caos penaliza siempre
        If AreLegionsOrCouncils(Atacante, Muerto) Then
            Call PenalizeFactionScoreLegionAndCouncil(Atacante, Muerto)
        End If
        If Not AlreadyKilledBy(Muerto, Atacante) Then
            Call RegisterRecentKiller(Muerto, Atacante)
            If UserList(Atacante).Faccion.CriminalesMatados < MAXUSERMATADOS Then
                UserList(Atacante).Faccion.CriminalesMatados = UserList(Atacante).Faccion.CriminalesMatados + 1
            End If
            If AttackerStatus = e_Facciones.Ciudadano Or AttackerStatus = e_Facciones.Armada Or AttackerStatus = e_Facciones.consejo Then
                Call HandleFactionScoreForKill(Atacante, Muerto)
            End If
        End If
    ElseIf Status(Muerto) = e_Facciones.Ciudadano Or Status(Muerto) = e_Facciones.Armada Or Status(Muerto) = e_Facciones.consejo Then
        If Not AlreadyKilledBy(Muerto, Atacante) Then
            Call RegisterRecentKiller(Muerto, Atacante)
            If UserList(Atacante).Faccion.ciudadanosMatados < MAXUSERMATADOS Then
                UserList(Atacante).Faccion.ciudadanosMatados = UserList(Atacante).Faccion.ciudadanosMatados + 1
            End If
            If AttackerStatus = e_Facciones.Criminal Or AttackerStatus = e_Facciones.Caos Or AttackerStatus = e_Facciones.concilio Then
                Call HandleFactionScoreForKill(Atacante, Muerto)
            End If
        End If
    End If
    Exit Sub
ContarMuerte_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ContarMuerte", Erl)
End Sub

Private Function ShouldApplyFactionBonus(ByVal attackerIndex As Integer, ByVal TargetIndex As Integer) As Boolean
    Dim Attacker As Byte
    Dim Target   As Byte
    Attacker = UserList(attackerIndex).Faccion.Status
    Target = UserList(TargetIndex).Faccion.Status
    Dim caosVsArmadaOrConsejo     As Boolean
    Dim concilioVsArmadaOrConsejo As Boolean
    Dim armadaVsCaosOrConcilio    As Boolean
    Dim consejoVsCaosOrConcilio   As Boolean
    caosVsArmadaOrConsejo = (Attacker = e_Facciones.Caos) And (Target = e_Facciones.Armada Or Target = e_Facciones.consejo)
    concilioVsArmadaOrConsejo = (Attacker = e_Facciones.concilio) And (Target = e_Facciones.Armada Or Target = e_Facciones.consejo)
    armadaVsCaosOrConcilio = (Attacker = e_Facciones.Armada) And (Target = e_Facciones.Caos Or Target = e_Facciones.concilio)
    consejoVsCaosOrConcilio = (Attacker = e_Facciones.consejo) And (Target = e_Facciones.Caos Or Target = e_Facciones.concilio)
    ShouldApplyFactionBonus = caosVsArmadaOrConsejo Or concilioVsArmadaOrConsejo Or armadaVsCaosOrConcilio Or consejoVsCaosOrConcilio
End Function

Sub HandleFactionScoreForKill(ByVal UserIndex As Integer, ByVal TargetIndex As Integer)
    Dim Score As Integer
    With UserList(UserIndex)
        Score = CalculateBaseFactionScore(UserIndex, TargetIndex)
        If ShouldApplyFactionBonus(UserIndex, TargetIndex) Then
            Score = Int(Score * 1.5)
        End If
        If Score > 20 Then
            Score = 20
        End If
        If GlobalFrameTime - .flags.LastHelpByTime < AssistHelpValidTime Then
            If IsValidUserRef(.flags.LastHelpUser) And .flags.LastHelpUser.ArrayIndex <> UserIndex Then
                Score = Score - 1
                Call HandleFactionScoreForAssist(.flags.LastHelpUser.ArrayIndex, TargetIndex)
            End If
        End If
        If GlobalFrameTime - UserList(TargetIndex).flags.LastAttackedByUserTime < AssistDamageValidTime Then
            If IsValidUserRef(UserList(TargetIndex).flags.LastAttacker) And UserList(TargetIndex).flags.LastAttacker.ArrayIndex <> UserIndex Then
                Score = Score - 1
                Call HandleFactionScoreForAssist(UserList(TargetIndex).flags.LastAttacker.ArrayIndex, TargetIndex)
            End If
        End If
        If AreLegionsOrCouncils(UserIndex, TargetIndex) Then
            Call PenalizeFactionScoreLegionAndCouncil(UserIndex, TargetIndex)
        Else
            'Mantener comportamiento original
            .Faccion.FactionScore = .Faccion.FactionScore + max(Score, 0)
        End If
    End With
End Sub

Sub HandleFactionScoreForAssist(ByVal UserIndex As Integer, ByVal TargetIndex As Integer)
    Dim Score As Integer
    With UserList(UserIndex)
        'Calcular el puntaje base de asistencia
        Score = 10 - max(CInt(.Stats.ELV) - CInt(UserList(TargetIndex).Stats.ELV), 0)
        Score = Score / 2
        If AreLegionsOrCouncils(UserIndex, TargetIndex) Then
            'Penalizar asistencias entre Legión y Concilio
            Dim newScore As Long
            newScore = .Faccion.FactionScore - Abs(Score)
            If newScore < 0 Then newScore = 0  ' Evitar que baje de 0
            .Faccion.FactionScore = newScore
        Else
            'Mantener comportamiento original
            .Faccion.FactionScore = .Faccion.FactionScore + max(Score, 0)
        End If
    End With
End Sub

Sub PenalizeFactionScoreLegionAndCouncil(ByVal Attacker As Integer, ByVal Target As Integer)
    On Error GoTo PenalizeFactionScoreLegionAndCouncil_Err
    With UserList(Attacker)
        Dim Score As Integer
        ' Calcular Score base según diferencia de niveles
        Score = CalculateBaseFactionScore(Attacker, Target)
        ' Aplicar bonus si corresponde
        If ShouldApplyFactionBonus(Attacker, Target) Then
            Score = Int(Score * 1.5)
        End If
        ' Limitar tope máximo
        If Score > 20 Then Score = 20
        ' Forzar penalización: siempre resta para Concilio–Caos
        Score = -Abs(Score)
        ' Aplicar y evitar bajar de 0
        Dim newScore As Long
        newScore = .Faccion.FactionScore + Score
        If newScore < 0 Then newScore = 0
        .Faccion.FactionScore = newScore
    End With
    Exit Sub
PenalizeFactionScoreLegionAndCouncil_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.PenalizeFactionScoreLegionAndCouncil", Erl)
End Sub

Private Function AreLegionsOrCouncils(ByVal Attacker As Integer, ByVal Target As Integer) As Boolean
    AreLegionsOrCouncils = ((UserList(Attacker).Faccion.Status = e_Facciones.concilio Or UserList(Attacker).Faccion.Status = e_Facciones.Caos) And (UserList( _
            Target).Faccion.Status = e_Facciones.concilio Or UserList(Target).Faccion.Status = e_Facciones.Caos))
End Function

Private Function CalculateBaseFactionScore(ByVal Attacker As Integer, ByVal Target As Integer) As Integer
    With UserList(Attacker)
        If CInt(.Stats.ELV) < CInt(UserList(Target).Stats.ELV) Then
            CalculateBaseFactionScore = 10 + CInt(UserList(Target).Stats.ELV) - max(CInt(.Stats.ELV), 0)
        Else
            CalculateBaseFactionScore = 10 - max(CInt(.Stats.ELV) - CInt(UserList(Target).Stats.ELV), 0)
        End If
    End With
End Function

Sub Tilelibre(ByRef pos As t_WorldPos, ByRef nPos As t_WorldPos, ByRef obj As t_Obj, ByRef Agua As Boolean, ByRef Tierra As Boolean, Optional ByVal InitialPos As Boolean = True)
    On Error GoTo Tilelibre_Err
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 23/01/2007
    '23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
    '**************************************************************
    Dim Notfound As Boolean
    Dim LoopC    As Integer
    Dim tX       As Integer
    Dim tY       As Integer
    Dim hayobj   As Boolean
    hayobj = False
    nPos.Map = pos.Map
    Do While Not LegalPos(pos.Map, nPos.x, nPos.y, Agua, Tierra) Or hayobj
        If LoopC > 15 Then
            Notfound = True
            Exit Do
        End If
        For tY = pos.y - LoopC To pos.y + LoopC
            For tX = pos.x - LoopC To pos.x + LoopC
                If LegalPos(nPos.Map, tX, tY, Agua, Tierra) Then
                    'We continue searching for a valid tile if
                    'there is already an item on the floor that differs from the item being dropped
                    'the item on the floor is the same but the elemental tags differ
                    'the amount of items exceeds the max quantity of items on the floor
                    hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex <> obj.ObjIndex)
                    If Not hayobj Then hayobj = MapData(nPos.Map, tX, tY).ObjInfo.ElementalTags > 0 And MapData(nPos.Map, tX, tY).ObjInfo.ElementalTags <> obj.ElementalTags
                    If Not hayobj Then hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.amount + obj.amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 And (InitialPos Or (tX <> pos.x And tY <> pos.y)) Then
                        nPos.x = tX
                        nPos.y = tY
                        tX = pos.x + LoopC
                        tY = pos.y + LoopC
                    End If
                End If
            Next tX
        Next tY
        LoopC = LoopC + 1
    Loop
    If Notfound = True Then
        nPos.x = 0
        nPos.y = 0
    End If
    Exit Sub
Tilelibre_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.Tilelibre", Erl)
End Sub

Sub WarpToLegalPos(ByVal UserIndex As Integer, _
                   ByVal Map As Integer, _
                   ByVal x As Byte, _
                   ByVal y As Byte, _
                   Optional ByVal FX As Boolean = False, _
                   Optional ByVal AguaValida As Boolean = False)
    On Error GoTo WarpToLegalPos_Err
    Dim LoopC As Integer
    Dim tX    As Integer
    Dim tY    As Integer
    Do While True
        If LoopC > 20 Then Exit Sub
        For tY = y - LoopC To y + LoopC
            For tX = x - LoopC To x + LoopC
                If LegalPos(Map, tX, tY, AguaValida, True, UserList(UserIndex).flags.Montado = 1, False, False) Then
                    If MapData(Map, tX, tY).trigger < 50 Then
                        Call WarpUserChar(UserIndex, Map, tX, tY, FX)
                        Exit Sub
                    End If
                End If
            Next tX
        Next tY
        LoopC = LoopC + 1
    Loop
    Call WarpUserChar(UserIndex, Map, x, y, FX)
    Exit Sub
WarpToLegalPos_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.WarpToLegalPos", Erl)
End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal FX As Boolean = False)
    On Error GoTo WarpUserChar_Err
    Dim OldMap As Integer
    Dim OldX   As Integer
    Dim OldY   As Integer
    With UserList(UserIndex)
        If Map <= 0 Then Exit Sub
        Call CancelarComercioUsuario(UserIndex, 1844)
        'Quitar el dialogo
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.charindex))
        Call WriteRemoveAllDialogs(UserIndex)
        OldMap = .pos.Map
        OldX = .pos.x
        OldY = .pos.y
        Call EraseUserChar(UserIndex, True, FX)
        If OldMap <> Map Then
            Call WriteChangeMap(UserIndex, Map)
            If MapInfo(OldMap).Seguro = 1 And MapInfo(Map).Seguro = 0 And .Stats.ELV < 42 Then
                ' Msg573=Estás saliendo de una zona segura, recuerda que aquí corres riesgo de ser atacado.
                Call WriteLocaleMsg(UserIndex, 573, e_FontTypeNames.FONTTYPE_WARNING)
            End If
            'Update new Map Users
            MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
            'Update old Map Users
            MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
            If MapInfo(OldMap).NumUsers < 0 Then
                MapInfo(OldMap).NumUsers = 0
            End If
            If .flags.Traveling = 1 Then
                .flags.Traveling = 0
                .Counters.goHome = 0
                ' Msg574=El viaje ha terminado.
                Call WriteLocaleMsg(UserIndex, 574, e_FontTypeNames.FONTTYPE_INFOBOLD)
            End If
        End If
        .pos.x = x
        .pos.y = y
        .pos.Map = Map
        If .Grupo.EnGrupo = True Then
            Call CompartirUbicacion(UserIndex)
        End If
        If FX Then
            Call MakeUserChar(True, Map, UserIndex, Map, x, y, 1)
        Else
            Call MakeUserChar(True, Map, UserIndex, Map, x, y, 0)
        End If
        Call WriteUserCharIndexInServer(UserIndex)
        If IsValidUserRef(.flags.GMMeSigue) Then
            Call WriteSendFollowingCharindex(.flags.GMMeSigue.ArrayIndex, .Char.charindex)
        End If
        'Seguis invisible al pasar de mapa
        If (.flags.invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then
            ' Si el mapa lo permite
            If MapInfo(Map).SinInviOcul Then
                .flags.invisible = 0
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                .Counters.Invisibilidad = 0
                .Counters.DisabledInvisibility = 0
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.charindex, False))
                ' Msg575=Una fuerza divina que vigila esta zona te ha vuelto visible.
                Call WriteLocaleMsg(UserIndex, 575, e_FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))
            End If
        End If
        'Reparacion temporal del bug de particulas. 08/07/09 LADDER
        If .flags.AdminInvisible = 0 Then
            If FX Then 'FX
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_WARP, x, y))
                UserList(UserIndex).Counters.timeFx = 3
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, e_GraphicEffects.ModernGmWarp, 0, .pos.x, .pos.y))
            End If
        Else
            Call SendData(ToIndex, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))
        End If
        If .NroMascotas > 0 Then Call WarpMascotas(UserIndex)
        If MapInfo(Map).zone = "DUNGEON" Or MapData(Map, x, y).trigger >= 9 Then
            If .flags.Montado > 0 Then
                Call DoMontar(UserIndex, ObjData(.invent.EquippedSaddleObjIndex), .invent.EquippedSaddleSlot)
            End If
        End If
        Call CancelarComercioUsuario(UserIndex, 1844)
    End With
    Exit Sub
WarpUserChar_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.WarpUserChar", Erl)
End Sub

Sub Cerrar_Usuario(ByVal UserIndex As Integer, Optional ByVal forceClose As Boolean = False)
    On Error GoTo Cerrar_Usuario_Err
    With UserList(UserIndex)
        If IsFeatureEnabled("debug_connections") Then
            Call AddLogToCircularBuffer("Cerrar_Usuario: " & UserIndex & ", force close: " & forceClose & ", usrLogged: " & .flags.UserLogged & ", Saliendo: " & .Counters.Saliendo)
        End If
        If .flags.UserLogged And Not .Counters.Saliendo Then
            .Counters.Saliendo = True
            .Counters.Salir = IntervaloCerrarConexion
            If .flags.Traveling = 1 Then
                ' Msg576=Se ha cancelado el viaje a casa
                Call WriteLocaleMsg(UserIndex, 576, e_FontTypeNames.FONTTYPE_INFO)
                .flags.Traveling = 0
                .Counters.goHome = 0
            End If
            If .flags.invisible + .flags.Oculto > 0 Then
                .flags.invisible = 0
                .flags.Oculto = 0
                .Counters.DisabledInvisibility = 0
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                ' Msg577=Has vuelto a ser visible
                Call WriteLocaleMsg(UserIndex, 577, e_FontTypeNames.FONTTYPE_INFO)
            End If
            'HarThaoS: Captura de bandera
            If .flags.jugando_captura = 1 Then
                If Not InstanciaCaptura Is Nothing Then
                    Call InstanciaCaptura.eliminarParticipante(InstanciaCaptura.GetPlayer(UserIndex))
                End If
            End If
            Call WriteLocaleMsg(UserIndex, 203, e_FontTypeNames.FONTTYPE_INFO, .Counters.Salir)
            If EsGM(UserIndex) Or MapInfo(.pos.Map).Seguro = 1 Or forceClose Then
                Call WriteDisconnect(UserIndex)
                Call CloseSocket(UserIndex)
            End If
        End If
    End With
    Exit Sub
Cerrar_Usuario_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.Cerrar_Usuario", Erl)
End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.
Public Sub CancelExit(ByVal UserIndex As Integer)
    On Error GoTo CancelExit_Err
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/02/08
    '
    '***************************************************
    If UserList(UserIndex).Counters.Saliendo And UserList(UserIndex).ConnectionDetails.ConnIDValida Then
        ' Is the user still connected?
        If UserList(UserIndex).ConnectionDetails.ConnIDValida Then
            UserList(UserIndex).Counters.Saliendo = False
            UserList(UserIndex).Counters.Salir = 0
            ' Msg578=/salir cancelado.
            Call WriteLocaleMsg(UserIndex, 578, e_FontTypeNames.FONTTYPE_WARNING)
        Else
            'Simply reset
            If UserList(UserIndex).flags.Privilegios = e_PlayerType.User And MapInfo(UserList(UserIndex).pos.Map).Seguro = 0 Then
                UserList(UserIndex).Counters.Salir = IntervaloCerrarConexion
            Else
                ' Msg579=Gracias por jugar Argentum Online.
                Call WriteLocaleMsg(UserIndex, 579, e_FontTypeNames.FONTTYPE_INFO)
                Call WriteDisconnect(UserIndex)
                Call CloseSocket(UserIndex)
            End If
            'UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And e_PlayerType.User) And MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0, IntervaloCerrarConexion, 0)
        End If
    End If
    Exit Sub
CancelExit_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.CancelExit", Erl)
End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)
    On Error GoTo VolverCriminal_Err
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 21/06/2006
    'Nacho: Actualiza el tag al cliente
    '**************************************************************
    With UserList(UserIndex)
        If MapData(.pos.Map, .pos.x, .pos.y).trigger = 6 Then Exit Sub
        If .flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero) Then
            If .Faccion.Status = e_Facciones.Armada Then
                '  NUNCA debería pasar, pero dejo un log por si las...
                Call TraceError(111, "Un personaje de la Armada Real atacó un ciudadano.", "UsUaRiOs.VolverCriminal")
                'Call ExpulsarFaccionReal(UserIndex)
            End If
        End If
        If .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then Exit Sub
        If .Faccion.Status = e_Facciones.Ciudadano Then
            .Faccion.FactionScore = 0
        End If
        .Faccion.Status = 0
        If MapInfo(.pos.Map).NoPKs And Not EsGM(UserIndex) And MapInfo(.pos.Map).Salida.Map <> 0 Then
            ' Msg580=En este mapa no se admiten criminales.
            Call WriteLocaleMsg(UserIndex, 580, e_FontTypeNames.FONTTYPE_INFO)
            Call WarpUserChar(UserIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
        Else
            Call RefreshCharStatus(UserIndex)
        End If
    End With
    Exit Sub
VolverCriminal_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.VolverCriminal", Erl)
End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 21/06/2006
    'Nacho: Actualiza el tag al cliente.
    '**************************************************************
    On Error GoTo VolverCiudadano_Err
    With UserList(UserIndex)
        If MapData(.pos.Map, .pos.x, .pos.y).trigger = 6 Then Exit Sub
        If .Faccion.Status = e_Facciones.Criminal Or .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then
            .Faccion.FactionScore = 0
        End If
        .Faccion.Status = e_Facciones.Ciudadano
        If MapInfo(.pos.Map).NoCiudadanos And Not EsGM(UserIndex) And MapInfo(.pos.Map).Salida.Map <> 0 Then
            ' Msg581=En este mapa no se admiten ciudadanos.
            Call WriteLocaleMsg(UserIndex, 581, e_FontTypeNames.FONTTYPE_INFO)
            Call WarpUserChar(UserIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
        Else
            Call RefreshCharStatus(UserIndex)
        End If
        Call WriteSafeModeOn(UserIndex)
        .flags.Seguro = True
    End With
    Exit Sub
VolverCiudadano_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.VolverCiudadano", Erl)
End Sub

Public Function getMaxInventorySlots(ByVal UserIndex As Integer) As Byte
    On Error GoTo getMaxInventorySlots_Err
    getMaxInventorySlots = MAX_USERINVENTORY_SLOTS
    With UserList(UserIndex)
        getMaxInventorySlots = get_num_inv_slots_from_tier(.Stats.tipoUsuario)
    End With
    Exit Function
getMaxInventorySlots_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.getMaxInventorySlots", Erl)
End Function

Private Sub WarpMascotas(ByVal UserIndex As Integer)
    On Error GoTo WarpMascotas_Err
    Dim i                As Integer
    Dim petType          As Integer
    Dim PermiteMascotas  As Boolean
    Dim Index            As Integer
    Dim iMinHP           As Integer
    Dim PetTiempoDeVida  As Integer
    Dim MascotaQuitada   As Boolean
    Dim ElementalQuitado As Boolean
    Dim SpawnInvalido    As Boolean
    PermiteMascotas = MapInfo(UserList(UserIndex).pos.Map).NoMascotas = False
    For i = 1 To MAXMASCOTAS
        Index = UserList(UserIndex).MascotasIndex(i).ArrayIndex
        If IsValidNpcRef(UserList(UserIndex).MascotasIndex(i)) Then
            iMinHP = NpcList(Index).Stats.MinHp
            PetTiempoDeVida = NpcList(Index).Contadores.TiempoExistencia
            Call SetUserRef(NpcList(Index).MaestroUser, 0)
            Call QuitarNPC(Index, eRemoveWarpPets)
            If PetTiempoDeVida > 0 Then
                Call QuitarMascota(UserIndex, Index)
                ElementalQuitado = True
            ElseIf Not PermiteMascotas Then
                Call ClearNpcRef(UserList(UserIndex).MascotasIndex(i))
                MascotaQuitada = True
            End If
        Else
            Call ClearNpcRef(UserList(UserIndex).MascotasIndex(i))
            iMinHP = 0
            PetTiempoDeVida = 0
        End If
        petType = UserList(UserIndex).MascotasType(i)
        If petType > 0 And PermiteMascotas And (UserList(UserIndex).flags.MascotasGuardadas = 0 Or UserList(UserIndex).MascotasIndex(i).ArrayIndex > 0) And PetTiempoDeVida = 0 Then
            Dim SpawnPos As t_WorldPos
            SpawnPos.Map = UserList(UserIndex).pos.Map
            SpawnPos.x = UserList(UserIndex).pos.x + RandomNumber(-3, 3)
            SpawnPos.y = UserList(UserIndex).pos.y + RandomNumber(-3, 3)
            Index = SpawnNpc(petType, SpawnPos, False, False, False, UserIndex)
            'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
            ' Exception: Pets don't spawn in water if they can't swim
            If Index > 0 Then
                Call SetNpcRef(UserList(UserIndex).MascotasIndex(i), Index)
                ' Nos aseguramos de que conserve el hp, si estaba danado
                If iMinHP Then NpcList(Index).Stats.MinHp = iMinHP
                Call SetUserRef(NpcList(Index).MaestroUser, UserIndex)
                Call FollowAmo(Index)
            Else
                SpawnInvalido = True
            End If
        End If
    Next i
    If MascotaQuitada Then
        If Not PermiteMascotas Then
            ' Msg582=Una fuerza superior impide que tus mascotas entren en este mapa. Estas te esperarán afuera.
            Call WriteLocaleMsg(UserIndex, 582, e_FontTypeNames.FONTTYPE_INFO)
        End If
    ElseIf SpawnInvalido Then
        ' Msg583=Tus mascotas no pueden transitar este mapa.
        Call WriteLocaleMsg(UserIndex, 583, e_FontTypeNames.FONTTYPE_INFO)
    ElseIf ElementalQuitado Then
        ' Msg584=Pierdes el control de tus mascotas invocadas.
        Call WriteLocaleMsg(UserIndex, 584, e_FontTypeNames.FONTTYPE_INFO)
    End If
    Exit Sub
WarpMascotas_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.WarpMascotas", Erl)
End Sub

Function TieneArmaduraCazador(ByVal UserIndex As Integer) As Boolean
    On Error GoTo TieneArmaduraCazador_Err
    If UserList(UserIndex).invent.EquippedArmorObjIndex > 0 Then
        If ObjData(UserList(UserIndex).invent.EquippedArmorObjIndex).Subtipo = 3 Then ' Aguante hardcodear números :D
            TieneArmaduraCazador = True
        End If
    End If
    Exit Function
TieneArmaduraCazador_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.TieneArmaduraCazador", Erl)
End Function

Public Sub SetModoConsulta(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 05/06/10
    '
    '***************************************************
    Dim sndNick As String
    With UserList(UserIndex)
        sndNick = .name
        If .flags.EnConsulta Then
            sndNick = sndNick & " [CONSULTA]"
        Else
            If .GuildIndex > 0 Then
                sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"
            End If
        End If
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, .Faccion.Status, sndNick))
    End With
End Sub

' Autor: WyroX - 20/01/2021
' Intenta moverlo hacia un "costado" según el heading indicado.
' Si no hay un lugar válido a los lados, lo mueve a la posición válida más cercana.
Sub MoveUserToSide(ByVal UserIndex As Integer, ByVal Heading As e_Heading)
    On Error GoTo Handler
    With UserList(UserIndex)
        ' Elegimos un lado al azar
        Dim r As Integer
        r = RandomNumber(0, 1) * 2 - 1 ' -1 o 1
        ' Roto el heading original hacia ese lado
        Heading = Rotate_Heading(Heading, r)
        ' Intento moverlo para ese lado
        If MoveUserChar(UserIndex, Heading) Then
            ' Le aviso al usuario que fue movido
            Call WriteForceCharMove(UserIndex, Heading)
            Exit Sub
        End If
        ' Si falló, intento moverlo para el lado opuesto
        Heading = InvertHeading(Heading)
        If MoveUserChar(UserIndex, Heading) Then
            ' Le aviso al usuario que fue movido
            Call WriteForceCharMove(UserIndex, Heading)
            Exit Sub
        End If
        ' Si ambos fallan, entonces lo dejo en la posición válida más cercana
        Dim NuevaPos As t_WorldPos
        Call ClosestLegalPos(.pos, NuevaPos, .flags.Navegando, .flags.Navegando = 0)
        Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.x, NuevaPos.y)
    End With
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.MoveUserToSide", Erl)
End Sub

' Autor: WyroX - 02/03/2021
' Quita parálisis, veneno, invisibilidad, estupidez, mimetismo, deja de descansar, de meditar y de ocultarse; y quita otros estados obsoletos (por si acaso)
Public Sub LimpiarEstadosAlterados(ByVal UserIndex As Integer)
    On Error GoTo Handler
    With UserList(UserIndex)
        .flags.DivineBlood = 0
        '<<<< Envenenamiento >>>>
        .flags.Envenenado = 0
        '<<<< Paralisis >>>>
        If .flags.Paralizado = 1 Then
            .flags.Paralizado = 0
            Call WriteParalizeOK(UserIndex)
        End If
        '<<<< Inmovilizado >>>>
        If .flags.Inmovilizado = 1 Then
            .flags.Inmovilizado = 0
            Call WriteInmovilizaOK(UserIndex)
        End If
        '<<< Estupidez >>>
        If .flags.Estupidez = 1 Then
            .flags.Estupidez = 0
            Call WriteDumbNoMore(UserIndex)
        End If
        '<<<< Descansando >>>>
        If .flags.Descansar Then
            .flags.Descansar = False
            Call WriteRestOK(UserIndex)
        End If
        '<<<< Meditando >>>>
        If .flags.Meditando Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0, .pos.x, .pos.y))
        End If
        '<<<< Stun >>>>
        .Counters.StunEndTime = 0
        '<<<< Invisible >>>>
        If (.flags.invisible = 1 Or .flags.Oculto = 1) And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .flags.invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
            .Counters.DisabledInvisibility = 0
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        End If
        '<<<< Mimetismo >>>>
        If .flags.Mimetizado > 0 Then
            If .flags.Navegando Then
                If .flags.Muerto = 0 Then
                    .Char.body = ObjData(UserList(UserIndex).invent.EquippedShipObjIndex).Ropaje
                Else
                    .Char.body = iFragataFantasmal
                End If
                Call ClearClothes(.Char)
            Else
                .Char.body = .CharMimetizado.body
                .Char.head = .CharMimetizado.head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Char.CartAnim = .CharMimetizado.CartAnim
            End If
            .Counters.Mimetismo = 0
            .flags.Mimetizado = e_EstadoMimetismo.Desactivado
        End If
        '<<<< Estados obsoletos >>>>
        .flags.Incinerado = 0
    End With
    Exit Sub
Handler:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.LimpiarEstadosAlterados", Erl)
End Sub

Public Sub DevolverPosAnterior(ByVal UserIndex As Integer)
    With UserList(UserIndex).flags
        Call WarpToLegalPos(UserIndex, .LastPos.Map, .LastPos.x, .LastPos.y, True)
    End With
End Sub

Public Function ActualizarVelocidadDeUsuario(ByVal UserIndex As Integer) As Single
    On Error GoTo ActualizarVelocidadDeUsuario_Err
    Dim velocidad As Single, modificadorItem As Single, modificadorHechizo As Single, JineteLevelSpeed As Single
    velocidad = VelocidadNormal
    modificadorItem = 1
    modificadorHechizo = 1
    JineteLevelSpeed = 1
    With UserList(UserIndex)
        If .flags.Muerto = 1 Then
            velocidad = VelocidadMuerto
            GoTo UpdateSpeed ' Los muertos no tienen modificadores de velocidad
        End If
        ' El traje para nadar es considerado barco, de subtipo = 0
        If (.flags.Navegando + .flags.Nadando > 0) And (.invent.EquippedShipObjIndex > 0) Then
            modificadorItem = ObjData(.invent.EquippedShipObjIndex).velocidad
        End If
        If (.flags.Montado = 1) And (.invent.EquippedSaddleObjIndex > 0) Then
            modificadorItem = ObjData(.invent.EquippedSaddleObjIndex).velocidad
            Select Case .Stats.JineteLevel
                Case 1
                    JineteLevelSpeed = SvrConfig.GetValue("JineteLevel1Speed")
                Case 2
                    JineteLevelSpeed = SvrConfig.GetValue("JineteLevel2Speed")
                Case 3
                    JineteLevelSpeed = SvrConfig.GetValue("JineteLevel3Speed")
                Case 4
                    JineteLevelSpeed = SvrConfig.GetValue("JineteLevel4Speed")
                Case 5
                    JineteLevelSpeed = SvrConfig.GetValue("JineteLevel5Speed")
                Case 6
                    JineteLevelSpeed = SvrConfig.GetValue("JineteLevel6Speed")
                Case 7
                    JineteLevelSpeed = SvrConfig.GetValue("JineteLevel7Speed")
                Case 8
                    JineteLevelSpeed = SvrConfig.GetValue("JineteLevel8Speed")
                Case 9
                    JineteLevelSpeed = SvrConfig.GetValue("JineteLevel9Speed")
                Case Else
                    JineteLevelSpeed = 1
            End Select
        End If
        ' Algun hechizo le afecto la velocidad
        If .flags.VelocidadHechizada > 0 Then
            modificadorHechizo = .flags.VelocidadHechizada
        End If
        If .invent.EquippedArmorObjIndex > 0 Then
            If ObjData(.invent.EquippedArmorObjIndex).velocidad <> 1 Then
                modificadorItem = modificadorItem * ObjData(.invent.EquippedArmorObjIndex).velocidad
            End If
        End If
        velocidad = VelocidadNormal * modificadorItem * JineteLevelSpeed * modificadorHechizo * max(0, (1 + .Modifiers.MovementSpeed))
UpdateSpeed:
        .Char.speeding = velocidad
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.charindex, .Char.speeding))
        Call WriteVelocidadToggle(UserIndex)
    End With
    Exit Function
ActualizarVelocidadDeUsuario_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.CalcularVelocidad_Err", Erl)
End Function

Public Sub ClearClothes(ByRef Char As t_Char)
    Char.ShieldAnim = NingunEscudo
    Char.WeaponAnim = NingunArma
    Char.CascoAnim = NingunCasco
    Char.CartAnim = NoCart
End Sub

Public Function IsStun(ByRef flags As t_UserFlags, ByRef Counters As t_UserCounters) As Boolean
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()
    ' Player is stunned if current tick has NOT yet passed the stun end deadline
    IsStun = Not DeadlinePassed(nowRaw, Counters.StunEndTime)
End Function

Public Function CanMove(ByRef flags As t_UserFlags, ByRef Counters As t_UserCounters) As Boolean
    CanMove = flags.Paralizado = 0 And flags.Inmovilizado = 0 And Not IsStun(flags, Counters) And Not flags.TranslationActive
End Function

Public Function StunPlayer(ByVal UserIndex As Integer, ByRef Counters As t_UserCounters) As Boolean
    On Error GoTo eh
    StunPlayer = False
    ' (Optional) your CanMove signature might be (counters, flags) — adjust order if needed
    If Not CanMove(UserList(UserIndex).flags, Counters) Then Exit Function
    If IsSet(UserList(UserIndex).flags.StatusMask, eCCInmunity) Then Exit Function
    Dim nowRaw As Long
    nowRaw = GetTickCountRaw()   ' <-- use raw tick
    ' Respect anti-chain-stun window: allow new stun only after immune window passes
    Dim immuneUntil As Long
    immuneUntil = AddMod32(Counters.StunEndTime, PlayerInmuneTime) ' old end + immunity
    If TickAfter(nowRaw, immuneUntil) Then
        ' Apply (or re-apply) stun: set absolute deadline using modulo-2^32 add
        Counters.StunEndTime = AddMod32(nowRaw, PlayerStunTime)
        StunPlayer = True
    End If
    Exit Function
eh:
End Function

Public Function CanUseItem(ByRef flags As t_UserFlags, ByRef Counters As t_UserCounters) As Boolean
    CanUseItem = True
End Function

Public Sub UpdateCd(ByVal UserIndex As Integer, ByVal cdType As e_CdTypes)
    UserList(UserIndex).CdTimes(cdType) = GetTickCountRaw()
    Call WriteUpdateCdType(UserIndex, cdType)
End Sub

Public Function IsVisible(ByRef User As t_User) As Boolean
    IsVisible = (Not (User.flags.invisible > 0 Or User.flags.Oculto > 0))
End Function

Public Function CanHelpUser(ByVal UserIndex As Integer, ByVal targetUserIndex As Integer) As e_InteractionResult
    CanHelpUser = eInteractionOk
    If UserList(UserIndex).flags.CurrentTeam > 0 And UserList(UserIndex).flags.CurrentTeam <> UserList(targetUserIndex).flags.CurrentTeam Then
        CanHelpUser = eDifferentTeam
        Exit Function
    End If
    If PeleaSegura(UserIndex, targetUserIndex) Then
        Exit Function
    End If
    Dim TargetStatus As e_Facciones
    TargetStatus = Status(targetUserIndex)
    Select Case Status(UserIndex)
        Case e_Facciones.Ciudadano, e_Facciones.Armada, e_Facciones.consejo
            If TargetStatus = e_Facciones.Caos Or TargetStatus = e_Facciones.concilio Then
                CanHelpUser = eOposingFaction
                Exit Function
            ElseIf TargetStatus = e_Facciones.Criminal Then
                If UserList(UserIndex).flags.Seguro Then
                    CanHelpUser = eCantHelpCriminal
                Else
                    If UserList(UserIndex).GuildIndex > 0 Then
                        'Si el clan es de alineación ciudadana.
                        If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                            'No lo dejo resucitarlo
                            CanHelpUser = eCantHelpCriminalClanRules
                            Exit Function
                            'Si es de alineación neutral, lo dejo resucitar y lo vuelvo criminal
                        ElseIf GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_NEUTRAL Then
                            Call VolverCriminal(UserIndex)
                            Call RefreshCharStatus(UserIndex)
                            Exit Function
                        End If
                    Else
                        Call VolverCriminal(UserIndex)
                        Call RefreshCharStatus(UserIndex)
                        Exit Function
                    End If
                End If
            End If
        Case e_Facciones.Caos, e_Facciones.concilio
            If Status(targetUserIndex) = e_Facciones.Armada Or Status(targetUserIndex) = e_Facciones.consejo Or Status(targetUserIndex) = e_Facciones.Ciudadano Then
                CanHelpUser = eOposingFaction
            End If
        Case Else
            Exit Function
    End Select
End Function

Public Function CanAttackUser(ByVal attackerIndex As Integer, _
                              ByVal AttackerVersionID As Integer, _
                              ByVal TargetIndex As Integer, _
                              ByVal TargetVersionID As Integer) As e_AttackInteractionResult
    If UserList(TargetIndex).flags.Muerto = 1 Then
        CanAttackUser = e_AttackInteractionResult.eDeathTarget
        Exit Function
    End If
    If attackerIndex = TargetIndex And AttackerVersionID = TargetVersionID Then
        CanAttackUser = e_AttackInteractionResult.eCantAttackYourself
        Exit Function
    End If
    If UserList(attackerIndex).flags.EnReto Then
        If Retos.Salas(UserList(attackerIndex).flags.SalaReto).TiempoItems > 0 Then
            CanAttackUser = e_AttackInteractionResult.eFightActive
            Exit Function
        End If
    End If
    If UserList(attackerIndex).Grupo.Id > 0 And UserList(TargetIndex).Grupo.Id > 0 And UserList(attackerIndex).Grupo.Id = UserList(TargetIndex).Grupo.Id Then
        CanAttackUser = eSameGroup
        Exit Function
    End If
    If UserList(attackerIndex).flags.EnConsulta Or UserList(TargetIndex).flags.EnConsulta Then
        CanAttackUser = eTalkWithMaster
        Exit Function
    End If
    If UserList(attackerIndex).flags.Maldicion = 1 Then
        CanAttackUser = eAttackerIsCursed
        Exit Function
    End If
    If UserList(attackerIndex).flags.Montado = 1 Then
        CanAttackUser = eMounted
        Exit Function
    End If
    If Not MapInfo(UserList(TargetIndex).pos.Map).FriendlyFire And UserList(TargetIndex).flags.CurrentTeam > 0 And UserList(TargetIndex).flags.CurrentTeam = UserList( _
            attackerIndex).flags.CurrentTeam Then
        CanAttackUser = eSameTeam
        Exit Function
    End If
    ' Nueva verificación específica para Captura la Bandera
    If UserList(attackerIndex).flags.jugando_captura = 1 And UserList(TargetIndex).flags.jugando_captura = 1 Then
        If UserList(attackerIndex).flags.CurrentTeam = UserList(TargetIndex).flags.CurrentTeam Then
            'Msg1102= ¡No puedes atacar a miembros de tu propio equipo!
            Call WriteLocaleMsg(attackerIndex, "1102", e_FontTypeNames.FONTTYPE_INFO)
            CanAttackUser = eSameTeam
            Exit Function
        End If
    End If
    Dim t As e_Trigger6
    'Estamos en una Arena? o un trigger zona segura?
    t = TriggerZonaPelea(attackerIndex, TargetIndex)
    If t = e_Trigger6.TRIGGER6_PERMITE Then
        CanAttackUser = eCanAttack
        Exit Function
    ElseIf PeleaSegura(attackerIndex, TargetIndex) Then
        CanAttackUser = eCanAttack
        Exit Function
    End If
    'Solo administradores pueden atacar a usuarios (PARA TESTING)
    If (UserList(attackerIndex).flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Admin)) = 0 Then
        CanAttackUser = eNotEnougthPrivileges
        Exit Function
    End If
    'Estas queriendo atacar a un GM?
    Dim rank As Integer
    rank = e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero
    If (UserList(TargetIndex).flags.Privilegios And rank) > (UserList(attackerIndex).flags.Privilegios And rank) Then
        CanAttackUser = eNotEnougthPrivileges
        Exit Function
    End If
    ' Seguro Clan
    If UserList(attackerIndex).GuildIndex > 0 Then
        If UserList(attackerIndex).flags.SeguroClan And NivelDeClan(UserList(attackerIndex).GuildIndex) >= 3 Then
            If UserList(attackerIndex).GuildIndex = UserList(TargetIndex).GuildIndex Then
                CanAttackUser = eSameClan
                Exit Function
            End If
        End If
    End If
    ' Es armada?
    If esArmada(attackerIndex) Then
        ' Si ataca otro armada
        If esArmada(TargetIndex) Then
            CanAttackUser = eSameFaction
            Exit Function
            ' Si ataca un ciudadano
        ElseIf esCiudadano(TargetIndex) Then
            CanAttackUser = eSameFaction
            Exit Function
        End If
        ' No es armada
    Else
        'Tenes puesto el seguro?
        If (esCiudadano(attackerIndex)) Then
            If (UserList(attackerIndex).flags.Seguro) Then
                If esCiudadano(TargetIndex) Then
                    CanAttackUser = eRemoveSafe
                    Exit Function
                ElseIf esArmada(TargetIndex) Then
                    CanAttackUser = eRemoveSafe
                    Exit Function
                End If
            End If
        ElseIf esCaos(attackerIndex) And esCaos(TargetIndex) Then
            If (UserList(attackerIndex).flags.LegionarySecure) Then
                CanAttackUser = eSameFaction
                Exit Function
            End If
        End If
    End If
    'Estas en un Mapa Seguro?
    If MapInfo(UserList(TargetIndex).pos.Map).Seguro = 1 Then
        If esArmada(attackerIndex) Then
            If UserList(attackerIndex).Faccion.RecompensasReal >= 3 Then
                If UserList(TargetIndex).pos.Map = 58 Or UserList(TargetIndex).pos.Map = 59 Or UserList(TargetIndex).pos.Map = 60 Then
                    CanAttackUser = eCanAttack
                    Exit Function
                End If
            End If
        End If
        If esCaos(attackerIndex) Then
            If UserList(attackerIndex).Faccion.RecompensasCaos >= 3 Then
                If UserList(TargetIndex).pos.Map = 195 Or UserList(TargetIndex).pos.Map = 196 Then
                    CanAttackUser = eCanAttack
                    Exit Function
                End If
            End If
        End If
        CanAttackUser = eSafeArea
        Exit Function
    End If
    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
    If MapData(UserList(TargetIndex).pos.Map, UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y).trigger = e_Trigger.ZonaSegura Or MapData(UserList( _
            attackerIndex).pos.Map, UserList(attackerIndex).pos.x, UserList(attackerIndex).pos.y).trigger = e_Trigger.ZonaSegura Then
        CanAttackUser = eSafeArea
        Exit Function
    End If
    CanAttackUser = eCanAttack
End Function

Public Function ModifyHealth(ByVal UserIndex As Integer, ByVal amount As Long, Optional ByVal MinValue = 0) As Boolean
    With UserList(UserIndex)
        ModifyHealth = False
        .Stats.MinHp = .Stats.MinHp + amount
        If .Stats.MinHp > .Stats.MaxHp Then
            .Stats.MinHp = .Stats.MaxHp
        End If
        If .Stats.MinHp <= MinValue Then
            .Stats.MinHp = MinValue
            ModifyHealth = True
        End If
        Call WriteUpdateHP(UserIndex)
    End With
End Function

Public Function ModifyStamina(ByVal UserIndex As Integer, ByVal amount As Integer, ByVal CancelIfNotEnought As Boolean, Optional ByVal MinValue = 0) As Boolean
    ModifyStamina = False
    With UserList(UserIndex)
        If CancelIfNotEnought And amount < 0 And .Stats.MinSta < Abs(amount) Then
            ModifyStamina = True
            Exit Function
        End If
        .Stats.MinSta = .Stats.MinSta + amount
        If .Stats.MinSta > .Stats.MaxSta Then
            .Stats.MinSta = .Stats.MaxSta
        End If
        If .Stats.MinSta < MinValue Then
            .Stats.MinSta = MinValue
            ModifyStamina = True
        End If
        Call WriteUpdateSta(UserIndex)
    End With
End Function

Public Function ModifyMana(ByVal UserIndex As Integer, ByVal amount As Integer, ByVal CancelIfNotEnought As Boolean, Optional ByVal MinValue = 0) As Boolean
    ModifyMana = False
    With UserList(UserIndex)
        If CancelIfNotEnought And amount < 0 And .Stats.MinMAN < Abs(amount) Then
            ModifyMana = True
            Exit Function
        End If
        .Stats.MinMAN = .Stats.MinMAN + amount
        If .Stats.MinMAN > .Stats.MaxMAN Then
            .Stats.MinMAN = .Stats.MaxMAN
        End If
        If .Stats.MinMAN < MinValue Then
            .Stats.MinMAN = MinValue
            ModifyMana = True
        End If
        Call WriteUpdateMana(UserIndex)
    End With
End Function

Public Sub ResurrectUser(ByVal targetUserIndex As Integer, Optional ByVal CasterUserIndex As Integer)
    ' Msg585=¡Has sido resucitado!
    Call WriteLocaleMsg(targetUserIndex, "585", e_FontTypeNames.FONTTYPE_INFO)
    Call SendData(SendTarget.ToPCArea, targetUserIndex, PrepareMessageParticleFX(UserList(targetUserIndex).Char.charindex, e_ParticleEffects.Resucitar, 250, True))
    Call SendData(SendTarget.ToPCArea, targetUserIndex, PrepareMessagePlayWave(117, UserList(targetUserIndex).pos.x, UserList(targetUserIndex).pos.y))
    Call RevivirUsuario(targetUserIndex, True, CasterUserIndex)
    Call WriteUpdateHungerAndThirst(targetUserIndex)
End Sub

Public Function DoDamageOrHeal(ByVal UserIndex As Integer, _
                               ByVal SourceIndex As Integer, _
                               ByVal SourceType As e_ReferenceType, _
                               ByVal amount As Long, _
                               ByVal DamageSourceType As e_DamageSourceType, _
                               ByVal DamageSourceIndex As Integer, _
                               Optional DoDamageText As Integer = 389, _
                               Optional GotDamageText As Integer = 34, _
                               Optional ByVal DamageColor As Long = vbRed) As e_DamageResult
    On Error GoTo DoDamageOrHeal_Err
    Dim DamageStr As String
    Dim Color     As Long
    DamageStr = PonerPuntos(amount)
    If amount > 0 Then
        Color = vbGreen
    Else
        Color = DamageColor
    End If
    If amount < 0 Then
        DamageStr = PonerPuntos(Math.Abs(amount))
        If SourceType = eUser Then
            If UserList(SourceIndex).ChatCombate = 1 And DoDamageText > 0 Then
                Call WriteLocaleMsg(SourceIndex, DoDamageText, e_FontTypeNames.FONTTYPE_FIGHT, UserList(UserIndex).name & "¬" & DamageStr)
            End If
            If UserList(UserIndex).ChatCombate = 1 And GotDamageText > 0 Then
                Call WriteLocaleMsg(UserIndex, GotDamageText, e_FontTypeNames.FONTTYPE_FIGHT, UserList(SourceIndex).name & "¬" & DamageStr)
            End If
        End If
        amount = EffectsOverTime.TargetApplyDamageReduction(UserList(UserIndex).EffectOverTime, amount, SourceIndex, SourceType, DamageSourceType)
        Call EffectsOverTime.TargetWasDamaged(UserList(UserIndex).EffectOverTime, SourceIndex, SourceType, DamageSourceType)
    End If
    With UserList(UserIndex)
        If IsVisible(UserList(UserIndex)) Then
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageTextOverChar(DamageStr, .Char.charindex, Color))
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageTextOverChar(DamageStr, .Char.charindex, Color))
        End If
        If ModifyHealth(UserIndex, amount) Then
            Call TargetWasDamaged(UserList(UserIndex).EffectOverTime, SourceIndex, SourceType, DamageSourceType)
            Call CustomScenarios.UserDie(UserIndex)
            If SourceType = eUser Then
                Call ContarMuerte(UserIndex, SourceIndex)
                Call PlayerKillPlayer(.pos.Map, SourceIndex, UserIndex, DamageSourceType, DamageSourceIndex)
                Call ActStats(UserIndex, SourceIndex)
            Else
                Call NPcKillPlayer(.pos.Map, SourceIndex, UserIndex, DamageSourceType, DamageSourceIndex)
                Call WriteNPCKillUser(UserIndex)
                If IsValidUserRef(NpcList(SourceIndex).MaestroUser) Then
                    Call AllFollowAmo(NpcList(SourceIndex).MaestroUser.ArrayIndex)
                    Call PlayerKillPlayer(.pos.Map, NpcList(SourceIndex).MaestroUser.ArrayIndex, UserIndex, e_DamageSourceType.e_pet, 0)
                Else
                    'Al matarlo no lo sigue mas
                    Call SetMovement(SourceIndex, NpcList(SourceIndex).flags.OldMovement)
                    NpcList(SourceIndex).Hostile = NpcList(SourceIndex).flags.OldHostil
                    NpcList(SourceIndex).flags.AttackedBy = vbNullString
                    Call SetUserRef(NpcList(SourceIndex).TargetUser, 0)
                End If
                Call UserMod.UserDie(UserIndex)
            End If
            DoDamageOrHeal = eDead
            Exit Function
        End If
    End With
    DoDamageOrHeal = eStillAlive
    Exit Function
DoDamageOrHeal_Err:
    Call TraceError(Err.Number, Err.Description, "UserMod.DoDamageOrHeal", Erl)
End Function

Public Function GetPhysicalDamageModifier(ByRef User As t_User) As Single
    GetPhysicalDamageModifier = max(1 + User.Modifiers.PhysicalDamageBonus, 0)
End Function

Public Function GetMagicDamageModifier(ByRef User As t_User) As Single
    GetMagicDamageModifier = max(1 + User.Modifiers.MagicDamageBonus, 0)
End Function

Public Function GetMagicDamageReduction(ByRef User As t_User) As Single
    GetMagicDamageReduction = max(1 - User.Modifiers.MagicDamageReduction, 0)
End Function

Public Function GetPhysicDamageReduction(ByRef User As t_User) As Single
    GetPhysicDamageReduction = max(1 - User.Modifiers.PhysicalDamageReduction, 0)
End Function

Public Sub RemoveInvisibility(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .flags.invisible + .flags.Oculto > 0 And .flags.NoDetectable = 0 Then
            .flags.invisible = 0
            .flags.Oculto = 0
            .Counters.Invisibilidad = 0
            .Counters.Ocultando = 0
            .Counters.DisabledInvisibility = 0
            ' Msg591=Tu invisibilidad ya no tiene efecto.
            Call WriteLocaleMsg(UserIndex, 591, e_FontTypeNames.FONTTYPE_INFOIAO)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        End If
    End With
End Sub

Public Function Inmovilize(ByVal SourceIndex As Integer, ByVal TargetIndex As Integer, ByVal Time As Integer, ByVal FX As Integer) As Boolean
    Call UsuarioAtacadoPorUsuario(SourceIndex, TargetIndex)
    If IsSet(UserList(TargetIndex).flags.StatusMask, eCCInmunity) Then
        Call WriteLocaleMsg(SourceIndex, MsgCCInunity, e_FontTypeNames.FONTTYPE_FIGHT)
        Exit Function
    End If
    If CanMove(UserList(TargetIndex).flags, UserList(TargetIndex).Counters) Then
        UserList(TargetIndex).Counters.Inmovilizado = Time
        UserList(TargetIndex).flags.Inmovilizado = 1
        Call SendData(SendTarget.ToPCAliveArea, TargetIndex, PrepareMessageCreateFX(UserList(TargetIndex).Char.charindex, FX, 0, UserList(TargetIndex).pos.x, UserList( _
                TargetIndex).pos.y))
        Call WriteInmovilizaOK(TargetIndex)
        Call WritePosUpdate(TargetIndex)
        Inmovilize = True
    End If
End Function

Public Function GetArmorPenetration(ByVal UserIndex As Integer, ByVal TargetArmor As Integer) As Integer
    Dim ArmorPenetration As Integer
    If Not IsFeatureEnabled("armor_penetration_feature") Then Exit Function
    With UserList(UserIndex)
        If .invent.EquippedWeaponObjIndex > 0 Then
            ArmorPenetration = ObjData(.invent.EquippedWeaponObjIndex).IgnoreArmorAmmount
            If ObjData(.invent.EquippedWeaponObjIndex).IgnoreArmorPercent > 0 Then
                ArmorPenetration = ArmorPenetration + TargetArmor * ObjData(.invent.EquippedWeaponObjIndex).IgnoreArmorPercent
            End If
        End If
    End With
    GetArmorPenetration = ArmorPenetration
End Function

Public Function GetEvasionBonus(ByRef User As t_User) As Integer
    GetEvasionBonus = User.Modifiers.EvasionBonus
End Function

Public Function GetHitBonus(ByRef User As t_User) As Integer
    GetHitBonus = User.Modifiers.HitBonus + GetWeaponHitBonus(User.invent.EquippedWeaponObjIndex, User.clase)
End Function

'Defines the healing bonus when using a potion, a spell or any other healing source
Public Function GetSelfHealingBonus(ByRef User As t_User) As Single
    GetSelfHealingBonus = max(1 + User.Modifiers.SelfHealingBonus, 0)
End Function

'Defines bonus when healing someone with magic
Public Function GetMagicHealingBonus(ByRef User As t_User) As Single
    GetMagicHealingBonus = max(1 + User.Modifiers.MagicHealingBonus, 0)
End Function

Public Function GetWeaponHitBonus(ByVal WeaponIndex As Integer, ByVal UserClass As e_Class)
    On Error GoTo GetWeaponHitBonus_Err
    If WeaponIndex = 0 Then Exit Function
    If Not IsFeatureEnabled("class_weapon_bonus") Or ObjData(WeaponIndex).WeaponType = 0 Then Exit Function
    GetWeaponHitBonus = ModClase(UserClass).WeaponHitBonus(ObjData(WeaponIndex).WeaponType)
    Exit Function
GetWeaponHitBonus_Err:
    Call TraceError(Err.Number, Err.Description, "UserMod.GetWeaponHitBonus WeaponIndex: " & WeaponIndex & " for class: " & UserClass, Erl)
End Function

Public Sub RemoveUserInvisibility(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        Dim RemoveHiddenState As Boolean
        ' Volver visible
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 And .flags.invisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            'Msg307=Has vuelto a ser visible.
            Call WriteLocaleMsg(UserIndex, 307, e_FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        End If
        If IsFeatureEnabled("remove-inv-on-attack") And Not MapInfo(.pos.Map).KeepInviOnAttack Then
            RemoveHiddenState = .flags.Oculto > 0 Or .flags.invisible > 0
        End If
        'I see you...
        If RemoveHiddenState And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .flags.invisible = 0
            .Counters.Invisibilidad = 0
            .Counters.TiempoOculto = 0
            .Counters.LastAttackTime = GlobalFrameTime
            If .flags.Navegando = 1 Then
                If .clase = e_Class.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call EquiparBarco(UserIndex)
                    ' Msg592=¡Has recuperado tu apariencia normal!
                    Call WriteLocaleMsg(UserIndex, 592, e_FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NoCart, NoBackPack)
                    Call RefreshCharStatus(UserIndex)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call WriteLocaleMsg(UserIndex, 307, e_FontTypeNames.FONTTYPE_INFO)
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                End If
            End If
        End If
    End With
End Sub

Public Function UserHasSpell(ByVal UserIndex As Integer, ByVal spellID As Integer) As Boolean
    With UserList(UserIndex)
        Dim i As Integer
        For i = LBound(.Stats.UserHechizos) To UBound(.Stats.UserHechizos)
            If .Stats.UserHechizos(i) = spellID Then
                UserHasSpell = True
                Exit Function
            End If
        Next i
    End With
End Function

Public Function GetLinearDamageBonus(ByVal UserIndex As Integer) As Integer
    GetLinearDamageBonus = UserList(UserIndex).Modifiers.PhysicalDamageLinearBonus
End Function

Public Function GetDefenseBonus(ByVal UserIndex As Integer) As Integer
    GetDefenseBonus = UserList(UserIndex).Modifiers.DefenseBonus
End Function

Public Function GetMaxMana(ByVal UserIndex As Integer) As Long
    With UserList(UserIndex)
        GetMaxMana = .Stats.UserAtributos(e_Atributos.Inteligencia) * ModClase(.clase).ManaInicial
        GetMaxMana = GetMaxMana + (ModClase(.clase).MultMana * .Stats.UserAtributos(e_Atributos.Inteligencia)) * (.Stats.ELV - 1)
    End With
End Function

Public Function GetHitModifier(ByVal UserIndex As Integer) As Long
    With UserList(UserIndex)
        If .Stats.ELV <= 36 Then
            GetHitModifier = (.Stats.ELV - 1) * ModClase(.clase).HitPre36
        Else
            GetHitModifier = 35 * ModClase(.clase).HitPre36
            GetHitModifier = GetHitModifier + (.Stats.ELV - 36) * ModClase(.clase).HitPost36
        End If
    End With
End Function

Public Function GetMaxStamina(ByVal UserIndex As Integer) As Integer
    With UserList(UserIndex)
        GetMaxStamina = 60 + (.Stats.ELV - 1) * ModClase(.clase).AumentoSta
    End With
End Function

Public Function GetMaxHp(ByVal UserIndex As Integer) As Integer
    With UserList(UserIndex)
        GetMaxHp = (ModClase(.clase).Vida - (21 - .Stats.UserAtributos(e_Atributos.Constitucion)) * 0.5) * (.Stats.ELV - 1) + .Stats.UserAtributos(e_Atributos.Constitucion)
    End With
End Function

Public Function GetUserSpouse(ByVal UserIndex As Integer) As String
    With UserList(UserIndex)
        If .flags.SpouseId = 0 Then
            Exit Function
        End If
        GetUserSpouse = GetUserName(.flags.SpouseId)
    End With
End Function

Public Sub RegisterNewAttack(ByVal TargetUser As Integer, ByVal attackerIndex As Integer)
    With UserList(TargetUser)
        If .Stats.MinHp > 0 Then
            Call SetUserRef(.flags.LastAttacker, attackerIndex)
            .flags.LastAttackedByUserTime = GlobalFrameTime
        End If
    End With
End Sub

Public Sub RegisterNewHelp(ByVal TargetUser As Integer, ByVal attackerIndex As Integer)
    With UserList(TargetUser)
        Call SetUserRef(.flags.LastHelpUser, attackerIndex)
        .flags.LastHelpByTime = GlobalFrameTime
    End With
End Sub

Public Sub SaveDCUserCache(ByVal UserIndex As Integer)
    On Error GoTo SaveDCUserCache_Err
    With UserList(UserIndex)
        Dim InsertIndex As Integer
        InsertIndex = RecentDCUserCache.LastIndex Mod UBound(RecentDCUserCache.LastDisconnectionInfo)
        Dim i As Integer
        For i = 0 To MaxRecentKillToStore
            RecentDCUserCache.LastDisconnectionInfo(InsertIndex).RecentKillers(i) = .flags.RecentKillers(i)
        Next i
        RecentDCUserCache.LastDisconnectionInfo(InsertIndex).RecentKillersIndex = .flags.LastKillerIndex
        RecentDCUserCache.LastDisconnectionInfo(InsertIndex).UserId = .Id
        RecentDCUserCache.LastIndex = RecentDCUserCache.LastIndex + 1
        If RecentDCUserCache.LastIndex > UBound(RecentDCUserCache.LastDisconnectionInfo) * 10 Then 'prevent overflow
            RecentDCUserCache.LastIndex = RecentDCUserCache.LastIndex \ 10
        End If
    End With
    Exit Sub
SaveDCUserCache_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SaveDCUserCache_Err", Erl)
    Resume Next
End Sub

Public Sub RestoreDCUserCache(ByVal UserIndex As Integer)
    On Error GoTo RestoreDCUserCache_Err
    With UserList(UserIndex)
        Dim StartIndex As Integer
        Dim EndIndex   As Integer
        Dim ArraySize  As Integer
        ArraySize = UBound(RecentDCUserCache.LastDisconnectionInfo)
        StartIndex = max(0, (RecentDCUserCache.LastIndex - ArraySize) Mod ArraySize)
        EndIndex = ((RecentDCUserCache.LastIndex - 1) Mod ArraySize)
        Dim i As Integer
        Dim j As Integer
        For i = StartIndex To EndIndex
            If RecentDCUserCache.LastDisconnectionInfo(i Mod ArraySize).UserId = .Id Then
                For j = 0 To MaxRecentKillToStore
                    .flags.RecentKillers(j) = RecentDCUserCache.LastDisconnectionInfo(i Mod ArraySize).RecentKillers(j)
                Next j
                .flags.LastKillerIndex = RecentDCUserCache.LastDisconnectionInfo(i Mod ArraySize).RecentKillersIndex
                Exit Sub
            End If
        Next i
    End With
    Exit Sub
RestoreDCUserCache_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.RestoreDCUserCache", Erl)
    Resume Next
End Sub

Public Function GetUserMRForNpc(ByVal UserIndex As Integer) As Integer
    With UserList(UserIndex)
        Dim MR As Integer
        MR = 0
        If .invent.EquippedArmorObjIndex > 0 Then
            MR = MR + ObjData(.invent.EquippedArmorObjIndex).ResistenciaMagica
        End If
        ' Resistencia mágica anillo
        If .invent.EquippedRingAccesoryObjIndex > 0 Then
            MR = MR + ObjData(.invent.EquippedRingAccesoryObjIndex).ResistenciaMagica
        End If
        ' Resistencia mágica escudo
        If .invent.EquippedShieldObjIndex > 0 Then
            MR = MR + ObjData(.invent.EquippedShieldObjIndex).ResistenciaMagica
        End If
        ' Resistencia mágica casco
        If .invent.EquippedHelmetObjIndex > 0 Then
            MR = MR + ObjData(.invent.EquippedHelmetObjIndex).ResistenciaMagica
        End If
        If IsFeatureEnabled("mr-magic-bonus-damage") Then
            MR = MR + .Stats.UserSkills(Resistencia) * MRSkillNpcProtectionModifier
        End If
        GetUserMRForNpc = MR + 100 * ModClase(.clase).ResistenciaMagica
    End With
End Function

Public Function GetUserMR(ByVal UserIndex As Integer) As Integer
    With UserList(UserIndex)
        Dim MR As Integer
        MR = 0
        If .invent.EquippedArmorObjIndex > 0 Then
            MR = MR + ObjData(.invent.EquippedArmorObjIndex).ResistenciaMagica
        End If
        ' Resistencia mágica anillo
        If .invent.EquippedRingAccesoryObjIndex > 0 Then
            MR = MR + ObjData(.invent.EquippedRingAccesoryObjIndex).ResistenciaMagica
        End If
        ' Resistencia mágica escudo
        If .invent.EquippedShieldObjIndex > 0 Then
            MR = MR + ObjData(.invent.EquippedShieldObjIndex).ResistenciaMagica
        End If
        ' Resistencia mágica casco
        If .invent.EquippedHelmetObjIndex > 0 Then
            MR = MR + ObjData(.invent.EquippedHelmetObjIndex).ResistenciaMagica
        End If
        If IsFeatureEnabled("mr-magic-bonus-damage") Then
            MR = MR + .Stats.UserSkills(Resistencia) * MRSkillProtectionModifier
        End If
        GetUserMR = MR + 100 * ModClase(.clase).ResistenciaMagica
    End With
End Function

Function LevelCanUseItem(ByVal UserIndex As Integer, ByRef obj As t_ObjData) As Boolean

    With UserList(UserIndex)
        If obj.MaxLEV <> 0 Then
            LevelCanUseItem = .Stats.ELV >= obj.MinELV And .Stats.ELV <= obj.MaxLEV
        Else
            LevelCanUseItem = .Stats.ELV >= obj.MinELV
        End If
    End With
    
End Function
