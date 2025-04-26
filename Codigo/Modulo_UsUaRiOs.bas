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
100     ReDim AvailableUserSlot.IndexInfo(Size)
        Dim i As Integer

102     For i = 1 To Size
104         AvailableUserSlot.IndexInfo(i) = Size - (i - 1)
106         UserList(AvailableUserSlot.IndexInfo(i)).flags.IsSlotFree = True
108     Next i

110     AvailableUserSlot.currentIndex = Size
        Exit Sub
ErrHandler_InitializeUserIndexHeap:
112     Call TraceError(Err.Number, Err.Description, "UserMod.InitializeUserIndexHeap", Erl)

End Sub

Public Function ReleaseUser(ByVal UserIndex As Integer) As Boolean
        On Error GoTo ErrHandler

100     If UserList(UserIndex).flags.IsSlotFree Then
102         ReleaseUser = False
            Exit Function

        End If

104     If IsFeatureEnabled("debug_id_assign") Then
106         Call LogError("Releasing usedid: " & UserIndex)

        End If

108     AvailableUserSlot.currentIndex = AvailableUserSlot.currentIndex + 1
110     Debug.Assert AvailableUserSlot.currentIndex <= UBound(AvailableUserSlot.IndexInfo)
112     AvailableUserSlot.IndexInfo(AvailableUserSlot.currentIndex) = UserIndex
114     UserList(UserIndex).flags.IsSlotFree = True
116     ReleaseUser = True
        Exit Function
ErrHandler:
118     ReleaseUser = False
120     Call TraceError(Err.Number, Err.Description, "UserMod.ReleaseUser", Erl)

End Function

Public Function GetAvailableUserSlot() As Integer
100     GetAvailableUserSlot = AvailableUserSlot.currentIndex

End Function

Public Function GetNextAvailableUserSlot() As Integer
        On Error GoTo ErrHandler

100     If (AvailableUserSlot.currentIndex = 0) Then
102         GetNextAvailableUserSlot = -1
104         Return

        End If

106     GetNextAvailableUserSlot = AvailableUserSlot.IndexInfo(AvailableUserSlot.currentIndex)
108     AvailableUserSlot.currentIndex = AvailableUserSlot.currentIndex - 1

110     If Not UserList(GetNextAvailableUserSlot).flags.IsSlotFree Then
112         Call TraceError(Err.Number, "Trying to active the same user slot twice", "UserMod.GetNextAvailableUserSlot", Erl)
114         GetNextAvailableUserSlot = -1

        End If

116     UserList(GetNextAvailableUserSlot).flags.IsSlotFree = False
        Exit Function
ErrHandler:
118     Call TraceError(Err.Number, Err.Description, "UserMod.GetNextAvailableUserSlot", Erl)

End Function

Public Function GetUserName(ByVal UserId As Long) As String
        On Error GoTo GetUserName_Err

100     If UserId <= 0 Then
102         GetUserName = ""
            Exit Function

        End If

104     If UserNameCache.Exists(UserId) Then
106         GetUserName = UserNameCache.Item(UserId)
            Exit Function

        End If

        Dim username As String
108     username = GetCharacterName(UserId)
110     Call RegisterUserName(UserId, username)
112     GetUserName = username
        Exit Function
GetUserName_Err:
114     Call TraceError(Err.Number, Err.Description, "UserMod.GetUserName", Erl)

End Function

Public Sub RegisterUserName(ByVal UserId As Long, ByVal username As String)

100     If UserNameCache.Exists(UserId) Then
102         UserNameCache.Item(UserId) = username
        Else
104         UserNameCache.Add UserId, username

        End If

End Sub

Public Function IsValidUserRef(ByRef UserRef As t_UserReference) As Boolean
100     IsValidUserRef = False

102     If UserRef.ArrayIndex <= 0 Or UserRef.ArrayIndex > UBound(UserList) Then
            Exit Function

        End If

104     If UserList(UserRef.ArrayIndex).VersionId <> UserRef.VersionId Then
            Exit Function

        End If

106     IsValidUserRef = True

End Function

Public Function SetUserRef(ByRef UserRef As t_UserReference, _
                           ByVal Index As Integer) As Boolean
100     SetUserRef = False
102     UserRef.ArrayIndex = Index

104     If Index <= 0 Or UserRef.ArrayIndex > UBound(UserList) Then
            Exit Function

        End If

106     UserRef.VersionId = UserList(Index).VersionId
108     SetUserRef = True

End Function

Public Sub ClearUserRef(ByRef UserRef As t_UserReference)
100     UserRef.ArrayIndex = 0
102     UserRef.VersionId = -1

End Sub

Public Sub IncreaseVersionId(ByVal UserIndex As Integer)

100     With UserList(UserIndex)

102         If .VersionId > 32760 Then
104             .VersionId = 0
            Else
106             .VersionId = .VersionId + 1

            End If

        End With

End Sub

Public Sub LogUserRefError(ByRef UserRef As t_UserReference, ByRef Text As String)
100     Call LogError("Failed to validate UserRef index(" & UserRef.ArrayIndex & ") version(" & UserRef.VersionId & ") got versionId: " & UserList(UserRef.ArrayIndex).VersionId & " At: " & Text)

End Sub

Public Function ConnectUser_Check(ByVal UserIndex As Integer, _
                                  ByVal name As String) As Boolean
        On Error GoTo Check_ConnectUser_Err
100     ConnectUser_Check = False

        'Controlamos no pasar el maximo de usuarios
102     If NumUsers >= MaxUsers Then
104         Call WriteShowMessageBox(UserIndex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
106         Call CloseSocket(UserIndex)
            Exit Function

        End If

108     If EnPausa Then
110         Call WritePauseToggle(UserIndex)
            ' Msg520=Servidor » Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.
112         Call WriteLocaleMsg(UserIndex, "520", e_FontTypeNames.FONTTYPE_SERVER)
114         Call CloseSocket(UserIndex)
            Exit Function

        End If

116     If Not EsGM(UserIndex) And ServerSoloGMs > 0 Then
118         Call WriteShowMessageBox(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
120         Call CloseSocket(UserIndex)
            Exit Function

        End If

122     With UserList(UserIndex)

124         If .flags.UserLogged Then
126             Call LogSecurity("User " & .name & " trying to log and already an already logged character from IP: " & .ConnectionDetails.IP)
128             Call CloseSocketSL(UserIndex)
130             Call Cerrar_Usuario(UserIndex)
                Exit Function

            End If

            '¿Ya esta conectado el personaje?
132         Dim tIndex As t_UserReference: tIndex = NameIndex(name)

134         If tIndex.ArrayIndex > 0 Then
136             If Not IsValidUserRef(tIndex) Then
138                 Call CloseSocket(tIndex.ArrayIndex)
140             ElseIf IsFeatureEnabled("override_same_ip_connection") And .ConnectionDetails.IP = UserList(tIndex.ArrayIndex).ConnectionDetails.IP Then
142                 Call WriteShowMessageBox(tIndex.ArrayIndex, "Alguien está ingresando con tu personaje. Si no has sido tú, por favor cambia la contraseña de tu cuenta.")
144                 Call CloseSocket(tIndex.ArrayIndex)
                Else

146                 If UserList(tIndex.ArrayIndex).Counters.Saliendo Then
148                     Call WriteShowMessageBox(UserIndex, "El personaje está saliendo.")
                    Else
150                     Call WriteShowMessageBox(UserIndex, "El personaje ya está conectado. Espere mientras es desconectado.")
                        ' Le avisamos al usuario que está jugando, en caso de que haya uno
152                     Call WriteShowMessageBox(tIndex.ArrayIndex, "Alguien está ingresando con tu personaje. Si no has sido tú, por favor cambia la contraseña de tu cuenta.")

                    End If

154                 Call CloseSocket(UserIndex)
                    Exit Function

                End If

            End If

            '¿Supera el máximo de usuarios por cuenta?
156         If MaxUsersPorCuenta > 0 Then
158             If ContarUsuariosMismaCuenta(.AccountID) >= MaxUsersPorCuenta Then
160                 If MaxUsersPorCuenta = 1 Then
162                     Call WriteShowMessageBox(UserIndex, "Ya hay un usuario conectado con esta cuenta.")
                    Else
164                     Call WriteShowMessageBox(UserIndex, "La cuenta ya alcanzó el máximo de " & MaxUsersPorCuenta & " usuarios conectados.")

                    End If

166                 Call CloseSocket(UserIndex)
                    Exit Function

                End If

            End If

168         .flags.Privilegios = UserDarPrivilegioLevel(name)

170         If EsRolesMaster(name) Then
172             .flags.Privilegios = .flags.Privilegios Or e_PlayerType.RoleMaster

            End If

174         If EsGM(UserIndex) Then
176             Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageConsoleMsg("Servidor » " & name & " se conecto al juego.", e_FontTypeNames.FONTTYPE_INFOBOLD))
178             Call LogGM(name, "Se conectó con IP: " & .ConnectionDetails.IP)

            End If

        End With

180     ConnectUser_Check = True
        Exit Function
Check_ConnectUser_Err:
182     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ConnectUser_Check", Erl)

End Function

Public Sub ConnectUser_Prepare(ByVal UserIndex As Integer, ByVal name As String)
        On Error GoTo Prepare_ConnectUser_Err

100     With UserList(UserIndex)
102         .flags.Escondido = 0
104         Call ClearNpcRef(.flags.TargetNPC)
106         .flags.TargetNpcTipo = e_NPCType.Comun
108         .flags.TargetObj = 0
110         Call SetUserRef(.flags.TargetUser, 0)
112         .Char.FX = 0
114         .Counters.CuentaRegresiva = -1
116         .name = name
            Dim UserRef As New clsUserRefWrapper
118         UserRef.SetFromIndex (UserIndex)
120         Set m_NameIndex(UCase$(name)) = UserRef
122         .showName = True
124         .NroMascotas = 0

        End With

        Exit Sub
Prepare_ConnectUser_Err:
126     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ConnectUser_Prepare", Erl)

End Sub

Public Function ConnectUser_Complete(ByVal UserIndex As Integer, _
                                     ByRef name As String, _
                                     Optional ByVal newUser As Boolean = False)
        On Error GoTo Complete_ConnectUser_Err
100     ConnectUser_Complete = False
        Dim n    As Integer
        Dim tStr As String
102     Call SendData(SendTarget.ToIndex, UserIndex, PrepareActiveToggles)

104     With UserList(UserIndex)

106         If .flags.Paralizado = 1 Then
108             .Counters.Paralisis = IntervaloParalizado

            End If

110         If .flags.Muerto = 0 Then
112             .Char = .OrigChar

114             If .Char.body = 0 Then
116                 Call SetNakedBody(UserList(UserIndex))

                End If

118             If .Char.head = 0 Then
120                 .Char.head = 1

                End If

            Else
122             .Char.body = iCuerpoMuerto
124             .Char.head = iCabezaMuerto
126             .Char.WeaponAnim = NingunArma
128             .Char.ShieldAnim = NingunEscudo
130             .Char.CascoAnim = NingunCasco
132             .Char.CartAnim = NoCart
134             .Char.Heading = e_Heading.SOUTH

            End If

136         .Stats.UserAtributos(e_Atributos.Fuerza) = 18 + ModRaza(.raza).Fuerza
138         .Stats.UserAtributos(e_Atributos.Agilidad) = 18 + ModRaza(.raza).Agilidad
140         .Stats.UserAtributos(e_Atributos.Inteligencia) = 18 + ModRaza(.raza).Inteligencia
142         .Stats.UserAtributos(e_Atributos.Constitucion) = 18 + ModRaza(.raza).Constitucion
144         .Stats.UserAtributos(e_Atributos.Carisma) = 18 + ModRaza(.raza).Carisma
146         .Stats.UserAtributosBackUP(e_Atributos.Fuerza) = .Stats.UserAtributos(e_Atributos.Fuerza)
148         .Stats.UserAtributosBackUP(e_Atributos.Agilidad) = .Stats.UserAtributos(e_Atributos.Agilidad)
150         .Stats.UserAtributosBackUP(e_Atributos.Inteligencia) = .Stats.UserAtributos(e_Atributos.Inteligencia)
152         .Stats.UserAtributosBackUP(e_Atributos.Constitucion) = .Stats.UserAtributos(e_Atributos.Constitucion)
154         .Stats.UserAtributosBackUP(e_Atributos.Carisma) = .Stats.UserAtributos(e_Atributos.Carisma)
156         .Stats.MaxMAN = UserMod.GetMaxMana(UserIndex)
158         .Stats.MaxSta = UserMod.GetMaxStamina(UserIndex)
160         .Stats.MinHIT = UserMod.GetHitModifier(UserIndex) + 1
162         .Stats.MaxHit = UserMod.GetHitModifier(UserIndex) + 2
164         .Stats.MinHp = Min(.Stats.MinHp, .Stats.MaxHp)
166         .Stats.MinMAN = Min(.Stats.MinMAN, UserMod.GetMaxMana(UserIndex))

            'Obtiene el indice-objeto del arma
168         If .invent.WeaponEqpSlot > 0 Then
170             If .invent.Object(.invent.WeaponEqpSlot).ObjIndex > 0 Then
172                 .invent.WeaponEqpObjIndex = .invent.Object(.invent.WeaponEqpSlot).ObjIndex

174                 If .flags.Muerto = 0 Then
176                     .Char.Arma_Aura = ObjData(.invent.WeaponEqpObjIndex).CreaGRH

                    End If

                Else
178                 .invent.WeaponEqpSlot = 0

                End If

            End If

            ' clear hotkey settings, the client should set this
180         For n = 0 To HotKeyCount - 1
182             .HotkeyList(n).Index = -1
184             .HotkeyList(n).LastKnownSlot = -1
186             .HotkeyList(n).Type = Unknown
188         Next n

            'Obtiene el indice-objeto del armadura
190         If .invent.ArmourEqpSlot > 0 Then
192             If .invent.Object(.invent.ArmourEqpSlot).ObjIndex > 0 Then
194                 .invent.ArmourEqpObjIndex = .invent.Object(.invent.ArmourEqpSlot).ObjIndex

196                 If .flags.Muerto = 0 Then
198                     .Char.Body_Aura = ObjData(.invent.ArmourEqpObjIndex).CreaGRH

                    End If

                Else
200                 .invent.ArmourEqpSlot = 0

                End If

202             .flags.Desnudo = 0
            Else
204             .flags.Desnudo = 1

            End If

            'Obtiene el indice-objeto del escudo
206         If .invent.EscudoEqpSlot > 0 Then
208             If .invent.Object(.invent.EscudoEqpSlot).ObjIndex > 0 Then
210                 .invent.EscudoEqpObjIndex = .invent.Object(.invent.EscudoEqpSlot).ObjIndex

212                 If .flags.Muerto = 0 Then
214                     .Char.Escudo_Aura = ObjData(.invent.EscudoEqpObjIndex).CreaGRH

                    End If

                Else
216                 .invent.EscudoEqpSlot = 0

                End If

            End If

            'Obtiene el indice-objeto del casco
218         If .invent.CascoEqpSlot > 0 Then
220             If .invent.Object(.invent.CascoEqpSlot).ObjIndex > 0 Then
222                 .invent.CascoEqpObjIndex = .invent.Object(.invent.CascoEqpSlot).ObjIndex

224                 If .flags.Muerto = 0 Then
226                     .Char.Head_Aura = ObjData(.invent.CascoEqpObjIndex).CreaGRH

                    End If

                Else
228                 .invent.CascoEqpSlot = 0

                End If

            End If

            'Obtiene el indice-objeto barco
230         If .invent.BarcoSlot > 0 Then
232             If .invent.Object(.invent.BarcoSlot).ObjIndex > 0 Then
234                 .invent.BarcoObjIndex = .invent.Object(.invent.BarcoSlot).ObjIndex
                Else
236                 .invent.BarcoSlot = 0

                End If

            End If

            'Obtiene el indice-objeto municion
238         If .invent.MunicionEqpSlot > 0 Then
240             If .invent.Object(.invent.MunicionEqpSlot).ObjIndex > 0 Then
242                 .invent.MunicionEqpObjIndex = .invent.Object(.invent.MunicionEqpSlot).ObjIndex
                Else
244                 .invent.MunicionEqpSlot = 0

                End If

            End If

            ' DM
246         If .invent.DañoMagicoEqpSlot > 0 Then
248             If .invent.Object(.invent.DañoMagicoEqpSlot).ObjIndex > 0 Then
250                 .invent.DañoMagicoEqpObjIndex = .invent.Object(.invent.DañoMagicoEqpSlot).ObjIndex

252                 If .flags.Muerto = 0 Then
254                     .Char.DM_Aura = ObjData(.invent.DañoMagicoEqpObjIndex).CreaGRH

                    End If

                Else
256                 .invent.DañoMagicoEqpSlot = 0

                End If

            End If

258         If .invent.MagicoSlot > 0 Then
260             .invent.MagicoObjIndex = .invent.Object(.invent.MagicoSlot).ObjIndex

262             If ObjData(.invent.MagicoObjIndex).CreaGRH <> "" Then
264                 .Char.Otra_Aura = ObjData(.invent.MagicoObjIndex).CreaGRH
266                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charIndex, .Char.Otra_Aura, False, 5))

                End If

268             If ObjData(.invent.MagicoObjIndex).Ropaje > 0 Then
270                 .Char.CartAnim = ObjData(.invent.MagicoObjIndex).Ropaje

                End If

            End If

            ' RM
272         If .invent.ResistenciaEqpSlot > 0 Then
274             If .invent.Object(.invent.ResistenciaEqpSlot).ObjIndex > 0 Then
276                 .invent.ResistenciaEqpObjIndex = .invent.Object(.invent.ResistenciaEqpSlot).ObjIndex

278                 If .flags.Muerto = 0 Then
280                     .Char.RM_Aura = ObjData(.invent.ResistenciaEqpObjIndex).CreaGRH

                    End If

                Else
282                 .invent.ResistenciaEqpSlot = 0

                End If

            End If

284         If .invent.MonturaSlot > 0 Then
286             If .invent.Object(.invent.MonturaSlot).ObjIndex > 0 Then
288                 .invent.MonturaObjIndex = .invent.Object(.invent.MonturaSlot).ObjIndex
                Else
290                 .invent.MonturaSlot = 0

                End If

            End If

292         If .invent.HerramientaEqpSlot > 0 Then
294             If .invent.Object(.invent.HerramientaEqpSlot).ObjIndex Then
296                 .invent.HerramientaEqpObjIndex = .invent.Object(.invent.HerramientaEqpSlot).ObjIndex
                Else
298                 .invent.HerramientaEqpSlot = 0

                End If

            End If

300         If .invent.MagicoSlot > 0 Then
302             If .invent.Object(.invent.MagicoSlot).ObjIndex Then
304                 .invent.MagicoObjIndex = .invent.Object(.invent.MagicoSlot).ObjIndex

306                 If .flags.Muerto = 0 Then
308                     .Char.Otra_Aura = ObjData(.invent.MagicoObjIndex).CreaGRH

                    End If

                Else
310                 .invent.MagicoSlot = 0

                End If

            End If

312         If .invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
314         If .invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
316         If .invent.WeaponEqpSlot = 0 And .invent.HerramientaEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
318         If .invent.MagicoSlot = 0 Then .Char.CartAnim = NoCart

            ' -----------------------------------------------------------------------
            '   FIN - INFORMACION INICIAL DEL PERSONAJE
            ' -----------------------------------------------------------------------
320         If Not ValidateChr(UserIndex) Then
322             Call WriteShowMessageBox(UserIndex, "Error en el personaje. Comuniquese con el staff.")
324             Call CloseSocket(UserIndex)
                Exit Function

            End If

326         .flags.SeguroParty = True
328         .flags.SeguroClan = True
330         .flags.SeguroResu = True
332         .CurrentInventorySlots = getMaxInventorySlots(UserIndex)
334         Call WriteInventoryUnlockSlots(UserIndex)
336         Call LoadUserIntervals(UserIndex)
338         Call WriteIntervals(UserIndex)
340         Call UpdateUserInv(True, UserIndex, 0)
342         Call UpdateUserHechizos(True, UserIndex, 0)
344         Call EnviarLlaves(UserIndex)

346         If .flags.Paralizado Then Call WriteParalizeOK(UserIndex)
348         If .flags.Inmovilizado Then Call WriteInmovilizaOK(UserIndex)

            ''
            'TODO : Feo, esto tiene que ser parche cliente
350         If .flags.Estupidez = 0 Then
352             Call WriteDumbNoMore(UserIndex)

            End If

354         .flags.Inmunidad = 1
356         .Counters.TiempoDeInmunidad = IntervaloPuedeSerAtacado
358         .Counters.TiempoDeInmunidadParalisisNoMagicas = 0

360         If MapInfo(.pos.Map).MapResource = 0 Then
362             .pos.Map = Ciudades(.Hogar).Map
364             .pos.x = Ciudades(.Hogar).x
366             .pos.y = Ciudades(.Hogar).y

            End If

            'Mapa válido
368         If Not MapaValido(.pos.Map) Then
370             Call WriteErrorMsg(UserIndex, "EL PJ se encuenta en un mapa invalido.")
372             Call CloseSocket(UserIndex)
                Exit Function

            End If

            'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
374         If MapData(.pos.Map, .pos.x, .pos.y).UserIndex <> 0 Or MapData(.pos.Map, .pos.x, .pos.y).NpcIndex <> 0 Then
                Dim FoundPlace As Boolean
                Dim esAgua     As Boolean
                Dim tX         As Long
                Dim tY         As Long
376             FoundPlace = False
378             esAgua = (MapData(.pos.Map, .pos.x, .pos.y).Blocked And FLAG_AGUA) <> 0

380             For tY = .pos.y - 1 To .pos.y + 1
382                 For tX = .pos.x - 1 To .pos.x + 1

384                     If esAgua Then

                            'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
386                         If LegalPos(.pos.Map, tX, tY, True, True, False, False, False) Then
388                             FoundPlace = True
                                Exit For

                            End If

                        Else

                            'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
390                         If LegalPos(.pos.Map, tX, tY, False, True, False, False, False) Then
392                             FoundPlace = True
                                Exit For

                            End If

                        End If

394                 Next tX

396                 If FoundPlace Then Exit For
398             Next tY

400             If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
402                 .pos.x = tX
404                 .pos.y = tY
                Else

                    'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
406                 If MapData(.pos.Map, .pos.x, .pos.y).UserIndex <> 0 Then

                        'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
408                     If IsValidUserRef(UserList(MapData(.pos.Map, .pos.x, .pos.y).UserIndex).ComUsu.DestUsu) Then

                            'Le avisamos al que estaba comerciando que se tuvo que ir.
410                         If UserList(UserList(MapData(.pos.Map, .pos.x, .pos.y).UserIndex).ComUsu.DestUsu.ArrayIndex).flags.UserLogged Then
412                             Call FinComerciarUsu(UserList(MapData(.pos.Map, .pos.x, .pos.y).UserIndex).ComUsu.DestUsu.ArrayIndex)
414                             Call WriteConsoleMsg(UserList(MapData(.pos.Map, .pos.x, .pos.y).UserIndex).ComUsu.DestUsu.ArrayIndex, "Comercio cancelado. El otro usuario se ha desconectado.", e_FontTypeNames.FONTTYPE_WARNING)

                            End If

                            'Lo sacamos.
416                         If UserList(MapData(.pos.Map, .pos.x, .pos.y).UserIndex).flags.UserLogged Then
418                             Call FinComerciarUsu(MapData(.pos.Map, .pos.x, .pos.y).UserIndex)
420                             Call WriteErrorMsg(MapData(.pos.Map, .pos.x, .pos.y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconectate...")

                            End If

                        End If

422                     Call CloseSocket(MapData(.pos.Map, .pos.x, .pos.y).UserIndex)

                    End If

                End If

            End If

            'If in the water, and has a boat, equip it!
            Dim trigger     As Integer
            Dim slotBarco   As Integer
            Dim itemBuscado As Integer
424         trigger = MapData(.pos.Map, .pos.x, .pos.y).trigger

426         If trigger = e_Trigger.DETALLEAGUA Then 'Esta en zona de caucho obj 199, 200
428             If .raza = e_Raza.Enano Or .raza = e_Raza.Gnomo Then
430                 itemBuscado = iObjTrajeBajoNw
                Else
432                 itemBuscado = iObjTrajeAltoNw

                End If

434             slotBarco = GetSlotInInventory(UserIndex, itemBuscado)

436             If slotBarco > -1 Then
438                 .invent.BarcoObjIndex = itemBuscado
440                 .invent.BarcoSlot = slotBarco

                End If

442         ElseIf trigger = e_Trigger.VALIDONADO Or trigger = e_Trigger.NADOCOMBINADO Then  'Esta en zona de nado comun obj 197
444             itemBuscado = iObjTraje
446             slotBarco = GetSlotInInventory(UserIndex, itemBuscado)

448             If slotBarco > -1 Then
450                 .invent.BarcoObjIndex = itemBuscado
452                 .invent.BarcoSlot = slotBarco

                End If

            End If

454         If .invent.BarcoObjIndex > 0 And (MapData(.pos.Map, .pos.x, .pos.y).Blocked And FLAG_AGUA) <> 0 Then
456             .flags.Navegando = 1
458             Call EquiparBarco(UserIndex)
460         ElseIf .flags.Navegando = 1 And (MapData(.pos.Map, .pos.x, .pos.y).Blocked And FLAG_AGUA) <> 0 Then
                Dim iSlot As Integer

462             For iSlot = 1 To UBound(.invent.Object)

464                 If .invent.Object(iSlot).ObjIndex > 0 Then
466                     If ObjData(.invent.Object(iSlot).ObjIndex).OBJType = otBarcos And ObjData(.invent.Object(iSlot).ObjIndex).Subtipo > 0 Then
468                         .invent.BarcoObjIndex = .invent.Object(iSlot).ObjIndex
470                         .invent.BarcoSlot = iSlot
                            Exit For

                        End If

                    End If

                Next

            End If

472         If .invent.MagicoObjIndex <> 0 Then
474             If ObjData(.invent.MagicoObjIndex).EfectoMagico = 11 Then .flags.Paraliza = 1

            End If

476         Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
478         Call WriteHora(UserIndex)
480         Call WriteChangeMap(UserIndex, .pos.Map) 'Carga el mapa
482         Call UpdateCharWithEquipedItems(UserIndex)

484         Select Case .flags.Privilegios

                Case e_PlayerType.Admin
486                 .flags.ChatColor = RGB(252, 195, 0)

488             Case e_PlayerType.Dios
490                 .flags.ChatColor = RGB(26, 209, 107)

492             Case e_PlayerType.SemiDios
494                 .flags.ChatColor = RGB(60, 150, 60)

496             Case e_PlayerType.Consejero
498                 .flags.ChatColor = RGB(170, 170, 170)

500             Case Else
502                 .flags.ChatColor = vbWhite

            End Select

504         Select Case .Faccion.Status

                Case e_Facciones.Ciudadano
506                 .flags.ChatColor = vbWhite

508             Case e_Facciones.Armada
510                 .flags.ChatColor = vbWhite

512             Case e_Facciones.consejo
514                 .flags.ChatColor = RGB(66, 201, 255)

516             Case e_Facciones.Criminal
518                 .flags.ChatColor = vbWhite

520             Case e_Facciones.Caos
522                 .flags.ChatColor = vbWhite

524             Case e_Facciones.concilio
526                 .flags.ChatColor = RGB(255, 102, 102)

            End Select

            ' Jopi: Te saco de los mapas de retos (si logueas ahi) 324 372 389 390
528         If Not EsGM(UserIndex) And (.pos.Map = 324 Or .pos.Map = 372 Or .pos.Map = 389 Or .pos.Map = 390) Then

                ' Si tiene una posicion a la que volver, lo mando ahi
530             If MapaValido(.flags.ReturnPos.Map) And .flags.ReturnPos.x > 0 And .flags.ReturnPos.x <= XMaxMapSize And .flags.ReturnPos.y > 0 And .flags.ReturnPos.y <= YMaxMapSize Then
532                 Call WarpToLegalPos(UserIndex, .flags.ReturnPos.Map, .flags.ReturnPos.x, .flags.ReturnPos.y, True)
                Else ' Lo mando a su hogar
534                 Call WarpToLegalPos(UserIndex, Ciudades(.Hogar).Map, Ciudades(.Hogar).x, Ciudades(.Hogar).y, True)

                End If

            End If

            ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
            #If ConUpTime Then
536             .LogOnTime = Now
            #End If
            'Crea  el personaje del usuario
538         Call MakeUserChar(True, .pos.Map, UserIndex, .pos.Map, .pos.x, .pos.y, 1)
540         Call WriteUserCharIndexInServer(UserIndex)
542         Call ActualizarVelocidadDeUsuario(UserIndex)

544         If .flags.Privilegios And (e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.Admin) Then
546             Call DoAdminInvisible(UserIndex)

            End If

548         Call WriteUpdateUserStats(UserIndex)
550         Call WriteUpdateHungerAndThirst(UserIndex)
552         Call WriteUpdateDM(UserIndex)
554         Call WriteUpdateRM(UserIndex)
556         Call SendMOTD(UserIndex)
            'Actualiza el Num de usuarios
558         NumUsers = NumUsers + 1
560         .flags.UserLogged = True
562         Call Execute("Update user set is_logged = true where id = ?", UserList(UserIndex).Id)
564         .Counters.LastSave = GetTickCount
566         MapInfo(.pos.Map).NumUsers = MapInfo(.pos.Map).NumUsers + 1

568         If .Stats.SkillPts > 0 Then
570             Call WriteSendSkills(UserIndex)
572             Call WriteLevelUp(UserIndex, .Stats.SkillPts)

            End If

574         If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers
576         If NumUsers > RecordUsuarios Then
578             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultáneamente: " & NumUsers & " usuarios.", e_FontTypeNames.FONTTYPE_INFO))
580             RecordUsuarios = NumUsers

            End If

582         Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageOnlineUser(NumUsers))
584         Call WriteFYA(UserIndex)
586         Call WriteBindKeys(UserIndex)

588         If .NroMascotas > 0 And MapInfo(.pos.Map).NoMascotas = 0 And .flags.MascotasGuardadas = 0 Then
                Dim i As Integer

590             For i = 1 To MAXMASCOTAS

592                 If .MascotasType(i) > 0 Then
594                     Call SetNpcRef(.MascotasIndex(i), SpawnNpc(.MascotasType(i), .pos, False, False, False, UserIndex))

596                     If .MascotasIndex(i).ArrayIndex > 0 Then
598                         Call SetUserRef(NpcList(.MascotasIndex(i).ArrayIndex).MaestroUser, UserIndex)
600                         Call FollowAmo(.MascotasIndex(i).ArrayIndex)

                        End If

                    End If

602             Next i

            End If

604         If .flags.Montado = 1 Then
606             Call WriteEquiteToggle(UserIndex)

            End If

608         Call ActualizarVelocidadDeUsuario(UserIndex)

610         If .GuildIndex > 0 Then

                'welcome to the show baby...
612             If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
                    ' Msg521=Tu estado no te permite entrar al clan.
614                 Call WriteLocaleMsg(UserIndex, "521", e_FontTypeNames.FONTTYPE_GUILD)

                End If

            End If

616         If LenB(.LastGuildRejection) <> 0 Then
618             Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & .LastGuildRejection)
620             .LastGuildRejection = vbNullString
622             Call SaveUserGuildRejectionReason(.name, vbNullString)

            End If

624         If Lloviendo Then Call WriteRainToggle(UserIndex)
626         If ServidorNublado Then Call WriteNubesToggle(UserIndex)
628         Call WriteLoggedMessage(UserIndex, newUser)

630         If .Stats.ELV = 1 Then
                ' Msg522=¡Bienvenido a las tierras de Argentum Online! ¡<nombre> que tengas buen viaje y mucha suerte!
632             Call WriteLocaleMsg(UserIndex, "522", e_FontTypeNames.FONTTYPE_GUILD, .name)
634         ElseIf .Stats.ELV < 14 Then
636             Call WriteLocaleMsg(UserIndex, "1439", e_FontTypeNames.FONTTYPE_GUILD, .name) ' Msg1439=¡Bienvenido de nuevo ¬1! Actualmente estas en el nivel ¬2 en ¬3, ¡buen viaje y mucha suerte!

            End If

638         If Status(UserIndex) = Criminal Or Status(UserIndex) = e_Facciones.Caos Then
640             Call WriteSafeModeOff(UserIndex)
642             .flags.Seguro = False
            Else
644             .flags.Seguro = True
646             Call WriteSafeModeOn(UserIndex)

            End If

648         If LenB(.MENSAJEINFORMACION) > 0 Then
                Dim Lines() As String
650             Lines = Split(.MENSAJEINFORMACION, vbNewLine)

652             For i = 0 To UBound(Lines)

654                 If LenB(Lines(i)) > 0 Then
656                     Call WriteConsoleMsg(UserIndex, Lines(i), e_FontTypeNames.FONTTYPE_New_DONADOR)

                    End If

                Next
658             .MENSAJEINFORMACION = vbNullString

            End If

660         If EventoActivo Then
662             Call WriteConsoleMsg(UserIndex, PublicidadEvento & ". Tiempo restante: " & TiempoRestanteEvento & " minuto(s).", e_FontTypeNames.FONTTYPE_New_Eventos)

            End If

664         Call WriteContadores(UserIndex)
666         Call WritePrivilegios(UserIndex)
668         Call RestoreDCUserCache(UserIndex)
670         Call CustomScenarios.UserConnected(UserIndex)
672         Call AntiCheat.OnNewPlayerConnect(UserIndex)

        End With

674     ConnectUser_Complete = True
        Exit Function
Complete_ConnectUser_Err:
676     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ConnectUser_Complete", Erl)

End Function

Sub ActStats(ByVal VictimIndex As Integer, ByVal attackerIndex As Integer)
        On Error GoTo ActStats_Err
        Dim DaExp       As Integer
        Dim EraCriminal As Byte
100     DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)

102     If UserList(attackerIndex).Stats.ELV < STAT_MAXELV Then
104         UserList(attackerIndex).Stats.Exp = UserList(attackerIndex).Stats.Exp + DaExp

106         If UserList(attackerIndex).Stats.Exp > MAXEXP Then UserList(attackerIndex).Stats.Exp = MAXEXP
108         Call WriteUpdateExp(attackerIndex)
110         Call CheckUserLevel(attackerIndex)

        End If

112     Call WriteLocaleMsg(attackerIndex, "184", e_FontTypeNames.FONTTYPE_FIGHT, UserList(VictimIndex).name)
114     Call WriteLocaleMsg(attackerIndex, "140", e_FontTypeNames.FONTTYPE_EXP, DaExp)
116     Call WriteLocaleMsg(VictimIndex, "185", e_FontTypeNames.FONTTYPE_FIGHT, UserList(attackerIndex).name)

118     If Not PeleaSegura(VictimIndex, attackerIndex) Then
120         EraCriminal = Status(attackerIndex)

122         If EraCriminal = 2 And Status(attackerIndex) < 2 Then
124             Call RefreshCharStatus(attackerIndex)
126         ElseIf EraCriminal < 2 And Status(attackerIndex) = 2 Then
128             Call RefreshCharStatus(attackerIndex)

            End If

        End If

130     Call UserMod.UserDie(VictimIndex)

132     If TriggerZonaPelea(attackerIndex, attackerIndex) <> TRIGGER6_PERMITE Then
134         If UserList(attackerIndex).Stats.UsuariosMatados < MAXUSERMATADOS Then
136             UserList(attackerIndex).Stats.UsuariosMatados = UserList(attackerIndex).Stats.UsuariosMatados + 1

            End If

        End If

        Exit Sub
ActStats_Err:
138     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ActStats", Erl)

End Sub

Sub RevivirUsuario(ByVal UserIndex As Integer, Optional ByVal MedianteHechizo As Boolean)
        On Error GoTo RevivirUsuario_Err

100     With UserList(UserIndex)
102         .flags.Muerto = 0
104         .Stats.MinHp = .Stats.MaxHp

            ' El comportamiento cambia si usamos el hechizo Resucitar
106         If MedianteHechizo Then
108             .Stats.MinHp = 1
110             .Stats.MinHam = 0
112             .Stats.MinAGU = 0
114             .Stats.MinMAN = 0
116             Call WriteUpdateHungerAndThirst(UserIndex)

            End If

118         Call WriteUpdateHP(UserIndex)
120         Call WriteUpdateMana(UserIndex)

122         If .flags.Navegando = 1 Then
124             Call EquiparBarco(UserIndex)
            Else
126             .Char.head = .OrigChar.head

128             If .invent.CascoEqpObjIndex > 0 Then
130                 .Char.CascoAnim = ObjData(.invent.CascoEqpObjIndex).CascoAnim

                End If

132             If .invent.EscudoEqpObjIndex > 0 Then
134                 .Char.ShieldAnim = ObjData(.invent.EscudoEqpObjIndex).ShieldAnim

                End If

136             If .invent.WeaponEqpObjIndex > 0 Then
138                 .Char.WeaponAnim = ObjData(.invent.WeaponEqpObjIndex).WeaponAnim

140                 If ObjData(.invent.WeaponEqpObjIndex).CreaGRH <> "" Then
142                     .Char.Arma_Aura = ObjData(.invent.WeaponEqpObjIndex).CreaGRH
144                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charIndex, .Char.Arma_Aura, False, 1))

                    End If

                End If

146             If .invent.ArmourEqpObjIndex > 0 Then
148                 .Char.body = ObtenerRopaje(UserIndex, ObjData(.invent.ArmourEqpObjIndex))

150                 If ObjData(.invent.ArmourEqpObjIndex).CreaGRH <> "" Then
152                     .Char.Body_Aura = ObjData(.invent.ArmourEqpObjIndex).CreaGRH
154                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charIndex, .Char.Body_Aura, False, 2))

                    End If

                Else
156                 Call SetNakedBody(UserList(UserIndex))

                End If

158             If .invent.EscudoEqpObjIndex > 0 Then
160                 .Char.ShieldAnim = ObjData(.invent.EscudoEqpObjIndex).ShieldAnim

162                 If ObjData(.invent.EscudoEqpObjIndex).CreaGRH <> "" Then
164                     .Char.Escudo_Aura = ObjData(.invent.EscudoEqpObjIndex).CreaGRH
166                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charIndex, .Char.Escudo_Aura, False, 3))

                    End If

                End If

168             If .invent.CascoEqpObjIndex > 0 Then
170                 .Char.CascoAnim = ObjData(.invent.CascoEqpObjIndex).CascoAnim

172                 If ObjData(.invent.CascoEqpObjIndex).CreaGRH <> "" Then
174                     .Char.Head_Aura = ObjData(.invent.CascoEqpObjIndex).CreaGRH
176                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charIndex, .Char.Head_Aura, False, 4))

                    End If

                End If

178             If .invent.MagicoObjIndex > 0 Then
180                 If ObjData(.invent.MagicoObjIndex).CreaGRH <> "" Then
182                     .Char.Otra_Aura = ObjData(.invent.MagicoObjIndex).CreaGRH
184                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charIndex, .Char.Otra_Aura, False, 5))

                    End If

186                 If ObjData(.invent.MagicoObjIndex).Ropaje > 0 Then
188                     .Char.CartAnim = ObjData(.invent.MagicoObjIndex).Ropaje

                    End If

                End If

190             If .invent.DañoMagicoEqpObjIndex > 0 Then
192                 If ObjData(.invent.DañoMagicoEqpObjIndex).CreaGRH <> "" Then
194                     .Char.DM_Aura = ObjData(.invent.DañoMagicoEqpObjIndex).CreaGRH
196                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charIndex, .Char.DM_Aura, False, 6))

                    End If

                End If

198             If .invent.ResistenciaEqpObjIndex > 0 Then
200                 If ObjData(.invent.ResistenciaEqpObjIndex).CreaGRH <> "" Then
202                     .Char.RM_Aura = ObjData(.invent.ResistenciaEqpObjIndex).CreaGRH
204                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charIndex, .Char.RM_Aura, False, 7))

                    End If

                End If

            End If

206         Call ActualizarVelocidadDeUsuario(UserIndex)
208         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)
210         Call MakeUserChar(True, UserList(UserIndex).pos.Map, UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, 0)

        End With

        Exit Sub
RevivirUsuario_Err:
212     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.RevivirUsuario", Erl)

End Sub

Sub ChangeUserChar(ByVal UserIndex As Integer, _
                   ByVal body As Integer, _
                   ByVal head As Integer, _
                   ByVal Heading As Byte, _
                   ByVal Arma As Integer, _
                   ByVal Escudo As Integer, _
                   ByVal Casco As Integer, _
                   ByVal Cart As Integer)
        On Error GoTo ChangeUserChar_Err

100     If IsSet(UserList(UserIndex).flags.StatusMask, e_StatusMask.eTransformed) Then Exit Sub

102     With UserList(UserIndex).Char
104         .body = body
106         .head = head
108         .Heading = Heading
110         .WeaponAnim = Arma
112         .ShieldAnim = Escudo
114         .CascoAnim = Casco
116         .CartAnim = Cart

        End With

118     If UserList(UserIndex).Char.charIndex > 0 Then
120         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(body, head, Heading, UserList(UserIndex).Char.charIndex, Arma, Escudo, Cart, UserList(UserIndex).Char.FX, UserList(UserIndex).Char.loops, Casco, False, UserList(UserIndex).flags.Navegando))

        End If

        Exit Sub
ChangeUserChar_Err:
122     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ChangeUserChar", Erl)

End Sub

Sub EraseUserChar(ByVal UserIndex As Integer, _
                  ByVal Desvanecer As Boolean, _
                  Optional ByVal FueWarp As Boolean = False)
        On Error GoTo ErrorHandler
        Dim Error As String
100     Error = "1"

102     If UserList(UserIndex).Char.charIndex = 0 Then Exit Sub
104     CharList(UserList(UserIndex).Char.charIndex) = 0

106     If UserList(UserIndex).Char.charIndex = LastChar Then

108         Do Until CharList(LastChar) > 0
110             LastChar = LastChar - 1

112             If LastChar <= 1 Then Exit Do
            Loop

        End If

114     Error = "2"
        #If UNIT_TEST = 0 Then
            'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
116         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(4, UserList(UserIndex).Char.charIndex, Desvanecer, FueWarp))
118         Error = "3"
120         Call QuitarUser(UserIndex, UserList(UserIndex).pos.Map)
122         Error = "4"
        #End If
124     MapData(UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y).UserIndex = 0
126     Error = "5"
128     UserList(UserIndex).Char.charIndex = 0
130     NumChars = NumChars - 1
132     Error = "6"
        Exit Sub
ErrorHandler:
134     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.EraseUserChar", Erl)

End Sub

Sub RefreshCharStatus(ByVal UserIndex As Integer)
        On Error GoTo RefreshCharStatus_Err
        'Refreshes the status and tag of UserIndex.
        Dim klan As String, name As String

100     If UserList(UserIndex).showName Then
102         If UserList(UserIndex).flags.Mimetizado = e_EstadoMimetismo.Desactivado Then
104             If UserList(UserIndex).GuildIndex > 0 Then
106                 klan = modGuilds.GuildName(UserList(UserIndex).GuildIndex)
108                 klan = " <" & klan & ">"

                End If

110             name = UserList(UserIndex).name & klan
            Else
112             name = UserList(UserIndex).NameMimetizado

            End If

114         If UserList(UserIndex).clase = e_Class.Pirat Then
116             If UserList(UserIndex).flags.Oculto = 1 Then
118                 name = vbNullString

                End If

            End If

        End If

120     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, UserList(UserIndex).Faccion.Status, name))
        Exit Sub
RefreshCharStatus_Err:
122     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.RefreshCharStatus", Erl)

End Sub

Sub MakeUserChar(ByVal toMap As Boolean, _
                 ByVal sndIndex As Integer, _
                 ByVal UserIndex As Integer, _
                 ByVal Map As Integer, _
                 ByVal x As Integer, _
                 ByVal y As Integer, _
                 Optional ByVal appear As Byte = 0)
        On Error GoTo HayError
        Dim charIndex As Integer
        Dim TempName  As String

100     If InMapBounds(Map, x, y) Then

102         With UserList(UserIndex)

                'If needed make a new character in list
104             If .Char.charIndex = 0 Then
106                 charIndex = NextOpenCharIndex
108                 .Char.charIndex = charIndex
110                 CharList(charIndex) = UserIndex

112                 If .Grupo.EnGrupo Then
114                     Call modSendData.SendData(ToGroup, UserIndex, PrepareUpdateGroupInfo(UserIndex))

                    End If

                End If

                'Place character on map if needed
116             If toMap Then MapData(Map, x, y).UserIndex = UserIndex
                'Send make character command to clients
                Dim klan       As String
                Dim clan_nivel As Byte

118             If Not toMap Then
120                 If .showName Then
122                     If .flags.Mimetizado = e_EstadoMimetismo.Desactivado Then
124                         If .GuildIndex > 0 Then
126                             klan = modGuilds.GuildName(.GuildIndex)
128                             clan_nivel = modGuilds.NivelDeClan(.GuildIndex)
130                             TempName = .name & " <" & klan & ">"
                            Else
132                             klan = vbNullString
134                             clan_nivel = 0

136                             If .flags.EnConsulta Then
138                                 TempName = .name & " [CONSULTA]"
                                Else
140                                 TempName = .name

                                End If

                            End If

                        Else
142                         TempName = .NameMimetizado

                        End If

                    End If

144                 Call WriteCharacterCreate(sndIndex, .Char.body, .Char.head, .Char.Heading, .Char.charIndex, x, y, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CartAnim, .Char.FX, 999, .Char.CascoAnim, TempName, .Faccion.Status, .flags.Privilegios, .Char.ParticulaFx, .Char.Head_Aura, .Char.Arma_Aura, .Char.Body_Aura, .Char.DM_Aura, .Char.RM_Aura, .Char.Otra_Aura, .Char.Escudo_Aura, .Char.speeding, 0, appear, .Grupo.Lider.ArrayIndex, .GuildIndex, clan_nivel, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMAN, .Stats.MaxMAN, 0, False, .flags.Navegando, .Stats.tipoUsuario, .flags.CurrentTeam, .flags.tiene_bandera)
                Else
                    'Hide the name and clan - set privs as normal user
146                 Call AgregarUser(UserIndex, .pos.Map, appear)

                End If

            End With

        End If

        Exit Sub
HayError:
        Dim Desc As String
148     Desc = Err.Description & vbNewLine & " Usuario: " & UserList(UserIndex).name & vbNewLine & "Pos: " & Map & "-" & x & "-" & y
150     Call TraceError(Err.Number, Err.Description, "Usuarios.MakeUserChar", Erl())
152     Call CloseSocket(UserIndex)

End Sub

Sub CheckUserLevel(ByVal UserIndex As Integer)
        'Chequea que el usuario no halla alcanzado el siguiente nivel,
        'de lo contrario le da la vida, mana, etc, correspodiente.
        'Integer - Modificacion de los valores
        'Corregido el BUG de STAT_MAXELV
        'Agrego modificaciones en ELU al subir de nivel.
        'Agrego modificaciones de la subida de mana de los magos por lvl.
        'Agrego diferencias entre el 18 y el 19 en Constitución.
        'Ahora el incremento de vida por Consitución se controla desde Balance.dat
        'Distribución normal de las vidas
        'Vuelvo a implementar vidas variables y les agrego un capeo min/max
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
100     Randomize Time

102     With UserList(UserIndex)
104         WasNewbie = EsNewbie(UserIndex)
106         experienceToLevelUp = ExpLevelUp(.Stats.ELV)

108         Do While .Stats.Exp >= experienceToLevelUp And .Stats.ELV < STAT_MAXELV
                'Store it!
                'Call Statistics.UserLevelUp(UserIndex)
110             UserList(UserIndex).Counters.timeFx = 3
112             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charIndex, 106, 0, .pos.x, .pos.y))
114             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .pos.x, .pos.y))
116             Call WriteLocaleMsg(UserIndex, "186", e_FontTypeNames.FONTTYPE_INFO)
118             .Stats.Exp = .Stats.Exp - experienceToLevelUp
120             Pts = Pts + ModClase(.clase).LevelSkillPoints
122             .Stats.ELV = .Stats.ELV + 1
124             experienceToLevelUp = ExpLevelUp(.Stats.ELV)
126             AumentoSta = .Stats.MaxSta
128             AumentoMANA = .Stats.MaxMAN
130             AumentoHIT = .Stats.MaxHit
132             .Stats.MaxMAN = UserMod.GetMaxMana(UserIndex)
134             .Stats.MaxSta = UserMod.GetMaxStamina(UserIndex)
136             .Stats.MinHIT = UserMod.GetHitModifier(UserIndex) + 1
138             .Stats.MaxHit = UserMod.GetHitModifier(UserIndex) + 2
140             AumentoSta = .Stats.MaxSta - AumentoSta
142             AumentoMANA = .Stats.MaxMAN - AumentoMANA
144             AumentoHIT = .Stats.MaxHit - AumentoHIT
                ' Shugar 15/7/2024
                ' Devuelvo el aumento de vida variable pero con capeo min/max
                ' Promedio sin vida variable
146             PromClaseRaza = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(e_Atributos.Constitucion)) * 0.5
                ' Promedio real del personaje
148             PromPersonaje = CalcularPromedioVida(UserIndex)
                ' Sesgo a favor del promedio sin vida variable
                ' Si DesbalancePromedioVidas = 0, el PromBias es el PromClaseRaza del manual
150             PromBias = PromClaseRaza + (PromClaseRaza - PromPersonaje) * DesbalancePromedioVidas
                ' Aumenta la vida un número entero al azar en un rango dado
                ' Min: PromClaseRaza - RangoVidas
                ' Max: PromClaseRaza + RangoVidas
                ' Media: PromBias
                ' Desviación: InfluenciaPromedioVidas
152             AumentoHP = RandomIntBiased(PromClaseRaza - RangoVidas, PromClaseRaza + RangoVidas, PromBias, InfluenciaPromedioVidas)

                ' Capeo de vida máxima a +10
154             If .Stats.MaxHp + AumentoHP > UserMod.GetMaxHp(UserIndex) + CapVidaMax Then
156                 AumentoHP = (UserMod.GetMaxHp(UserIndex) + CapVidaMax) - .Stats.MaxHp

                End If

                ' Capeo de vida mínima a -10
158             If .Stats.MaxHp + AumentoHP < UserMod.GetMaxHp(UserIndex) + CapVidaMin Then
160                 AumentoHP = (UserMod.GetMaxHp(UserIndex) + CapVidaMin) - .Stats.MaxHp

                End If

                ' Aumento la vida máxima del personaje
162             .Stats.MaxHp = .Stats.MaxHp + AumentoHP

                'Notificamos al user
164             If AumentoHP > 0 Then
                    'Msg197= Has ganado " & AumentoHP & " puntos de vida.", e_FontTypeNames.FONTTYPE_INFO)
166                 Call WriteLocaleMsg(UserIndex, "197", e_FontTypeNames.FONTTYPE_INFO, AumentoHP)

                End If

168             If AumentoSta > 0 Then
                    'Msg198= Has ganado " & AumentoSTA & " puntos de vitalidad.", e_FontTypeNames.FONTTYPE_INFO)
170                 Call WriteLocaleMsg(UserIndex, "198", e_FontTypeNames.FONTTYPE_INFO, AumentoSta)

                End If

172             If AumentoMANA > 0 Then
                    'Msg199= Has ganado " & AumentoMANA & " puntos de magia."
174                 Call WriteLocaleMsg(UserIndex, "199", e_FontTypeNames.FONTTYPE_INFO, AumentoMANA)

                End If

176             If AumentoHIT > 0 Then
178                 Call WriteLocaleMsg(UserIndex, "200", e_FontTypeNames.FONTTYPE_INFO, AumentoHIT)
                    'Msg1292= Tu golpe aumento en ¬1 puntos.
180                 Call WriteLocaleMsg(UserIndex, "1292", e_FontTypeNames.FONTTYPE_INFO, AumentoHIT)

                End If

182             PasoDeNivel = True
184             .Stats.MinHp = .Stats.MaxHp

                ' Call UpdateUserInv(True, UserIndex, 0)
186             If SvrConfig.GetValue("OroPorNivel") > 0 Then
188                 If EsNewbie(UserIndex) Then
                        Dim OroRecompenza As Long
190                     OroRecompenza = SvrConfig.GetValue("OroPorNivel") * .Stats.ELV * SvrConfig.GetValue("GoldMult")
192                     .Stats.GLD = .Stats.GLD + OroRecompenza
                        'Msg1293= Has ganado ¬1 monedas de oro.
194                     Call WriteLocaleMsg(UserIndex, "1293", e_FontTypeNames.FONTTYPE_INFO, OroRecompenza)
196                     Call WriteLocaleMsg(UserIndex, "29", e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(OroRecompenza))

                    End If

                End If

            Loop

198         If PasoDeNivel Then
200             If .Stats.ELV >= STAT_MAXELV Then .Stats.Exp = 0
202             Call UpdateUserInv(True, UserIndex, 0)
                'Call CheckearRecompesas(UserIndex, 3)
204             Call WriteUpdateUserStats(UserIndex)

206             If Pts > 0 Then
208                 .Stats.SkillPts = .Stats.SkillPts + Pts
210                 Call WriteLevelUp(UserIndex, .Stats.SkillPts)
212                 Call WriteLocaleMsg(UserIndex, "187", e_FontTypeNames.FONTTYPE_INFO, Pts)
                    'Msg1294= Has ganado un total de ¬1 skillpoints.
214                 Call WriteLocaleMsg(UserIndex, "1294", e_FontTypeNames.FONTTYPE_INFO, Pts)

                End If

216             If Not EsNewbie(UserIndex) And WasNewbie Then
218                 Call QuitarNewbieObj(UserIndex)
220             ElseIf .Stats.ELV >= MapInfo(.pos.Map).MaxLevel And Not EsGM(UserIndex) Then

222                 If MapInfo(.pos.Map).Salida.Map <> 0 Then
                        ' Msg523=Tu nivel no te permite seguir en el mapa.
224                     Call WriteLocaleMsg(UserIndex, "523", e_FontTypeNames.FONTTYPE_INFO)
226                     Call WarpUserChar(UserIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)

                    End If

                End If

            End If

        End With

        Exit Sub
ErrHandler:
228     Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.Description)

End Sub

Public Sub SwapTargetUserPos(ByVal TargetUser As Integer, _
                             ByRef NewTargetPos As t_WorldPos)
        Dim Heading As e_Heading
100     Heading = UserList(TargetUser).Char.Heading
102     UserList(TargetUser).pos = NewTargetPos
104     Call WritePosUpdate(TargetUser)

106     If UserList(TargetUser).flags.AdminInvisible = 0 Then
108         Call SendData(SendTarget.ToPCAreaButIndex, TargetUser, PrepareMessageCharacterMove(UserList(TargetUser).Char.charIndex, UserList(TargetUser).pos.x, UserList(TargetUser).pos.y), True)
        Else
110         Call SendData(SendTarget.ToAdminAreaButIndex, TargetUser, PrepareMessageCharacterMove(UserList(TargetUser).Char.charIndex, UserList(TargetUser).pos.x, UserList(TargetUser).pos.y))

        End If

112     If IsValidUserRef(UserList(TargetUser).flags.GMMeSigue) Then
114         Call WriteForceCharMoveSiguiendo(UserList(TargetUser).flags.GMMeSigue.ArrayIndex, Heading)

        End If

116     Call WriteForceCharMove(TargetUser, Heading)
        'Update map and char
118     UserList(TargetUser).Char.Heading = Heading
120     MapData(UserList(TargetUser).pos.Map, UserList(TargetUser).pos.x, UserList(TargetUser).pos.y).UserIndex = TargetUser
        'Actualizamos las areas de ser necesario
122     Call ModAreas.CheckUpdateNeededUser(TargetUser, Heading, 0)

End Sub

Function TranslateUserPos(ByVal UserIndex As Integer, _
                          ByRef NewPos As t_WorldPos, _
                          ByVal Speed As Long)
        On Error GoTo TranslateUserPos_Err
        Dim OriginalPos As t_WorldPos

100     With UserList(UserIndex)
102         OriginalPos = .pos

104         If MapInfo(.pos.Map).NumUsers > 1 Or IsValidUserRef(.flags.GMMeSigue) Then
106             If MapData(NewPos.Map, NewPos.x, NewPos.y).UserIndex > 0 Then
108                 Call SwapTargetUserPos(MapData(NewPos.Map, NewPos.x, NewPos.y).UserIndex, .pos)

                End If

            End If

110         If .flags.AdminInvisible = 0 Then
112             If IsValidUserRef(.flags.GMMeSigue) Then
114                 Call SendData(SendTarget.ToPCAreaButFollowerAndIndex, UserIndex, PrepareCharacterTranslate(.Char.charIndex, NewPos.x, NewPos.y, Speed))
116                 Call WriteForceCharMoveSiguiendo(.flags.GMMeSigue.ArrayIndex, .Char.Heading)
                Else
                    'Mando a todos menos a mi donde estoy
118                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCharacterTranslate(.Char.charIndex, NewPos.x, NewPos.y, Speed), True)

                End If

            Else
120             Call SendData(SendTarget.ToAdminAreaButIndex, UserIndex, PrepareCharacterTranslate(.Char.charIndex, NewPos.x, NewPos.y, Speed))

            End If

            'Update map and user pos
122         If MapData(.pos.Map, .pos.x, .pos.y).UserIndex = UserIndex Then
124             MapData(.pos.Map, .pos.x, .pos.y).UserIndex = 0

            End If

126         .pos = NewPos
128         MapData(.pos.Map, .pos.x, .pos.y).UserIndex = UserIndex
130         Call WritePosUpdate(UserIndex)
            'Actualizamos las áreas de ser necesario
132         Call ModAreas.CheckUpdateNeededUser(UserIndex, .Char.Heading, 0)

134         If .Counters.Trabajando Then
136             Call WriteMacroTrabajoToggle(UserIndex, False)

            End If

        End With

        Exit Function
TranslateUserPos_Err:
138     Call LogError("Error en la subrutina TranslateUserPos - Error : " & Err.Number & " - Description : " & Err.Description)

End Function

Public Sub SwapNpcPos(ByVal UserIndex As Integer, _
                      ByRef TargetPos As t_WorldPos, _
                      ByVal nHeading As e_Heading)
        Dim NpcIndex         As Integer
        Dim Opposite_Heading As e_Heading
100     NpcIndex = MapData(TargetPos.Map, TargetPos.x, TargetPos.y).NpcIndex

102     If NpcIndex <= 0 Then Exit Sub
104     Opposite_Heading = InvertHeading(nHeading)
106     Call HeadtoPos(Opposite_Heading, NpcList(NpcIndex).pos)
108     Call SendData(SendTarget.ToNPCAliveArea, NpcIndex, PrepareMessageCharacterMove(NpcList(NpcIndex).Char.charIndex, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y), False)
110     MapData(NpcList(NpcIndex).pos.Map, NpcList(NpcIndex).pos.x, NpcList(NpcIndex).pos.y).NpcIndex = NpcIndex
112     MapData(TargetPos.Map, TargetPos.x, TargetPos.y).NpcIndex = 0
114     Call CheckUpdateNeededNpc(NpcIndex, Opposite_Heading)

End Sub

Function MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As e_Heading) As Boolean
        ' 20/01/2021 - WyroX: Lo convierto a función y saco los WritePosUpdate, ahora están en el paquete
        On Error GoTo MoveUserChar_Err
        Dim nPos             As t_WorldPos
        Dim nPosOriginal     As t_WorldPos
        Dim nPosMuerto       As t_WorldPos
        Dim IndexMover       As Integer
        Dim Opposite_Heading As e_Heading

100     With UserList(UserIndex)
102         nPos = .pos
104         Call HeadtoPos(nHeading, nPos)

106         If Not LegalWalk(.pos.Map, nPos.x, nPos.y, nHeading, .flags.Navegando = 1, .flags.Navegando = 0, .flags.Montado, , UserIndex) Then
                Exit Function

            End If

108         If .flags.Navegando And .invent.BarcoObjIndex = iObjTraje And Not (MapData(.pos.Map, nPos.x, nPos.y).trigger = e_Trigger.DETALLEAGUA Or MapData(.pos.Map, nPos.x, nPos.y).trigger = e_Trigger.NADOCOMBINADO Or MapData(.pos.Map, nPos.x, nPos.y).trigger = e_Trigger.VALIDONADO Or MapData(.pos.Map, nPos.x, nPos.y).trigger = e_Trigger.NADOBAJOTECHO) Then
                Exit Function

            End If

110         If .Accion.AccionPendiente = True Then
112             .Counters.TimerBarra = 0
114             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.charIndex, .Accion.Particula, .Counters.TimerBarra, True))
116             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.charIndex, .Counters.TimerBarra, e_AccionBarra.CancelarAccion))
118             .Accion.AccionPendiente = False
120             .Accion.Particula = 0
122             .Accion.TipoAccion = e_AccionBarra.CancelarAccion
124             .Accion.HechizoPendiente = 0
126             .Accion.RunaObj = 0
128             .Accion.ObjSlot = 0
130             .Accion.AccionPendiente = False

            End If

132         Call SwapNpcPos(UserIndex, nPos, nHeading)

            'Si no estoy solo en el mapa...
134         If MapInfo(.pos.Map).NumUsers > 1 Or IsValidUserRef(.flags.GMMeSigue) Then
                ' Intercambia posición si hay un casper o gm invisible
136             IndexMover = MapData(nPos.Map, nPos.x, nPos.y).UserIndex

138             If IndexMover <> 0 Then
                    ' Sólo puedo patear caspers/gms invisibles si no es él un gm invisible
                    ' If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Function
140                 Call WritePosUpdate(IndexMover)
142                 Opposite_Heading = InvertHeading(nHeading)
144                 Call HeadtoPos(Opposite_Heading, UserList(IndexMover).pos)

                    ' Si es un admin invisible, no se avisa a los demas clientes
146                 If UserList(IndexMover).flags.AdminInvisible = 0 Then
148                     Call SendData(SendTarget.ToPCAreaButIndex, IndexMover, PrepareMessageCharacterMove(UserList(IndexMover).Char.charIndex, UserList(IndexMover).pos.x, UserList(IndexMover).pos.y), True)
                    Else
150                     Call SendData(SendTarget.ToAdminAreaButIndex, IndexMover, PrepareMessageCharacterMove(UserList(IndexMover).Char.charIndex, UserList(IndexMover).pos.x, UserList(IndexMover).pos.y))

                    End If

152                 If IsValidUserRef(UserList(IndexMover).flags.GMMeSigue) Then
154                     Call WriteForceCharMoveSiguiendo(UserList(IndexMover).flags.GMMeSigue.ArrayIndex, Opposite_Heading)

                    End If

156                 Call WriteForceCharMove(IndexMover, Opposite_Heading)
                    'Update map and char
158                 UserList(IndexMover).Char.Heading = Opposite_Heading
160                 MapData(UserList(IndexMover).pos.Map, UserList(IndexMover).pos.x, UserList(IndexMover).pos.y).UserIndex = IndexMover
                    'Actualizamos las areas de ser necesario
162                 Call ModAreas.CheckUpdateNeededUser(IndexMover, Opposite_Heading, 0)

                End If

164             If .flags.AdminInvisible = 0 Then
166                 If IsValidUserRef(.flags.GMMeSigue) Then
168                     Call SendData(SendTarget.ToPCAreaButFollowerAndIndex, UserIndex, PrepareMessageCharacterMove(.Char.charIndex, nPos.x, nPos.y))
170                     Call WriteForceCharMoveSiguiendo(.flags.GMMeSigue.ArrayIndex, nHeading)
                    Else
                        'Mando a todos menos a mi donde estoy
172                     Call SendData(SendTarget.ToPCAliveAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.charIndex, nPos.x, nPos.y), True)
                        Dim LoopC     As Integer
                        Dim tempIndex As Integer
                        'Togle para alternar el paso para los invis
174                     .flags.stepToggle = Not .flags.stepToggle

176                     If Not EsGM(UserIndex) Then
178                         If .flags.invisible + .flags.Oculto > 0 And .flags.Navegando = 0 Then

180                             For LoopC = 1 To ConnGroups(UserList(UserIndex).pos.Map).CountEntrys
182                                 tempIndex = ConnGroups(UserList(UserIndex).pos.Map).UserEntrys(LoopC)

184                                 If tempIndex <> UserIndex And Not EsGM(tempIndex) Then
186                                     If Abs(nPos.x - UserList(tempIndex).pos.x) <= RANGO_VISION_X And Abs(nPos.y - UserList(tempIndex).pos.y) <= RANGO_VISION_Y Then
188                                         If UserList(tempIndex).ConnectionDetails.ConnIDValida Then
190                                             If UserList(tempIndex).flags.Muerto = 0 Or MapInfo(UserList(tempIndex).pos.Map).Seguro = 1 Then
192                                                 If Not CheckGuildSend(UserList(UserIndex), UserList(tempIndex)) Then
194                                                     If .Counters.timeFx + .Counters.timeChat = 0 Then
196                                                         If Distancia(nPos, UserList(tempIndex).pos) > DISTANCIA_ENVIO_DATOS Then
                                                                'Mandamos los pasos para los pjs q estan lejos para que simule que caminen.
                                                                'Mando tambien el char para q lo borre
198                                                             Call WritePlayWaveStep(tempIndex, .Char.charIndex, MapData(nPos.Map, nPos.x, nPos.y).Graphic(1), MapData(nPos.Map, nPos.x, nPos.y).Graphic(2), Distance(nPos.x, nPos.y, UserList(tempIndex).pos.x, UserList(tempIndex).pos.y), Sgn(nPos.x - UserList(tempIndex).pos.x), .flags.stepToggle)
                                                            Else
200                                                             Call WritePosUpdateChar(tempIndex, nPos.x, nPos.y, .Char.charIndex)

                                                            End If

                                                        End If

                                                    End If

                                                End If

                                            End If

                                        End If

                                    End If

202                             Next LoopC

                            End If

                            Dim x As Byte, y As Byte

                            'Esto es para q si me acerco a un usuario que esta invisible y no se mueve me notifique su posicion
204                         For x = nPos.x - DISTANCIA_ENVIO_DATOS To nPos.x + DISTANCIA_ENVIO_DATOS
206                             For y = nPos.y - DISTANCIA_ENVIO_DATOS To nPos.y + DISTANCIA_ENVIO_DATOS
208                                 tempIndex = MapData(.pos.Map, x, y).UserIndex

210                                 If tempIndex > 0 And tempIndex <> UserIndex And Not EsGM(tempIndex) Then
212                                     If UserList(tempIndex).flags.invisible + UserList(tempIndex).flags.Oculto > 0 And UserList(tempIndex).flags.Navegando = 0 And (.GuildIndex = 0 Or .GuildIndex <> UserList(tempIndex).GuildIndex Or modGuilds.NivelDeClan(.GuildIndex) < 6) Then
214                                         Call WritePosUpdateChar(UserIndex, x, y, UserList(tempIndex).Char.charIndex)

                                        End If

                                    End If

216                             Next y
218                         Next x

                        End If

                    End If

                Else
220                 Call SendData(SendTarget.ToAdminAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.charIndex, nPos.x, nPos.y))

                End If

            End If

            'Update map and user pos
222         If MapData(.pos.Map, .pos.x, .pos.y).UserIndex = UserIndex Then
224             MapData(.pos.Map, .pos.x, .pos.y).UserIndex = 0

            End If

226         .pos = nPos
228         .Char.Heading = nHeading
230         MapData(.pos.Map, .pos.x, .pos.y).UserIndex = UserIndex
            'Actualizamos las áreas de ser necesario
232         Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading, 0)

234         If .Counters.Trabajando Then
236             Call WriteMacroTrabajoToggle(UserIndex, False)

            End If

238         If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1

        End With

240     MoveUserChar = True
        Exit Function
MoveUserChar_Err:
242     Call TraceError(Err.Number, Err.Description + " UI:" + UserIndex, "UsUaRiOs.MoveUserChar", Erl)

End Function

Public Function InvertHeading(ByVal nHeading As e_Heading) As e_Heading
        On Error GoTo InvertHeading_Err

        'Returns the heading opposite to the one passed by val.
100     Select Case nHeading

            Case e_Heading.EAST
102             InvertHeading = e_Heading.WEST

104         Case e_Heading.WEST
106             InvertHeading = e_Heading.EAST

108         Case e_Heading.SOUTH
110             InvertHeading = e_Heading.NORTH

112         Case e_Heading.NORTH
114             InvertHeading = e_Heading.SOUTH

        End Select

        Exit Function
InvertHeading_Err:
116     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.InvertHeading", Erl)

End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, _
                  ByVal Slot As Byte, _
                  ByRef Object As t_UserOBJ)
        On Error GoTo ChangeUserInv_Err
100     UserList(UserIndex).invent.Object(Slot) = Object
102     Call WriteChangeInventorySlot(UserIndex, Slot)
        Exit Sub
ChangeUserInv_Err:
104     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ChangeUserInv", Erl)

End Sub

Function NextOpenCharIndex() As Integer
        On Error GoTo NextOpenCharIndex_Err
        Dim LoopC As Long

100     For LoopC = 1 To MAXCHARS

102         If CharList(LoopC) = 0 Then
104             NextOpenCharIndex = LoopC
106             NumChars = NumChars + 1

108             If LoopC > LastChar Then LastChar = LoopC
                Exit Function

            End If

110     Next LoopC

        Exit Function
NextOpenCharIndex_Err:
112     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.NextOpenCharIndex", Erl)

End Function

Function NextOpenUser() As Integer
        On Error GoTo NextOpenUser_Err
        Dim LoopC As Long

100     If IsFeatureEnabled("use_old_user_slot_check") Then

102         For LoopC = 1 To MaxUsers + 1

104             If LoopC > MaxUsers Then Exit For
106             If (Not UserList(LoopC).ConnectionDetails.ConnIDValida And UserList(LoopC).flags.UserLogged = False) Then Exit For
108         Next LoopC
110         NextOpenUser = LoopC
        Else
112         NextOpenUser = GetNextAvailableUserSlot

        End If

        Exit Function
NextOpenUser_Err:
114     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.NextOpenUser", Erl)

End Function

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        On Error GoTo SendUserStatsTxt_Err
        Dim GuildI As Integer
        'Msg1295= Estadisticas de: ¬1
100     Call WriteLocaleMsg(sendIndex, "1295", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).name)
102     Call WriteConsoleMsg(sendIndex, "Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & ExpLevelUp(UserList(UserIndex).Stats.ELV), e_FontTypeNames.FONTTYPE_INFO)
104     Call WriteConsoleMsg(sendIndex, "Salud: " & UserList(UserIndex).Stats.MinHp & "/" & UserList(UserIndex).Stats.MaxHp & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta, e_FontTypeNames.FONTTYPE_INFO)

106     If UserList(UserIndex).invent.WeaponEqpObjIndex > 0 Then
108         Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHit & " (" & ObjData(UserList(UserIndex).invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).invent.WeaponEqpObjIndex).MaxHit & ")", e_FontTypeNames.FONTTYPE_INFO)
        Else
110         Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHit, e_FontTypeNames.FONTTYPE_INFO)

        End If

112     If UserList(UserIndex).invent.ArmourEqpObjIndex > 0 Then
114         If UserList(UserIndex).invent.EscudoEqpObjIndex > 0 Then
116             Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).invent.ArmourEqpObjIndex).MinDef + ObjData(UserList(UserIndex).invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).invent.ArmourEqpObjIndex).MaxDef + ObjData(UserList(UserIndex).invent.EscudoEqpObjIndex).MaxDef, e_FontTypeNames.FONTTYPE_INFO)
            Else
118             Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).invent.ArmourEqpObjIndex).MaxDef, e_FontTypeNames.FONTTYPE_INFO)

            End If

        Else
            'Msg1098= (CUERPO) Min Def/Max Def: 0
120         Call WriteLocaleMsg(sendIndex, "1098", e_FontTypeNames.FONTTYPE_INFO)

        End If

122     If UserList(UserIndex).invent.CascoEqpObjIndex > 0 Then
124         Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).invent.CascoEqpObjIndex).MaxDef, e_FontTypeNames.FONTTYPE_INFO)
        Else
            'Msg1099= (CABEZA) Min Def/Max Def: 0
126         Call WriteLocaleMsg(sendIndex, "1099", e_FontTypeNames.FONTTYPE_INFO)

        End If

128     GuildI = UserList(UserIndex).GuildIndex

130     If GuildI > 0 Then
            'Msg1296= Clan: ¬1
132         Call WriteLocaleMsg(sendIndex, "1296", e_FontTypeNames.FONTTYPE_INFO, modGuilds.GuildName(GuildI))

134         If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(UserList(sendIndex).name) Then
                'Msg1100= Status: Líder
136             Call WriteLocaleMsg(sendIndex, "1100", e_FontTypeNames.FONTTYPE_INFO)

            End If

            'guildpts no tienen objeto
        End If

        #If ConUpTime Then
            Dim TempDate As Date
            Dim TempSecs As Long
            Dim TempStr  As String
138         TempDate = Now - UserList(UserIndex).LogOnTime
140         TempSecs = (UserList(UserIndex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
142         TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
144         Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), e_FontTypeNames.FONTTYPE_INFO)
            'Msg1297= Total: ¬1
146         Call WriteLocaleMsg(sendIndex, "1297", e_FontTypeNames.FONTTYPE_INFO, TempStr)
        #End If
148     Call LoadPatronCreditsFromDB(UserIndex)
        'Msg1298= Oro: ¬1
150     Call WriteLocaleMsg(sendIndex, "1298", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).Stats.GLD)
152     Call WriteConsoleMsg(sendIndex, "Dados: " & UserList(UserIndex).Stats.UserAtributos(e_Atributos.Fuerza) & ", " & UserList(UserIndex).Stats.UserAtributos(e_Atributos.Agilidad) & ", " & UserList(UserIndex).Stats.UserAtributos(e_Atributos.Inteligencia) & ", " & UserList(UserIndex).Stats.UserAtributos(e_Atributos.Constitucion) & ", " & UserList(UserIndex).Stats.UserAtributos(e_Atributos.Carisma), e_FontTypeNames.FONTTYPE_INFO)
        'Msg1299= Veces que Moriste: ¬1
154     Call WriteLocaleMsg(sendIndex, "1299", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).flags.VecesQueMoriste)
156     Call WriteLocaleMsg(sendIndex, MsgFactionScore, e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).Faccion.FactionScore)
        'Msg1300= Creditos Patreon: ¬1
158     Call WriteLocaleMsg(sendIndex, "1300", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).Stats.Creditos)
        Exit Sub
SendUserStatsTxt_Err:
160     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserStatsTxt", Erl)

End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        On Error GoTo SendUserMiniStatsTxt_Err

        'Shows the users Stats when the user is online.
        ' Agrego de funciones y mejora de distribución de parámetros.
100     With UserList(UserIndex)
            'Msg1301= Pj: ¬1
102         Call WriteLocaleMsg(sendIndex, "1301", e_FontTypeNames.FONTTYPE_INFO, .name)
            'Msg1302= Ciudadanos Matados: ¬1
104         Call WriteLocaleMsg(sendIndex, "1302", e_FontTypeNames.FONTTYPE_INFO, .Faccion.ciudadanosMatados)
            'Msg1303= Criminales Matados: ¬1
106         Call WriteLocaleMsg(sendIndex, "1303", e_FontTypeNames.FONTTYPE_INFO, .Faccion.CriminalesMatados)
            'Msg1304= UsuariosMatados: ¬1
108         Call WriteLocaleMsg(sendIndex, "1304", e_FontTypeNames.FONTTYPE_INFO, .Stats.UsuariosMatados)
            'Msg1305= NPCsMuertos: ¬1
110         Call WriteLocaleMsg(sendIndex, "1305", e_FontTypeNames.FONTTYPE_INFO, .Stats.NPCsMuertos)
            'Msg1306= Clase: ¬1
112         Call WriteLocaleMsg(sendIndex, "1306", e_FontTypeNames.FONTTYPE_INFO, ListaClases(.clase))
            'Msg1307= Pena: ¬1
114         Call WriteLocaleMsg(sendIndex, "1307", e_FontTypeNames.FONTTYPE_INFO, .Counters.Pena)

116         If .GuildIndex > 0 Then
                'Msg1308= Clan: ¬1
118             Call WriteLocaleMsg(sendIndex, "1308", e_FontTypeNames.FONTTYPE_INFO, GuildName(.GuildIndex))

            End If

            'Msg1309= Oro en billetera: ¬1
120         Call WriteLocaleMsg(sendIndex, "1309", e_FontTypeNames.FONTTYPE_INFO, .Stats.GLD)
            'Msg1310= Oro en banco: ¬1
122         Call WriteLocaleMsg(sendIndex, "1310", e_FontTypeNames.FONTTYPE_INFO, .Stats.Banco)

        End With

        Exit Sub
SendUserMiniStatsTxt_Err:
124     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserMiniStatsTxt", Erl)

End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        On Error GoTo SendUserInvTxt_Err
        Dim j As Long
100     Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)
        'Msg1311= Tiene ¬1 objetos.
102     Call WriteLocaleMsg(sendIndex, "1311", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).invent.NroItems)

104     For j = 1 To UserList(UserIndex).CurrentInventorySlots

106         If UserList(UserIndex).invent.Object(j).ObjIndex > 0 Then
108             Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(UserList(UserIndex).invent.Object(j).ObjIndex).name & " Cantidad:" & UserList(UserIndex).invent.Object(j).amount, e_FontTypeNames.FONTTYPE_INFO)

            End If

110     Next j

        Exit Sub
SendUserInvTxt_Err:
112     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserInvTxt", Erl)

End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        On Error GoTo SendUserSkillsTxt_Err
        Dim j As Integer
100     Call WriteConsoleMsg(sendIndex, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)

102     For j = 1 To NUMSKILLS
104         Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), e_FontTypeNames.FONTTYPE_INFO)
        Next
        'Msg1312=  SkillLibres:¬1
106     Call WriteLocaleMsg(sendIndex, "1312", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).Stats.SkillPts)
        Exit Sub
SendUserSkillsTxt_Err:
108     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserSkillsTxt", Erl)

End Sub

Function DameUserIndexConNombre(ByVal nombre As String) As Integer
        On Error GoTo DameUserIndexConNombre_Err
        Dim LoopC As Integer
100     LoopC = 1
102     nombre = UCase$(nombre)

104     Do Until UCase$(UserList(LoopC).name) = nombre
106         LoopC = LoopC + 1

108         If LoopC > MaxUsers Then
110             DameUserIndexConNombre = 0
                Exit Function

            End If

        Loop
112     DameUserIndexConNombre = LoopC
        Exit Function
DameUserIndexConNombre_Err:
114     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.DameUserIndexConNombre", Erl)

End Function

Sub NPCAtacado(ByVal NpcIndex As Integer, _
               ByVal UserIndex As Integer, _
               Optional ByVal AffectsOwner As Boolean = True)
        On Error GoTo NPCAtacado_Err
        ' WyroX: El usuario pierde la protección
100     UserList(UserIndex).Counters.TiempoDeInmunidad = 0
102     UserList(UserIndex).flags.Inmunidad = 0

        'Guardamos el usuario que ataco el npc.
104     If Not IsSet(NpcList(NpcIndex).flags.StatusMask, eTaunted) And NpcList(NpcIndex).Movement <> Estatico And NpcList(NpcIndex).flags.AttackedFirstBy = vbNullString Then
106         Call SetUserRef(NpcList(NpcIndex).TargetUser, UserIndex)
108         NpcList(NpcIndex).Hostile = 1

110         If AffectsOwner Then
112             NpcList(NpcIndex).flags.AttackedBy = UserList(UserIndex).name
114             NpcList(NpcIndex).flags.AttackedTime = GlobalFrameTime

            End If

        End If

        'Guarda el NPC que estas atacando ahora.
116     If AffectsOwner Then Call SetNpcRef(UserList(UserIndex).flags.NPCAtacado, NpcIndex)
118     If NpcList(NpcIndex).flags.Faccion = Armada And Status(UserIndex) = e_Facciones.Ciudadano Then
120         Call VolverCriminal(UserIndex)

        End If

122     If IsValidUserRef(NpcList(NpcIndex).MaestroUser) And NpcList(NpcIndex).MaestroUser.ArrayIndex <> UserIndex Then
124         Call AllMascotasAtacanUser(UserIndex, NpcList(NpcIndex).MaestroUser.ArrayIndex)

        End If

        Exit Sub
NPCAtacado_Err:
126     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.NPCAtacado", Erl)

End Sub

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)
        On Error GoTo SubirSkill_Err
        Dim Lvl As Integer, maxPermitido As Integer
100     Lvl = UserList(UserIndex).Stats.ELV

102     If UserList(UserIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub

        ' Se suben 5 skills cada dos niveles como máximo.
104     If (Lvl Mod 2 = 0) Then ' El level es numero par
106         maxPermitido = (Lvl \ 2) * 5
        Else ' El level es numero impar
            ' Esta cuenta signifca, que si el nivel anterior terminaba en 5 ahora
            ' suma dos puntos mas, sino 3. Lo de siempre.
108         maxPermitido = (Lvl \ 2) * 5 + 3 - (((((Lvl - 1) \ 2) * 5) Mod 10) \ 5)

        End If

110     If UserList(UserIndex).Stats.UserSkills(Skill) >= maxPermitido Then Exit Sub
112     If UserList(UserIndex).Stats.MinHam > 0 And UserList(UserIndex).Stats.MinAGU > 0 Then
            Dim Aumenta As Integer
            Dim Prob    As Integer
            Dim Menor   As Byte
114         Menor = 10

116         Select Case Lvl

                Case Is <= 12
118                 Prob = 15

120             Case Is <= 24
122                 Prob = 30

124             Case Else
126                 Prob = 50

            End Select

128         Aumenta = RandomNumber(1, Prob * DificultadSubirSkill)

130         If UserList(UserIndex).flags.PendienteDelExperto = 1 Then
132             Menor = 15

            End If

134         If Aumenta < Menor Then
136             UserList(UserIndex).Stats.UserSkills(Skill) = UserList(UserIndex).Stats.UserSkills(Skill) + 1
138             Call WriteConsoleMsg(UserIndex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts.", e_FontTypeNames.FONTTYPE_INFO)
                Dim BonusExp As Long
140             BonusExp = 5& * SvrConfig.GetValue("ExpMult")
                'Msg1313= ¡Has ganado ¬1 puntos de experiencia!
142             Call WriteLocaleMsg(UserIndex, "1313", e_FontTypeNames.FONTTYPE_INFOIAO, BonusExp)

144             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
146                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + BonusExp

148                 If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
150                 UserList(UserIndex).flags.ModificoSkills = True

152                 If UserList(UserIndex).ChatCombate = 1 Then
154                     Call WriteLocaleMsg(UserIndex, "140", e_FontTypeNames.FONTTYPE_EXP, BonusExp)

                    End If

156                 Call WriteUpdateExp(UserIndex)
158                 Call CheckUserLevel(UserIndex)

                End If

            End If

        End If

        Exit Sub
SubirSkill_Err:
160     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SubirSkill", Erl)

End Sub

Public Sub SubirSkillDeArmaActual(ByVal UserIndex As Integer)
        On Error GoTo SubirSkillDeArmaActual_Err

100     With UserList(UserIndex)

102         If .invent.WeaponEqpObjIndex > 0 Then

                ' Arma con proyectiles, subimos armas a distancia
104             If ObjData(.invent.WeaponEqpObjIndex).Proyectil Then
106                 Call SubirSkill(UserIndex, e_Skill.Proyectiles)
108             ElseIf ObjData(.invent.WeaponEqpObjIndex).WeaponType = eKnuckle Then
110                 Call SubirSkill(UserIndex, e_Skill.Wrestling)
                    ' Sino, subimos combate con armas
                Else
112                 Call SubirSkill(UserIndex, e_Skill.Armas)

                End If

                ' Si no está usando un arma, subimos combate sin armas
            Else
114             Call SubirSkill(UserIndex, e_Skill.Wrestling)

            End If

        End With

        Exit Sub
SubirSkillDeArmaActual_Err:
116     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SubirSkillDeArmaActual", Erl)

End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'
Sub UserDie(ByVal UserIndex As Integer)
        'Ahora se resetea el counter del invi
        On Error GoTo ErrorHandler
        Dim i  As Long
        Dim aN As Integer

100     With UserList(UserIndex)
102         .Counters.Mimetismo = 0
104         .flags.Mimetizado = e_EstadoMimetismo.Desactivado
106         Call RefreshCharStatus(UserIndex)
            'Sonido
108         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(IIf(.genero = e_Genero.Hombre, e_SoundIndex.MUERTE_HOMBRE, e_SoundIndex.MUERTE_MUJER), .pos.x, .pos.y))
            'Quitar el dialogo del user muerto
110         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.charIndex))
112         .Stats.MinHp = 0
114         .Stats.MinSta = 0
116         .Stats.shield = 0
118         .flags.AtacadoPorUser = 0
120         .flags.incinera = 0
122         .flags.Paraliza = 0
124         .flags.Envenena = 0
126         .flags.Estupidiza = 0
128         Call ClearEffectList(.EffectOverTime, e_EffectType.eAny, True)
130         Call ClearModifiers(.Modifiers)
132         .flags.Muerto = 1
134         Call WriteUpdateHP(UserIndex)
136         Call WriteUpdateSta(UserIndex)
138         Call ClearAttackerNpc(UserIndex)

140         If MapData(.pos.Map, .pos.x, .pos.y).trigger <> e_Trigger.ZONAPELEA And MapInfo(.pos.Map).DropItems Then
142             If (.flags.Privilegios And e_PlayerType.User) <> 0 Then
144                 If .flags.PendienteDelSacrificio = 0 Then
146                     Call TirarTodosLosItems(UserIndex)
                    Else
                        Dim MiObj As t_Obj
148                     MiObj.amount = 1
150                     MiObj.ObjIndex = PENDIENTE
152                     Call QuitarObjetos(PENDIENTE, 1, UserIndex)

                    End If

                End If

            End If

154         Call Desequipar(UserIndex, .invent.ArmourEqpSlot)
156         Call Desequipar(UserIndex, .invent.WeaponEqpSlot)
158         Call Desequipar(UserIndex, .invent.EscudoEqpSlot)
160         Call Desequipar(UserIndex, .invent.CascoEqpSlot)
162         Call Desequipar(UserIndex, .invent.DañoMagicoEqpSlot)
164         Call Desequipar(UserIndex, .invent.HerramientaEqpSlot)
166         Call Desequipar(UserIndex, .invent.MonturaSlot)
168         Call Desequipar(UserIndex, .invent.MunicionEqpSlot)
170         Call Desequipar(UserIndex, .invent.MagicoSlot)
172         Call Desequipar(UserIndex, .invent.ResistenciaEqpSlot)

            'desequipar montura
174         If .flags.Montado > 0 Then
176             Call DoMontar(UserIndex, ObjData(.invent.MonturaObjIndex), .invent.MonturaSlot)

            End If

            ' << Reseteamos los posibles FX sobre el personaje >>
178         If .Char.loops = INFINITE_LOOPS Then
180             .Char.FX = 0
182             .Char.loops = 0

            End If

184         If TriggerZonaPelea(UserIndex, UserIndex) <> TRIGGER6_PERMITE Then
186             .flags.VecesQueMoriste = .flags.VecesQueMoriste + 1

            End If

            ' << Restauramos los atributos >>
188         If .flags.TomoPocion Then

190             For i = 1 To 4
192                 .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
194             Next i

196             Call WriteFYA(UserIndex)

            End If

            ' << Frenamos el contador de la droga >>
198         .flags.DuracionEfecto = 0

            '<< Cambiamos la apariencia del char >>
200         If .flags.Navegando = 0 Then
202             .Char.body = iCuerpoMuerto
204             .Char.head = 0
206             .Char.ShieldAnim = NingunEscudo
208             .Char.WeaponAnim = NingunArma
210             .Char.CascoAnim = NingunCasco
212             .Char.CartAnim = NoCart
            Else
214             Call EquiparBarco(UserIndex)

            End If

216         Call ActualizarVelocidadDeUsuario(UserIndex)
218         Call LimpiarEstadosAlterados(UserIndex)

220         For i = 1 To MAXMASCOTAS

222             If .MascotasIndex(i).ArrayIndex > 0 Then
224                 If IsValidNpcRef(.MascotasIndex(i)) Then
226                     Call MuereNpc(.MascotasIndex(i).ArrayIndex, 0)
                    Else
228                     Call ClearNpcRef(.MascotasIndex(i))

                    End If

                End If

230         Next i

232         If .clase = e_Class.Druid Then
                Dim Params() As Variant
                Dim ParamC   As Long
234             ReDim Params(MAXMASCOTAS * 3 - 1)
236             ParamC = 0

238             For i = 1 To MAXMASCOTAS
240                 Params(ParamC) = .Id
242                 ParamC = ParamC + 1
244                 Params(ParamC) = i
246                 ParamC = ParamC + 1
248                 Params(ParamC) = 0
250                 ParamC = ParamC + 1
252             Next i

254             Call Execute(QUERY_UPSERT_PETS, Params)

            End If

256         If (.flags.MascotasGuardadas = 0) Then
258             .NroMascotas = 0

            End If

            '<< Actualizamos clientes >>
260         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NoCart)

262         If MapInfo(.pos.Map).Seguro = 0 Then
                ' Msg524=Escribe /HOGAR si deseas regresar rápido a tu hogar.
264             Call WriteLocaleMsg(UserIndex, "524", e_FontTypeNames.FONTTYPE_New_Naranja)

            End If

266         If .flags.EnReto Then
268             Call MuereEnReto(UserIndex)

            End If

270         If .flags.jugando_captura = 1 Then
272             If Not InstanciaCaptura Is Nothing Then
274                 Call InstanciaCaptura.muereUsuario(UserIndex)

                End If

            End If

            'Borramos todos los personajes del area
            'HarThaoS: Mando un 5 en head para que cuente como muerto el area y no recalcule las posiciones.
276         Call CheckUpdateNeededUser(UserIndex, 5, 0, .flags.Muerto)
            Dim LoopC     As Long
            Dim tempIndex As Integer
            Dim Map       As Integer
            Dim AreaX     As Integer
            Dim AreaY     As Integer
278         AreaX = UserList(UserIndex).AreasInfo.AreaPerteneceX
280         AreaY = UserList(UserIndex).AreasInfo.AreaPerteneceY

282         For LoopC = 1 To ConnGroups(UserList(UserIndex).pos.Map).CountEntrys
284             tempIndex = ConnGroups(UserList(UserIndex).pos.Map).UserEntrys(LoopC)

286             If UserList(tempIndex).AreasInfo.AreaReciveX And AreaX Then  'Esta en el area?
288                 If UserList(tempIndex).AreasInfo.AreaReciveY And AreaY Then
290                     If UserList(tempIndex).ConnectionDetails.ConnIDValida Then

                            'Si no soy el que se murió
292                         If UserIndex <> tempIndex And (Not EsGM(UserIndex)) And MapInfo(UserList(UserIndex).pos.Map).Seguro = 0 And UserList(tempIndex).flags.AdminInvisible = 1 Then
294                             If UserList(UserIndex).GuildIndex = 0 Then
296                                 Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageCharacterRemove(3, UserList(tempIndex).Char.charIndex, True))
                                Else

298                                 If UserList(UserIndex).GuildIndex <> UserList(tempIndex).GuildIndex Then
300                                     Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageCharacterRemove(3, UserList(tempIndex).Char.charIndex, True))

                                    End If

                                End If

                            End If

                        End If

                    End If

                End If

302         Next LoopC

        End With

        Exit Sub
ErrorHandler:
304     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.UserDie", Erl)

End Sub

Public Function AlreadyKilledBy(ByVal TargetIndex As Integer, _
                                ByVal killerIndex As Integer) As Boolean
        Dim TargetPos As Integer

100     With UserList(TargetIndex)
102         TargetPos = Min(.flags.LastKillerIndex, MaxRecentKillToStore)
            Dim i As Integer

104         For i = 0 To TargetPos

106             If .flags.RecentKillers(i).UserId = UserList(killerIndex).Id And (GlobalFrameTime - .flags.RecentKillers(i).KillTime) < FactionReKillTime Then
108                 AlreadyKilledBy = True
                    Exit Function

                End If

110         Next i

        End With

End Function

Public Sub RegisterRecentKiller(ByVal TargetIndex As Integer, _
                                ByVal killerIndex As Integer)
        Dim InsertIndex As Integer

100     With UserList(TargetIndex)
102         InsertIndex = .flags.LastKillerIndex Mod MaxRecentKillToStore
104         .flags.RecentKillers(InsertIndex).UserId = UserList(killerIndex).Id
106         .flags.RecentKillers(InsertIndex).KillTime = GlobalFrameTime
108         .flags.LastKillerIndex = .flags.LastKillerIndex + 1

110         If .flags.LastKillerIndex > MaxRecentKillToStore * 10 Then 'prevent overflow
112             .flags.LastKillerIndex = .flags.LastKillerIndex \ 10

            End If

        End With

End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
        On Error GoTo ContarMuerte_Err

100     If EsNewbie(Muerto) Then Exit Sub
102     If PeleaSegura(Atacante, Muerto) Then Exit Sub

        'Si se llevan más de 10 niveles no le cuento la muerte.
104     If CInt(UserList(Atacante).Stats.ELV) - CInt(UserList(Muerto).Stats.ELV) > 10 Then Exit Sub
        Dim AttackerStatus As e_Facciones
106     AttackerStatus = Status(Atacante)

108     If Status(Muerto) = e_Facciones.Criminal Or Status(Muerto) = e_Facciones.Caos Or Status(Muerto) = e_Facciones.concilio Then
110         If Not AlreadyKilledBy(Muerto, Atacante) Then
112             Call RegisterRecentKiller(Muerto, Atacante)

114             If UserList(Atacante).Faccion.CriminalesMatados < MAXUSERMATADOS Then
116                 UserList(Atacante).Faccion.CriminalesMatados = UserList(Atacante).Faccion.CriminalesMatados + 1

                End If

118             If AttackerStatus = e_Facciones.Ciudadano Or AttackerStatus = e_Facciones.Armada Or AttackerStatus = e_Facciones.consejo Then
120                 Call HandleFactionScoreForKill(Atacante, Muerto)

                End If

            End If

122     ElseIf Status(Muerto) = e_Facciones.Ciudadano Or Status(Muerto) = e_Facciones.Armada Or Status(Muerto) = e_Facciones.consejo Then

124         If Not AlreadyKilledBy(Muerto, Atacante) Then
126             Call RegisterRecentKiller(Muerto, Atacante)

128             If UserList(Atacante).Faccion.ciudadanosMatados < MAXUSERMATADOS Then
130                 UserList(Atacante).Faccion.ciudadanosMatados = UserList(Atacante).Faccion.ciudadanosMatados + 1

                End If

132             If AttackerStatus = e_Facciones.Criminal Or AttackerStatus = e_Facciones.Caos Or AttackerStatus = e_Facciones.concilio Then
134                 Call HandleFactionScoreForKill(Atacante, Muerto)

                End If

            End If

        End If

        Exit Sub
ContarMuerte_Err:
136     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ContarMuerte", Erl)

End Sub

Sub HandleFactionScoreForKill(ByVal UserIndex As Integer, ByVal TargetIndex As Integer)
        Dim Score As Integer

100     With UserList(UserIndex)
102         Score = 10 - max(CInt(.Stats.ELV) - CInt(UserList(TargetIndex).Stats.ELV), 0)

104         If GlobalFrameTime - .flags.LastHelpByTime < AssistHelpValidTime Then
106             If IsValidUserRef(.flags.LastHelpUser) And .flags.LastHelpUser.ArrayIndex <> UserIndex Then
108                 Score = Score - 1
110                 Call HandleFactionScoreForAssist(.flags.LastHelpUser.ArrayIndex, TargetIndex)

                End If

            End If

112         If GlobalFrameTime - UserList(TargetIndex).flags.LastAttackedByUserTime < AssistDamageValidTime Then
114             If IsValidUserRef(UserList(TargetIndex).flags.LastAttacker) And UserList(TargetIndex).flags.LastAttacker.ArrayIndex <> UserIndex Then
116                 Score = Score - 1
118                 Call HandleFactionScoreForAssist(UserList(TargetIndex).flags.LastAttacker.ArrayIndex, TargetIndex)

                End If

            End If

120         .Faccion.FactionScore = .Faccion.FactionScore + max(Score, 0)

        End With

End Sub

Sub HandleFactionScoreForAssist(ByVal UserIndex As Integer, ByVal TargetIndex As Integer)
        Dim Score As Integer

100     With UserList(UserIndex)
102         Score = 10 - max(CInt(.Stats.ELV) - CInt(UserList(TargetIndex).Stats.ELV), 0)
104         Score = Score / 2
106         .Faccion.FactionScore = .Faccion.FactionScore + max(Score, 0)

        End With

End Sub

Sub Tilelibre(ByRef pos As t_WorldPos, _
              ByRef nPos As t_WorldPos, _
              ByRef obj As t_Obj, _
              ByRef Agua As Boolean, _
              ByRef Tierra As Boolean, _
              Optional ByVal InitialPos As Boolean = True)
        On Error GoTo Tilelibre_Err
        'El agua es ahora un TileLibre agregando las condiciones necesarias.
        Dim Notfound As Boolean
        Dim LoopC    As Integer
        Dim tX       As Integer
        Dim tY       As Integer
        Dim hayobj   As Boolean
100     hayobj = False
102     nPos.Map = pos.Map

104     Do While Not LegalPos(pos.Map, nPos.x, nPos.y, Agua, Tierra) Or hayobj

106         If LoopC > 15 Then
108             Notfound = True
                Exit Do

            End If

110         For tY = pos.y - LoopC To pos.y + LoopC
112             For tX = pos.x - LoopC To pos.x + LoopC

114                 If LegalPos(nPos.Map, tX, tY, Agua, Tierra) Then
                        'We continue if: a - the item is different from 0 and the dropped item or b - the Amount dropped + Amount in map exceeds MAX_INVENTORY_OBJS
116                     hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex <> obj.ObjIndex)

118                     If Not hayobj Then hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.amount + obj.amount > MAX_INVENTORY_OBJS)
120                     If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 And (InitialPos Or (tX <> pos.x And tY <> pos.y)) Then
122                         nPos.x = tX
124                         nPos.y = tY
126                         tX = pos.x + LoopC
128                         tY = pos.y + LoopC

                        End If

                    End If

130             Next tX
132         Next tY

134         LoopC = LoopC + 1
        Loop

136     If Notfound = True Then
138         nPos.x = 0
140         nPos.y = 0

        End If

        Exit Sub
Tilelibre_Err:
142     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.Tilelibre", Erl)

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

100     Do While True

102         If LoopC > 20 Then Exit Sub

104         For tY = y - LoopC To y + LoopC
106             For tX = x - LoopC To x + LoopC

108                 If LegalPos(Map, tX, tY, AguaValida, True, UserList(UserIndex).flags.Montado = 1, False, False) Then
110                     If MapData(Map, tX, tY).trigger < 50 Then
112                         Call WarpUserChar(UserIndex, Map, tX, tY, FX)
                            Exit Sub

                        End If

                    End If

114             Next tX
116         Next tY

118         LoopC = LoopC + 1
        Loop
120     Call WarpUserChar(UserIndex, Map, x, y, FX)
        Exit Sub
WarpToLegalPos_Err:
122     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.WarpToLegalPos", Erl)

End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, _
                 ByVal Map As Integer, _
                 ByVal x As Integer, _
                 ByVal y As Integer, _
                 Optional ByVal FX As Boolean = False)
        On Error GoTo WarpUserChar_Err
        Dim OldMap As Integer
        Dim OldX   As Integer
        Dim OldY   As Integer

100     With UserList(UserIndex)

102         If Map <= 0 Then Exit Sub
104         If IsValidUserRef(.ComUsu.DestUsu) Then
106             If UserList(.ComUsu.DestUsu.ArrayIndex).flags.UserLogged Then
108                 If UserList(.ComUsu.DestUsu.ArrayIndex).ComUsu.DestUsu.ArrayIndex = UserIndex Then
                        'Msg1101= Comercio cancelado por el otro usuario
110                     Call WriteLocaleMsg(.ComUsu.DestUsu.ArrayIndex, "1101", e_FontTypeNames.FONTTYPE_TALK)
112                     Call FinComerciarUsu(.ComUsu.DestUsu.ArrayIndex)

                    End If

                End If

            End If

            'Quitar el dialogo
114         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.charIndex))
116         Call WriteRemoveAllDialogs(UserIndex)
118         OldMap = .pos.Map
120         OldX = .pos.x
122         OldY = .pos.y
124         Call EraseUserChar(UserIndex, True, FX)

126         If OldMap <> Map Then
128             Call WriteChangeMap(UserIndex, Map)

130             If MapInfo(OldMap).Seguro = 1 And MapInfo(Map).Seguro = 0 And .Stats.ELV < 42 Then
                    ' Msg573=Estás saliendo de una zona segura, recuerda que aquí corres riesgo de ser atacado.
132                 Call WriteLocaleMsg(UserIndex, "573", e_FontTypeNames.FONTTYPE_WARNING)

                End If

                'Update new Map Users
134             MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
                'Update old Map Users
136             MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1

138             If MapInfo(OldMap).NumUsers < 0 Then
140                 MapInfo(OldMap).NumUsers = 0

                End If

142             If .flags.Traveling = 1 Then
144                 .flags.Traveling = 0
146                 .Counters.goHome = 0
                    ' Msg574=El viaje ha terminado.
148                 Call WriteLocaleMsg(UserIndex, "574", e_FontTypeNames.FONTTYPE_INFOBOLD)

                End If

            End If

150         .pos.x = x
152         .pos.y = y
154         .pos.Map = Map

156         If .Grupo.EnGrupo = True Then
158             Call CompartirUbicacion(UserIndex)

            End If

160         If FX Then
162             Call MakeUserChar(True, Map, UserIndex, Map, x, y, 1)
            Else
164             Call MakeUserChar(True, Map, UserIndex, Map, x, y, 0)

            End If

166         Call WriteUserCharIndexInServer(UserIndex)

168         If IsValidUserRef(.flags.GMMeSigue) Then
170             Call WriteSendFollowingCharindex(.flags.GMMeSigue.ArrayIndex, .Char.charIndex)

            End If

            'Seguis invisible al pasar de mapa
172         If (.flags.invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then

                ' Si el mapa lo permite
174             If MapInfo(Map).SinInviOcul Then
176                 .flags.invisible = 0
178                 .flags.Oculto = 0
180                 .Counters.TiempoOculto = 0
182                 .Counters.Invisibilidad = 0
184                 .Counters.DisabledInvisibility = 0
186                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.charIndex, False))
                    ' Msg575=Una fuerza divina que vigila esta zona te ha vuelto visible.
188                 Call WriteLocaleMsg(UserIndex, "575", e_FontTypeNames.FONTTYPE_INFO)
                Else
190                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charIndex, True))

                End If

            End If

            'Reparacion temporal del bug de particulas. 08/07/09 LADDER
192         If .flags.AdminInvisible = 0 Then
194             If FX Then 'FX
196                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_WARP, x, y))
198                 UserList(UserIndex).Counters.timeFx = 3
200                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charIndex, e_FXIDs.FXWARP, 0, .pos.x, .pos.y))

                End If

            Else
202             Call SendData(ToIndex, UserIndex, PrepareMessageSetInvisible(.Char.charIndex, True))

            End If

204         If .NroMascotas > 0 Then Call WarpMascotas(UserIndex)
206         If MapInfo(Map).zone = "DUNGEON" Or MapData(Map, x, y).trigger >= 9 Then
208             If .flags.Montado > 0 Then
210                 Call DoMontar(UserIndex, ObjData(.invent.MonturaObjIndex), .invent.MonturaSlot)

                End If

            End If

        End With

        Exit Sub
WarpUserChar_Err:
212     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.WarpUserChar", Erl)

End Sub

Sub Cerrar_Usuario(ByVal UserIndex As Integer, _
                   Optional ByVal forceClose As Boolean = False)
        On Error GoTo Cerrar_Usuario_Err

100     With UserList(UserIndex)

102         If IsFeatureEnabled("debug_connections") Then
104             Call AddLogToCircularBuffer("Cerrar_Usuario: " & UserIndex & ", force close: " & forceClose & ", usrLogged: " & .flags.UserLogged & ", Saliendo: " & .Counters.Saliendo)

            End If

106         If .flags.UserLogged And Not .Counters.Saliendo Then
108             .Counters.Saliendo = True
110             .Counters.Salir = IntervaloCerrarConexion

112             If .flags.Traveling = 1 Then
                    ' Msg576=Se ha cancelado el viaje a casa
114                 Call WriteLocaleMsg(UserIndex, "576", e_FontTypeNames.FONTTYPE_INFO)
116                 .flags.Traveling = 0
118                 .Counters.goHome = 0

                End If

120             If .flags.invisible + .flags.Oculto > 0 Then
122                 .flags.invisible = 0
124                 .flags.Oculto = 0
126                 .Counters.DisabledInvisibility = 0
128                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charIndex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                    ' Msg577=Has vuelto a ser visible
130                 Call WriteLocaleMsg(UserIndex, "577", e_FontTypeNames.FONTTYPE_INFO)

                End If

                'HarThaoS: Captura de bandera
132             If .flags.jugando_captura = 1 Then
134                 If Not InstanciaCaptura Is Nothing Then
136                     Call InstanciaCaptura.eliminarParticipante(InstanciaCaptura.GetPlayer(UserIndex))

                    End If

                End If

138             Call WriteLocaleMsg(UserIndex, "203", e_FontTypeNames.FONTTYPE_INFO, .Counters.Salir)

140             If EsGM(UserIndex) Or MapInfo(.pos.Map).Seguro = 1 Or forceClose Then
142                 Call WriteDisconnect(UserIndex)
144                 Call CloseSocket(UserIndex)

                End If

            End If

        End With

        Exit Sub
Cerrar_Usuario_Err:
146     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.Cerrar_Usuario", Erl)

End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.
Public Sub CancelExit(ByVal UserIndex As Integer)
        On Error GoTo CancelExit_Err

100     If UserList(UserIndex).Counters.Saliendo And UserList(UserIndex).ConnectionDetails.ConnIDValida Then

            ' Is the user still connected?
102         If UserList(UserIndex).ConnectionDetails.ConnIDValida Then
104             UserList(UserIndex).Counters.Saliendo = False
106             UserList(UserIndex).Counters.Salir = 0
                ' Msg578=/salir cancelado.
108             Call WriteLocaleMsg(UserIndex, "578", e_FontTypeNames.FONTTYPE_WARNING)
            Else

                'Simply reset
110             If UserList(UserIndex).flags.Privilegios = e_PlayerType.User And MapInfo(UserList(UserIndex).pos.Map).Seguro = 0 Then
112                 UserList(UserIndex).Counters.Salir = IntervaloCerrarConexion
                Else
                    ' Msg579=Gracias por jugar Argentum Online.
114                 Call WriteLocaleMsg(UserIndex, "579", e_FontTypeNames.FONTTYPE_INFO)
116                 Call WriteDisconnect(UserIndex)
118                 Call CloseSocket(UserIndex)

                End If

                'UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And e_PlayerType.User) And MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0, IntervaloCerrarConexion, 0)
            End If

        End If

        Exit Sub
CancelExit_Err:
120     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.CancelExit", Erl)

End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)
        On Error GoTo VolverCriminal_Err

        'Actualiza el tag al cliente
100     With UserList(UserIndex)

102         If MapData(.pos.Map, .pos.x, .pos.y).trigger = 6 Then Exit Sub
104         If .flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero) Then
106             If .Faccion.Status = e_Facciones.Armada Then
                    ' WyroX: NUNCA debería pasar, pero dejo un log por si las...
108                 Call TraceError(111, "Un personaje de la Armada Real atacó un ciudadano.", "UsUaRiOs.VolverCriminal")

                    'Call ExpulsarFaccionReal(UserIndex)
                End If

            End If

110         If .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then Exit Sub
112         If .Faccion.Status = e_Facciones.Ciudadano Then
114             .Faccion.FactionScore = 0

            End If

116         .Faccion.Status = 0

118         If MapInfo(.pos.Map).NoPKs And Not EsGM(UserIndex) And MapInfo(.pos.Map).Salida.Map <> 0 Then
                ' Msg580=En este mapa no se admiten criminales.
120             Call WriteLocaleMsg(UserIndex, "580", e_FontTypeNames.FONTTYPE_INFO)
122             Call WarpUserChar(UserIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
            Else
124             Call RefreshCharStatus(UserIndex)

            End If

        End With

        Exit Sub
VolverCriminal_Err:
126     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.VolverCriminal", Erl)

End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
        'Actualiza el tag al cliente.
        On Error GoTo VolverCiudadano_Err

100     With UserList(UserIndex)

102         If MapData(.pos.Map, .pos.x, .pos.y).trigger = 6 Then Exit Sub
104         If .Faccion.Status = e_Facciones.Criminal Or .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then
106             .Faccion.FactionScore = 0

            End If

108         .Faccion.Status = e_Facciones.Ciudadano

110         If MapInfo(.pos.Map).NoCiudadanos And Not EsGM(UserIndex) And MapInfo(.pos.Map).Salida.Map <> 0 Then
                ' Msg581=En este mapa no se admiten ciudadanos.
112             Call WriteLocaleMsg(UserIndex, "581", e_FontTypeNames.FONTTYPE_INFO)
114             Call WarpUserChar(UserIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
            Else
116             Call RefreshCharStatus(UserIndex)

            End If

118         Call WriteSafeModeOn(UserIndex)
120         .flags.Seguro = True

        End With

        Exit Sub
VolverCiudadano_Err:
122     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.VolverCiudadano", Erl)

End Sub

Public Function getMaxInventorySlots(ByVal UserIndex As Integer) As Byte
        On Error GoTo getMaxInventorySlots_Err
100     getMaxInventorySlots = MAX_USERINVENTORY_SLOTS

102     With UserList(UserIndex)
104         getMaxInventorySlots = get_num_inv_slots_from_tier(.Stats.tipoUsuario)

        End With

        Exit Function
getMaxInventorySlots_Err:
106     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.getMaxInventorySlots", Erl)

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
100     PermiteMascotas = MapInfo(UserList(UserIndex).pos.Map).NoMascotas = False

102     For i = 1 To MAXMASCOTAS
104         Index = UserList(UserIndex).MascotasIndex(i).ArrayIndex

106         If IsValidNpcRef(UserList(UserIndex).MascotasIndex(i)) Then
108             iMinHP = NpcList(Index).Stats.MinHp
110             PetTiempoDeVida = NpcList(Index).Contadores.TiempoExistencia
112             Call SetUserRef(NpcList(Index).MaestroUser, 0)
114             Call QuitarNPC(Index, eRemoveWarpPets)

116             If PetTiempoDeVida > 0 Then
118                 Call QuitarMascota(UserIndex, Index)
120                 ElementalQuitado = True
122             ElseIf Not PermiteMascotas Then
124                 Call ClearNpcRef(UserList(UserIndex).MascotasIndex(i))
126                 MascotaQuitada = True

                End If

            Else
128             Call ClearNpcRef(UserList(UserIndex).MascotasIndex(i))
130             iMinHP = 0
132             PetTiempoDeVida = 0

            End If

134         petType = UserList(UserIndex).MascotasType(i)

136         If petType > 0 And PermiteMascotas And (UserList(UserIndex).flags.MascotasGuardadas = 0 Or UserList(UserIndex).MascotasIndex(i).ArrayIndex > 0) And PetTiempoDeVida = 0 Then
                Dim SpawnPos As t_WorldPos
138             SpawnPos.Map = UserList(UserIndex).pos.Map
140             SpawnPos.x = UserList(UserIndex).pos.x + RandomNumber(-3, 3)
142             SpawnPos.y = UserList(UserIndex).pos.y + RandomNumber(-3, 3)
144             Index = SpawnNpc(petType, SpawnPos, False, False, False, UserIndex)

                'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
                ' Exception: Pets don't spawn in water if they can't swim
146             If Index > 0 Then
148                 Call SetNpcRef(UserList(UserIndex).MascotasIndex(i), Index)

                    ' Nos aseguramos de que conserve el hp, si estaba danado
150                 If iMinHP Then NpcList(Index).Stats.MinHp = iMinHP
152                 Call SetUserRef(NpcList(Index).MaestroUser, UserIndex)
154                 Call FollowAmo(Index)
                Else
156                 SpawnInvalido = True

                End If

            End If

158     Next i

160     If MascotaQuitada Then
162         If Not PermiteMascotas Then
                ' Msg582=Una fuerza superior impide que tus mascotas entren en este mapa. Estas te esperarán afuera.
164             Call WriteLocaleMsg(UserIndex, "582", e_FontTypeNames.FONTTYPE_INFO)

            End If

166     ElseIf SpawnInvalido Then
            ' Msg583=Tus mascotas no pueden transitar este mapa.
168         Call WriteLocaleMsg(UserIndex, "583", e_FontTypeNames.FONTTYPE_INFO)
170     ElseIf ElementalQuitado Then
            ' Msg584=Pierdes el control de tus mascotas invocadas.
172         Call WriteLocaleMsg(UserIndex, "584", e_FontTypeNames.FONTTYPE_INFO)

        End If

        Exit Sub
WarpMascotas_Err:
174     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.WarpMascotas", Erl)

End Sub

Function TieneArmaduraCazador(ByVal UserIndex As Integer) As Boolean
        On Error GoTo TieneArmaduraCazador_Err

100     If UserList(UserIndex).invent.ArmourEqpObjIndex > 0 Then
102         If ObjData(UserList(UserIndex).invent.ArmourEqpObjIndex).Subtipo = 3 Then ' Aguante hardcodear números :D
104             TieneArmaduraCazador = True

            End If

        End If

        Exit Function
TieneArmaduraCazador_Err:
106     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.TieneArmaduraCazador", Erl)

End Function

Public Sub SetModoConsulta(ByVal UserIndex As Integer)
        Dim sndNick As String

100     With UserList(UserIndex)
102         sndNick = .name

104         If .flags.EnConsulta Then
106             sndNick = sndNick & " [CONSULTA]"
            Else

108             If .GuildIndex > 0 Then
110                 sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"

                End If

            End If

112         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, .Faccion.Status, sndNick))

        End With

End Sub

' Autor: WyroX - 20/01/2021
' Intenta moverlo hacia un "costado" según el heading indicado.
' Si no hay un lugar válido a los lados, lo mueve a la posición válida más cercana.
Sub MoveUserToSide(ByVal UserIndex As Integer, ByVal Heading As e_Heading)
        On Error GoTo Handler

100     With UserList(UserIndex)
            ' Elegimos un lado al azar
            Dim r As Integer
102         r = RandomNumber(0, 1) * 2 - 1 ' -1 o 1
            ' Roto el heading original hacia ese lado
104         Heading = Rotate_Heading(Heading, r)

            ' Intento moverlo para ese lado
106         If MoveUserChar(UserIndex, Heading) Then
                ' Le aviso al usuario que fue movido
108             Call WriteForceCharMove(UserIndex, Heading)
                Exit Sub

            End If

            ' Si falló, intento moverlo para el lado opuesto
110         Heading = InvertHeading(Heading)

112         If MoveUserChar(UserIndex, Heading) Then
                ' Le aviso al usuario que fue movido
114             Call WriteForceCharMove(UserIndex, Heading)
                Exit Sub

            End If

            ' Si ambos fallan, entonces lo dejo en la posición válida más cercana
            Dim NuevaPos As t_WorldPos
116         Call ClosestLegalPos(.pos, NuevaPos, .flags.Navegando, .flags.Navegando = 0)
118         Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.x, NuevaPos.y)

        End With

        Exit Sub
Handler:
120     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.MoveUserToSide", Erl)

End Sub

' Autor: WyroX - 02/03/2021
' Quita parálisis, veneno, invisibilidad, estupidez, mimetismo, deja de descansar, de meditar y de ocultarse; y quita otros estados obsoletos (por si acaso)
Public Sub LimpiarEstadosAlterados(ByVal UserIndex As Integer)
        On Error GoTo Handler

100     With UserList(UserIndex)
            '<<<< Envenenamiento >>>>
102         .flags.Envenenado = 0

            '<<<< Paralisis >>>>
104         If .flags.Paralizado = 1 Then
106             .flags.Paralizado = 0
108             Call WriteParalizeOK(UserIndex)

            End If

            '<<<< Inmovilizado >>>>
110         If .flags.Inmovilizado = 1 Then
112             .flags.Inmovilizado = 0
114             Call WriteInmovilizaOK(UserIndex)

            End If

            '<<< Estupidez >>>
116         If .flags.Estupidez = 1 Then
118             .flags.Estupidez = 0
120             Call WriteDumbNoMore(UserIndex)

            End If

            '<<<< Descansando >>>>
122         If .flags.Descansar Then
124             .flags.Descansar = False
126             Call WriteRestOK(UserIndex)

            End If

            '<<<< Meditando >>>>
128         If .flags.Meditando Then
130             .flags.Meditando = False
132             .Char.FX = 0
134             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charIndex, 0, .pos.x, .pos.y))

            End If

            '<<<< Stun >>>>
136         .Counters.StunEndTime = 0

            '<<<< Invisible >>>>
138         If (.flags.invisible = 1 Or .flags.Oculto = 1) And .flags.AdminInvisible = 0 Then
140             .flags.Oculto = 0
142             .flags.invisible = 0
144             .Counters.TiempoOculto = 0
146             .Counters.Invisibilidad = 0
148             .Counters.DisabledInvisibility = 0
150             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charIndex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))

            End If

            '<<<< Mimetismo >>>>
152         If .flags.Mimetizado > 0 Then
154             If .flags.Navegando Then
156                 If .flags.Muerto = 0 Then
158                     .Char.body = ObjData(UserList(UserIndex).invent.BarcoObjIndex).Ropaje
                    Else
160                     .Char.body = iFragataFantasmal

                    End If

162                 Call ClearClothes(.Char)
                Else
164                 .Char.body = .CharMimetizado.body
166                 .Char.head = .CharMimetizado.head
168                 .Char.CascoAnim = .CharMimetizado.CascoAnim
170                 .Char.ShieldAnim = .CharMimetizado.ShieldAnim
172                 .Char.WeaponAnim = .CharMimetizado.WeaponAnim
174                 .Char.CartAnim = .CharMimetizado.CartAnim

                End If

176             .Counters.Mimetismo = 0
178             .flags.Mimetizado = e_EstadoMimetismo.Desactivado

            End If

            '<<<< Estados obsoletos >>>>
180         .flags.Incinerado = 0

        End With

        Exit Sub
Handler:
182     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.LimpiarEstadosAlterados", Erl)

End Sub

Public Sub DevolverPosAnterior(ByVal UserIndex As Integer)

100     With UserList(UserIndex).flags
102         Call WarpToLegalPos(UserIndex, .LastPos.Map, .LastPos.x, .LastPos.y, True)

        End With

End Sub

Public Function ActualizarVelocidadDeUsuario(ByVal UserIndex As Integer) As Single
        On Error GoTo ActualizarVelocidadDeUsuario_Err
        Dim velocidad As Single, modificadorItem As Single, modificadorHechizo As Single
100     velocidad = VelocidadNormal
102     modificadorItem = 1
104     modificadorHechizo = 1

106     With UserList(UserIndex)

108         If .flags.Muerto = 1 Then
110             velocidad = VelocidadMuerto
112             GoTo UpdateSpeed ' Los muertos no tienen modificadores de velocidad

            End If

            ' El traje para nadar es considerado barco, de subtipo = 0
114         If (.flags.Navegando + .flags.Nadando > 0) And (.invent.BarcoObjIndex > 0) Then
116             modificadorItem = ObjData(.invent.BarcoObjIndex).velocidad

            End If

118         If (.flags.Montado = 1) And (.invent.MonturaObjIndex > 0) Then
120             modificadorItem = ObjData(.invent.MonturaObjIndex).velocidad

            End If

            ' Algun hechizo le afecto la velocidad
122         If .flags.VelocidadHechizada > 0 Then
124             modificadorHechizo = .flags.VelocidadHechizada

            End If

126         velocidad = VelocidadNormal * modificadorItem * modificadorHechizo * max(0, (1 + .Modifiers.MovementSpeed))
UpdateSpeed:
128         .Char.speeding = velocidad
130         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.charIndex, .Char.speeding))
132         Call WriteVelocidadToggle(UserIndex)

        End With

        Exit Function
ActualizarVelocidadDeUsuario_Err:
134     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.CalcularVelocidad_Err", Erl)

End Function

Public Sub ClearClothes(ByRef Char As t_Char)
100     Char.ShieldAnim = NingunEscudo
102     Char.WeaponAnim = NingunArma
104     Char.CascoAnim = NingunCasco
106     Char.CartAnim = NoCart

End Sub

Public Function IsStun(ByRef flags As t_UserFlags, _
                       ByRef Counters As t_UserCounters) As Boolean
100     IsStun = Counters.StunEndTime > GetTickCount()

End Function

Public Function CanMove(ByRef flags As t_UserFlags, _
                        ByRef Counters As t_UserCounters) As Boolean
100     CanMove = flags.Paralizado = 0 And flags.Inmovilizado = 0 And Not IsStun(flags, Counters) And Not flags.TranslationActive

End Function

Public Function StunPlayer(ByVal UserIndex As Integer, _
                           ByRef Counters As t_UserCounters) As Boolean
        Dim CurrTime As Long
100     StunPlayer = False

102     If Not CanMove(UserList(UserIndex).flags, Counters) Then Exit Function
104     If IsSet(UserList(UserIndex).flags.StatusMask, eCCInmunity) Then Exit Function
106     CurrTime = GetTickCount()

108     If CurrTime > Counters.StunEndTime + PlayerInmuneTime Then
110         Counters.StunEndTime = GetTickCount() + PlayerStunTime
112         StunPlayer = True

        End If

End Function

Public Function CanUseItem(ByRef flags As t_UserFlags, _
                           ByRef Counters As t_UserCounters) As Boolean
100     CanUseItem = True

End Function

Public Sub UpdateCd(ByVal UserIndex As Integer, ByVal cdType As e_CdTypes)
100     UserList(UserIndex).CdTimes(cdType) = GetTickCount()
102     Call WriteUpdateCdType(UserIndex, cdType)

End Sub

Public Function IsVisible(ByRef User As t_User) As Boolean
100     IsVisible = (Not (User.flags.invisible > 0 Or User.flags.Oculto > 0))

End Function

Public Function CanHelpUser(ByVal UserIndex As Integer, _
                            ByVal targetUserIndex As Integer) As e_InteractionResult
100     CanHelpUser = eInteractionOk

102     If UserList(UserIndex).flags.CurrentTeam > 0 And UserList(UserIndex).flags.CurrentTeam <> UserList(targetUserIndex).flags.CurrentTeam Then
104         CanHelpUser = eDifferentTeam
            Exit Function

        End If

106     If PeleaSegura(UserIndex, targetUserIndex) Then
            Exit Function

        End If

        Dim TargetStatus As e_Facciones
108     TargetStatus = Status(targetUserIndex)

110     Select Case Status(UserIndex)

            Case e_Facciones.Ciudadano, e_Facciones.Armada, e_Facciones.consejo

112             If TargetStatus = e_Facciones.Caos Or TargetStatus = e_Facciones.concilio Then
114                 CanHelpUser = eOposingFaction
                    Exit Function
116             ElseIf TargetStatus = e_Facciones.Criminal Then

118                 If UserList(UserIndex).flags.Seguro Then
120                     CanHelpUser = eCantHelpCriminal
                    Else

122                     If UserList(UserIndex).GuildIndex > 0 Then

                            'Si el clan es de alineación ciudadana.
124                         If GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                                'No lo dejo resucitarlo
126                             CanHelpUser = eCantHelpCriminalClanRules
                                Exit Function
                                'Si es de alineación neutral, lo dejo resucitar y lo vuelvo criminal
128                         ElseIf GuildAlignmentIndex(UserList(UserIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_NEUTRAL Then
130                             Call VolverCriminal(UserIndex)
132                             Call RefreshCharStatus(UserIndex)
                                Exit Function

                            End If

                        Else
134                         Call VolverCriminal(UserIndex)
136                         Call RefreshCharStatus(UserIndex)
                            Exit Function

                        End If

                    End If

                End If

138         Case e_Facciones.Caos, e_Facciones.concilio

140             If Status(targetUserIndex) = e_Facciones.Armada Or Status(targetUserIndex) = e_Facciones.consejo Or Status(targetUserIndex) = e_Facciones.Ciudadano Then
142                 CanHelpUser = eOposingFaction

                End If

144         Case Else
                Exit Function

        End Select

End Function

Public Function CanAttackUser(ByVal attackerIndex As Integer, _
                              ByVal AttackerVersionID As Integer, _
                              ByVal TargetIndex As Integer, _
                              ByVal TargetVersionID As Integer) As e_AttackInteractionResult

100     If UserList(TargetIndex).flags.Muerto = 1 Then
102         CanAttackUser = e_AttackInteractionResult.eDeathTarget
            Exit Function

        End If

104     If attackerIndex = TargetIndex And AttackerVersionID = TargetVersionID Then
106         CanAttackUser = e_AttackInteractionResult.eCantAttackYourself
            Exit Function

        End If

108     If UserList(attackerIndex).flags.EnReto Then
110         If Retos.Salas(UserList(attackerIndex).flags.SalaReto).TiempoItems > 0 Then
112             CanAttackUser = e_AttackInteractionResult.eFightActive
                Exit Function

            End If

        End If

114     If UserList(attackerIndex).Grupo.Id > 0 And UserList(TargetIndex).Grupo.Id > 0 And UserList(attackerIndex).Grupo.Id = UserList(TargetIndex).Grupo.Id Then
116         CanAttackUser = eSameGroup
            Exit Function

        End If

118     If UserList(attackerIndex).flags.EnConsulta Or UserList(TargetIndex).flags.EnConsulta Then
120         CanAttackUser = eTalkWithMaster
            Exit Function

        End If

122     If UserList(attackerIndex).flags.Maldicion = 1 Then
124         CanAttackUser = eAttackerIsCursed
            Exit Function

        End If

126     If UserList(attackerIndex).flags.Montado = 1 Then
128         CanAttackUser = eMounted
            Exit Function

        End If

130     If Not MapInfo(UserList(TargetIndex).pos.Map).FriendlyFire And UserList(TargetIndex).flags.CurrentTeam > 0 And UserList(TargetIndex).flags.CurrentTeam = UserList(attackerIndex).flags.CurrentTeam Then
132         CanAttackUser = eSameTeam
            Exit Function

        End If

        ' Nueva verificación específica para Captura la Bandera
134     If UserList(attackerIndex).flags.jugando_captura = 1 And UserList(TargetIndex).flags.jugando_captura = 1 Then
136         If UserList(attackerIndex).flags.CurrentTeam = UserList(TargetIndex).flags.CurrentTeam Then
                'Msg1102= ¡No puedes atacar a miembros de tu propio equipo!
138             Call WriteLocaleMsg(attackerIndex, "1102", e_FontTypeNames.FONTTYPE_INFO)
140             CanAttackUser = eSameTeam
                Exit Function

            End If

        End If

        Dim t As e_Trigger6
        'Estamos en una Arena? o un trigger zona segura?
142     t = TriggerZonaPelea(attackerIndex, TargetIndex)

144     If t = e_Trigger6.TRIGGER6_PERMITE Then
146         CanAttackUser = eCanAttack
            Exit Function
148     ElseIf PeleaSegura(attackerIndex, TargetIndex) Then
150         CanAttackUser = eCanAttack
            Exit Function

        End If

        'Solo administradores pueden atacar a usuarios (PARA TESTING)
152     If (UserList(attackerIndex).flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Admin)) = 0 Then
154         CanAttackUser = eNotEnougthPrivileges
            Exit Function

        End If

        'Estas queriendo atacar a un GM?
        Dim rank As Integer
156     rank = e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero

158     If (UserList(TargetIndex).flags.Privilegios And rank) > (UserList(attackerIndex).flags.Privilegios And rank) Then
160         CanAttackUser = eNotEnougthPrivileges
            Exit Function

        End If

        ' Seguro Clan
162     If UserList(attackerIndex).GuildIndex > 0 Then
164         If UserList(attackerIndex).flags.SeguroClan And NivelDeClan(UserList(attackerIndex).GuildIndex) >= 3 Then
166             If UserList(attackerIndex).GuildIndex = UserList(TargetIndex).GuildIndex Then
168                 CanAttackUser = eSameClan
                    Exit Function

                End If

            End If

        End If

        ' Es armada?
170     If esArmada(attackerIndex) Then

            ' Si ataca otro armada
172         If esArmada(TargetIndex) Then
174             CanAttackUser = eSameFaction
                Exit Function
                ' Si ataca un ciudadano
176         ElseIf esCiudadano(TargetIndex) Then
178             CanAttackUser = eSameFaction
                Exit Function

            End If

            ' No es armada
        Else

            'Tenes puesto el seguro?
180         If (esCiudadano(attackerIndex)) Then
182             If (UserList(attackerIndex).flags.Seguro) Then
184                 If esCiudadano(TargetIndex) Then
186                     CanAttackUser = eRemoveSafe
                        Exit Function
188                 ElseIf esArmada(TargetIndex) Then
190                     CanAttackUser = eRemoveSafe
                        Exit Function

                    End If

                End If

192         ElseIf esCaos(attackerIndex) And esCaos(TargetIndex) Then
194             CanAttackUser = eSameFaction
                Exit Function

            End If

        End If

        'Estas en un Mapa Seguro?
196     If MapInfo(UserList(TargetIndex).pos.Map).Seguro = 1 Then
198         If esArmada(attackerIndex) Then
200             If UserList(attackerIndex).Faccion.RecompensasReal >= 3 Then
202                 If UserList(TargetIndex).pos.Map = 58 Or UserList(TargetIndex).pos.Map = 59 Or UserList(TargetIndex).pos.Map = 60 Then
204                     CanAttackUser = eCanAttack
                        Exit Function

                    End If

                End If

            End If

206         If esCaos(attackerIndex) Then
208             If UserList(attackerIndex).Faccion.RecompensasCaos >= 3 Then
210                 If UserList(TargetIndex).pos.Map = 195 Or UserList(TargetIndex).pos.Map = 196 Then
212                     CanAttackUser = eCanAttack
                        Exit Function

                    End If

                End If

            End If

214         CanAttackUser = eSafeArea
            Exit Function

        End If

        'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
216     If MapData(UserList(TargetIndex).pos.Map, UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y).trigger = e_Trigger.ZonaSegura Or MapData(UserList(attackerIndex).pos.Map, UserList(attackerIndex).pos.x, UserList(attackerIndex).pos.y).trigger = e_Trigger.ZonaSegura Then
218         CanAttackUser = eSafeArea
            Exit Function

        End If

220     CanAttackUser = eCanAttack

End Function

Public Function ModifyHealth(ByVal UserIndex As Integer, _
                             ByVal amount As Long, _
                             Optional ByVal MinValue = 0) As Boolean

100     With UserList(UserIndex)
102         ModifyHealth = False
104         .Stats.MinHp = .Stats.MinHp + amount

106         If .Stats.MinHp > .Stats.MaxHp Then
108             .Stats.MinHp = .Stats.MaxHp

            End If

110         If .Stats.MinHp <= MinValue Then
112             .Stats.MinHp = MinValue
114             ModifyHealth = True

            End If

116         Call WriteUpdateHP(UserIndex)

        End With

End Function

Public Function ModifyStamina(ByVal UserIndex As Integer, _
                              ByVal amount As Integer, _
                              ByVal CancelIfNotEnought As Boolean, _
                              Optional ByVal MinValue = 0) As Boolean
100     ModifyStamina = False

102     With UserList(UserIndex)

104         If CancelIfNotEnought And amount < 0 And .Stats.MinSta < Abs(amount) Then
106             ModifyStamina = True
                Exit Function

            End If

108         .Stats.MinSta = .Stats.MinSta + amount

110         If .Stats.MinSta > .Stats.MaxSta Then
112             .Stats.MinSta = .Stats.MaxSta

            End If

114         If .Stats.MinSta < MinValue Then
116             .Stats.MinSta = MinValue
118             ModifyStamina = True

            End If

120         Call WriteUpdateSta(UserIndex)

        End With

End Function

Public Function ModifyMana(ByVal UserIndex As Integer, _
                           ByVal amount As Integer, _
                           ByVal CancelIfNotEnought As Boolean, _
                           Optional ByVal MinValue = 0) As Boolean
100     ModifyMana = False

102     With UserList(UserIndex)

104         If CancelIfNotEnought And amount < 0 And .Stats.MinMAN < Abs(amount) Then
106             ModifyMana = True
                Exit Function

            End If

108         .Stats.MinMAN = .Stats.MinMAN + amount

110         If .Stats.MinMAN > .Stats.MaxMAN Then
112             .Stats.MinMAN = .Stats.MaxMAN

            End If

114         If .Stats.MinMAN < MinValue Then
116             .Stats.MinMAN = MinValue
118             ModifyMana = True

            End If

120         Call WriteUpdateMana(UserIndex)

        End With

End Function

Public Sub ResurrectUser(ByVal UserIndex As Integer)
        ' Msg585=¡Has sido resucitado!
100     Call WriteLocaleMsg(UserIndex, "585", e_FontTypeNames.FONTTYPE_INFO)
102     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charIndex, e_ParticulasIndex.Resucitar, 250, True))
104     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(117, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
106     Call RevivirUsuario(UserIndex, True)
108     Call WriteUpdateHungerAndThirst(UserIndex)

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
100     DamageStr = PonerPuntos(amount)

102     If amount > 0 Then
104         Color = vbGreen
        Else
106         Color = DamageColor

        End If

108     If amount < 0 Then
110         DamageStr = PonerPuntos(Math.Abs(amount))

112         If SourceType = eUser Then
114             If UserList(SourceIndex).ChatCombate = 1 And DoDamageText > 0 Then
116                 Call WriteLocaleMsg(SourceIndex, DoDamageText, e_FontTypeNames.FONTTYPE_FIGHT, UserList(UserIndex).name & "¬" & DamageStr)

                End If

118             If UserList(UserIndex).ChatCombate = 1 And GotDamageText > 0 Then
120                 Call WriteLocaleMsg(UserIndex, GotDamageText, e_FontTypeNames.FONTTYPE_FIGHT, UserList(SourceIndex).name & "¬" & DamageStr)

                End If

            End If

122         amount = EffectsOverTime.TargetApplyDamageReduction(UserList(UserIndex).EffectOverTime, amount, SourceIndex, SourceType, DamageSourceType)
124         Call EffectsOverTime.TargetWasDamaged(UserList(UserIndex).EffectOverTime, SourceIndex, SourceType, DamageSourceType)

        End If

126     With UserList(UserIndex)

128         If IsVisible(UserList(UserIndex)) Then
130             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageTextOverChar(DamageStr, .Char.charIndex, Color))
            Else
132             Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageTextOverChar(DamageStr, .Char.charIndex, Color))

            End If

134         If ModifyHealth(UserIndex, amount) Then
136             Call TargetWasDamaged(UserList(UserIndex).EffectOverTime, SourceIndex, SourceType, DamageSourceType)
138             Call CustomScenarios.UserDie(UserIndex)

140             If SourceType = eUser Then
142                 Call ContarMuerte(UserIndex, SourceIndex)
144                 Call PlayerKillPlayer(.pos.Map, SourceIndex, UserIndex, DamageSourceType, DamageSourceIndex)
146                 Call ActStats(UserIndex, SourceIndex)
                Else
148                 Call NPcKillPlayer(.pos.Map, SourceIndex, UserIndex, DamageSourceType, DamageSourceIndex)
150                 Call WriteNPCKillUser(UserIndex)

152                 If IsValidUserRef(NpcList(SourceIndex).MaestroUser) Then
154                     Call AllFollowAmo(NpcList(SourceIndex).MaestroUser.ArrayIndex)
156                     Call PlayerKillPlayer(.pos.Map, NpcList(SourceIndex).MaestroUser.ArrayIndex, UserIndex, e_DamageSourceType.e_pet, 0)
                    Else
                        'Al matarlo no lo sigue mas
158                     Call SetMovement(SourceIndex, NpcList(SourceIndex).flags.OldMovement)
160                     NpcList(SourceIndex).Hostile = NpcList(SourceIndex).flags.OldHostil
162                     NpcList(SourceIndex).flags.AttackedBy = vbNullString
164                     Call SetUserRef(NpcList(SourceIndex).TargetUser, 0)

                    End If

166                 Call UserMod.UserDie(UserIndex)

                End If

168             DoDamageOrHeal = eDead
                Exit Function

            End If

        End With

170     DoDamageOrHeal = eStillAlive
        Exit Function
DoDamageOrHeal_Err:
172     Call TraceError(Err.Number, Err.Description, "UserMod.DoDamageOrHeal", Erl)

End Function

Public Function GetPhysicalDamageModifier(ByRef User As t_User) As Single
100     GetPhysicalDamageModifier = max(1 + User.Modifiers.PhysicalDamageBonus, 0)

End Function

Public Function GetMagicDamageModifier(ByRef User As t_User) As Single
100     GetMagicDamageModifier = max(1 + User.Modifiers.MagicDamageBonus, 0)

End Function

Public Function GetMagicDamageReduction(ByRef User As t_User) As Single
100     GetMagicDamageReduction = max(1 - User.Modifiers.MagicDamageReduction, 0)

End Function

Public Function GetPhysicDamageReduction(ByRef User As t_User) As Single
100     GetPhysicDamageReduction = max(1 - User.Modifiers.PhysicalDamageReduction, 0)

End Function

Public Sub RemoveInvisibility(ByVal UserIndex As Integer)

100     With UserList(UserIndex)

102         If .flags.invisible + .flags.Oculto > 0 And .flags.NoDetectable = 0 Then
104             .flags.invisible = 0
106             .flags.Oculto = 0
108             .Counters.Invisibilidad = 0
110             .Counters.Ocultando = 0
112             .Counters.DisabledInvisibility = 0
                ' Msg591=Tu invisibilidad ya no tiene efecto.
114             Call WriteLocaleMsg(UserIndex, "591", e_FontTypeNames.FONTTYPE_INFOIAO)
116             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charIndex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))

            End If

        End With

End Sub

Public Function Inmovilize(ByVal SourceIndex As Integer, _
                           ByVal TargetIndex As Integer, _
                           ByVal Time As Integer, _
                           ByVal FX As Integer) As Boolean
100     Call UsuarioAtacadoPorUsuario(SourceIndex, TargetIndex)

102     If IsSet(UserList(TargetIndex).flags.StatusMask, eCCInmunity) Then
104         Call WriteLocaleMsg(SourceIndex, MsgCCInunity, e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Function

        End If

106     If CanMove(UserList(TargetIndex).flags, UserList(TargetIndex).Counters) Then
108         UserList(TargetIndex).Counters.Inmovilizado = Time
110         UserList(TargetIndex).flags.Inmovilizado = 1
112         Call SendData(SendTarget.ToPCAliveArea, TargetIndex, PrepareMessageCreateFX(UserList(TargetIndex).Char.charIndex, FX, 0, UserList(TargetIndex).pos.x, UserList(TargetIndex).pos.y))
114         Call WriteInmovilizaOK(TargetIndex)
116         Call WritePosUpdate(TargetIndex)
118         Inmovilize = True

        End If

End Function

Public Function GetArmorPenetration(ByVal UserIndex As Integer, _
                                    ByVal TargetArmor As Integer) As Integer
        Dim ArmorPenetration As Integer

100     If Not IsFeatureEnabled("armor_penetration_feature") Then Exit Function

102     With UserList(UserIndex)

104         If .invent.WeaponEqpObjIndex > 0 Then
106             ArmorPenetration = ObjData(.invent.WeaponEqpObjIndex).IgnoreArmorAmmount

108             If ObjData(.invent.WeaponEqpObjIndex).IgnoreArmorPercent > 0 Then
110                 ArmorPenetration = ArmorPenetration + TargetArmor * ObjData(.invent.WeaponEqpObjIndex).IgnoreArmorPercent

                End If

            End If

        End With

112     GetArmorPenetration = ArmorPenetration

End Function

Public Function GetEvasionBonus(ByRef User As t_User) As Integer
100     GetEvasionBonus = User.Modifiers.EvasionBonus

End Function

Public Function GetHitBonus(ByRef User As t_User) As Integer
100     GetHitBonus = User.Modifiers.HitBonus + GetWeaponHitBonus(User.invent.WeaponEqpObjIndex, User.clase)

End Function

'Defines the healing bonus when using a potion, a spell or any other healing source
Public Function GetSelfHealingBonus(ByRef User As t_User) As Single
100     GetSelfHealingBonus = max(1 + User.Modifiers.SelfHealingBonus, 0)

End Function

'Defines bonus when healing someone with magic
Public Function GetMagicHealingBonus(ByRef User As t_User) As Single
100     GetMagicHealingBonus = max(1 + User.Modifiers.MagicHealingBonus, 0)

End Function

Public Function GetWeaponHitBonus(ByVal WeaponIndex As Integer, _
                                  ByVal UserClass As e_Class)
        On Error GoTo GetWeaponHitBonus_Err

100     If WeaponIndex = 0 Then Exit Function
102     If Not IsFeatureEnabled("class_weapon_bonus") Or ObjData(WeaponIndex).WeaponType = 0 Then Exit Function
104     GetWeaponHitBonus = ModClase(UserClass).WeaponHitBonus(ObjData(WeaponIndex).WeaponType)
        Exit Function
GetWeaponHitBonus_Err:
106     Call TraceError(Err.Number, Err.Description, "UserMod.GetWeaponHitBonus WeaponIndex: " & WeaponIndex & " for class: " & UserClass, Erl)

End Function

Public Sub RemoveUserInvisibility(ByVal UserIndex As Integer)

100     With UserList(UserIndex)
            Dim RemoveHiddenState As Boolean

            ' Volver visible
102         If .flags.Oculto = 1 And .flags.AdminInvisible = 0 And .flags.invisible = 0 Then
104             .flags.Oculto = 0
106             .Counters.TiempoOculto = 0
                'Msg307=Has vuelto a ser visible.
108             Call WriteLocaleMsg(UserIndex, "307", e_FontTypeNames.FONTTYPE_INFO)
110             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charIndex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))

            End If

112         If IsFeatureEnabled("remove-inv-on-attack") And Not MapInfo(.pos.Map).KeepInviOnAttack Then
114             RemoveHiddenState = .flags.Oculto > 0 Or .flags.invisible > 0

            End If

            'I see you...
116         If RemoveHiddenState And .flags.AdminInvisible = 0 Then
118             .flags.Oculto = 0
120             .flags.invisible = 0
122             .Counters.Invisibilidad = 0
124             .Counters.TiempoOculto = 0
126             .Counters.LastAttackTime = GlobalFrameTime

128             If .flags.Navegando = 1 Then
130                 If .clase = e_Class.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
132                     Call EquiparBarco(UserIndex)
                        ' Msg592=¡Has recuperado tu apariencia normal!
134                     Call WriteLocaleMsg(UserIndex, "592", e_FontTypeNames.FONTTYPE_INFO)
136                     Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NoCart)
138                     Call RefreshCharStatus(UserIndex)

                    End If

                Else

140                 If .flags.invisible = 0 Then
142                     Call WriteLocaleMsg(UserIndex, "307", e_FontTypeNames.FONTTYPE_INFO)
144                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charIndex, False, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))

                    End If

                End If

            End If

        End With

End Sub

Public Function UserHasSpell(ByVal UserIndex As Integer, _
                             ByVal SpellId As Integer) As Boolean

100     With UserList(UserIndex)
            Dim i As Integer

102         For i = LBound(.Stats.UserHechizos) To UBound(.Stats.UserHechizos)

104             If .Stats.UserHechizos(i) = SpellId Then
106                 UserHasSpell = True
                    Exit Function

                End If

108         Next i

        End With

End Function

Public Function GetLinearDamageBonus(ByVal UserIndex As Integer) As Integer
100     GetLinearDamageBonus = UserList(UserIndex).Modifiers.PhysicalDamageLinearBonus

End Function

Public Function GetDefenseBonus(ByVal UserIndex As Integer) As Integer
100     GetDefenseBonus = UserList(UserIndex).Modifiers.DefenseBonus

End Function

Public Function GetMaxMana(ByVal UserIndex As Integer) As Long

100     With UserList(UserIndex)
102         GetMaxMana = .Stats.UserAtributos(e_Atributos.Inteligencia) * ModClase(.clase).ManaInicial
104         GetMaxMana = GetMaxMana + (ModClase(.clase).MultMana * .Stats.UserAtributos(e_Atributos.Inteligencia)) * (.Stats.ELV - 1)

        End With

End Function

Public Function GetHitModifier(ByVal UserIndex As Integer) As Long

100     With UserList(UserIndex)

102         If .Stats.ELV <= 36 Then
104             GetHitModifier = (.Stats.ELV - 1) * ModClase(.clase).HitPre36
            Else
106             GetHitModifier = 35 * ModClase(.clase).HitPre36
108             GetHitModifier = GetHitModifier + (.Stats.ELV - 36) * ModClase(.clase).HitPost36

            End If

        End With

End Function

Public Function GetMaxStamina(ByVal UserIndex As Integer) As Integer

100     With UserList(UserIndex)
102         GetMaxStamina = 60 + (.Stats.ELV - 1) * ModClase(.clase).AumentoSta

        End With

End Function

Public Function GetMaxHp(ByVal UserIndex As Integer) As Integer

100     With UserList(UserIndex)
102         GetMaxHp = (ModClase(.clase).Vida - (21 - .Stats.UserAtributos(e_Atributos.Constitucion)) * 0.5) * (.Stats.ELV - 1) + .Stats.UserAtributos(e_Atributos.Constitucion)

        End With

End Function

Public Function GetUserSpouse(ByVal UserIndex As Integer) As String

100     With UserList(UserIndex)

102         If .flags.SpouseId = 0 Then
                Exit Function

            End If

104         GetUserSpouse = GetUserName(.flags.SpouseId)

        End With

End Function

Public Sub RegisterNewAttack(ByVal TargetUser As Integer, ByVal attackerIndex As Integer)

100     With UserList(TargetUser)

102         If .Stats.MinHp > 0 Then
104             Call SetUserRef(.flags.LastAttacker, attackerIndex)
106             .flags.LastAttackedByUserTime = GlobalFrameTime

            End If

        End With

End Sub

Public Sub RegisterNewHelp(ByVal TargetUser As Integer, ByVal attackerIndex As Integer)

100     With UserList(TargetUser)
102         Call SetUserRef(.flags.LastHelpUser, attackerIndex)
104         .flags.LastHelpByTime = GlobalFrameTime

        End With

End Sub

Public Sub SaveDCUserCache(ByVal UserIndex As Integer)
        On Error GoTo SaveDCUserCache_Err

100     With UserList(UserIndex)
            Dim InsertIndex As Integer
102         InsertIndex = RecentDCUserCache.LastIndex Mod UBound(RecentDCUserCache.LastDisconnectionInfo)
            Dim i As Integer

104         For i = 0 To MaxRecentKillToStore
106             RecentDCUserCache.LastDisconnectionInfo(InsertIndex).RecentKillers(i) = .flags.RecentKillers(i)
108         Next i

110         RecentDCUserCache.LastDisconnectionInfo(InsertIndex).RecentKillersIndex = .flags.LastKillerIndex
112         RecentDCUserCache.LastDisconnectionInfo(InsertIndex).UserId = .Id
114         RecentDCUserCache.LastIndex = RecentDCUserCache.LastIndex + 1

116         If RecentDCUserCache.LastIndex > UBound(RecentDCUserCache.LastDisconnectionInfo) * 10 Then 'prevent overflow
118             RecentDCUserCache.LastIndex = RecentDCUserCache.LastIndex \ 10

            End If

        End With

        Exit Sub
SaveDCUserCache_Err:
120     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SaveDCUserCache_Err", Erl)

122     Resume Next

End Sub

Public Sub RestoreDCUserCache(ByVal UserIndex As Integer)
        On Error GoTo RestoreDCUserCache_Err

100     With UserList(UserIndex)
            Dim StartIndex As Integer
            Dim EndIndex   As Integer
            Dim ArraySize  As Integer
102         ArraySize = UBound(RecentDCUserCache.LastDisconnectionInfo)
104         StartIndex = max(0, (RecentDCUserCache.LastIndex - ArraySize) Mod ArraySize)
106         EndIndex = ((RecentDCUserCache.LastIndex - 1) Mod ArraySize)
            Dim i As Integer
            Dim j As Integer

108         For i = StartIndex To EndIndex

110             If RecentDCUserCache.LastDisconnectionInfo(i Mod ArraySize).UserId = .Id Then

112                 For j = 0 To MaxRecentKillToStore
114                     .flags.RecentKillers(j) = RecentDCUserCache.LastDisconnectionInfo(i Mod ArraySize).RecentKillers(j)
116                 Next j

118                 .flags.LastKillerIndex = RecentDCUserCache.LastDisconnectionInfo(i Mod ArraySize).RecentKillersIndex
                    Exit Sub

                End If

120         Next i

        End With

        Exit Sub
RestoreDCUserCache_Err:
122     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.RestoreDCUserCache", Erl)

124     Resume Next

End Sub

Public Function GetUserMRForNpc(ByVal UserIndex As Integer) As Integer

100     With UserList(UserIndex)
            Dim MR As Integer
102         MR = 0

104         If .invent.ArmourEqpObjIndex > 0 Then
106             MR = MR + ObjData(.invent.ArmourEqpObjIndex).ResistenciaMagica

            End If

            ' Resistencia mágica anillo
108         If .invent.ResistenciaEqpObjIndex > 0 Then
110             MR = MR + ObjData(.invent.ResistenciaEqpObjIndex).ResistenciaMagica

            End If

            ' Resistencia mágica escudo
112         If .invent.EscudoEqpObjIndex > 0 Then
114             MR = MR + ObjData(.invent.EscudoEqpObjIndex).ResistenciaMagica

            End If

            ' Resistencia mágica casco
116         If .invent.CascoEqpObjIndex > 0 Then
118             MR = MR + ObjData(.invent.CascoEqpObjIndex).ResistenciaMagica

            End If

120         If IsFeatureEnabled("mr-magic-bonus-damage") Then
122             MR = MR + .Stats.UserSkills(Resistencia) * MRSkillNpcProtectionModifier

            End If

124         GetUserMRForNpc = MR + 100 * ModClase(.clase).ResistenciaMagica

        End With

End Function

Public Function GetUserMR(ByVal UserIndex As Integer) As Integer

100     With UserList(UserIndex)
            Dim MR As Integer
102         MR = 0

104         If .invent.ArmourEqpObjIndex > 0 Then
106             MR = MR + ObjData(.invent.ArmourEqpObjIndex).ResistenciaMagica

            End If

            ' Resistencia mágica anillo
108         If .invent.ResistenciaEqpObjIndex > 0 Then
110             MR = MR + ObjData(.invent.ResistenciaEqpObjIndex).ResistenciaMagica

            End If

            ' Resistencia mágica escudo
112         If .invent.EscudoEqpObjIndex > 0 Then
114             MR = MR + ObjData(.invent.EscudoEqpObjIndex).ResistenciaMagica

            End If

            ' Resistencia mágica casco
116         If .invent.CascoEqpObjIndex > 0 Then
118             MR = MR + ObjData(.invent.CascoEqpObjIndex).ResistenciaMagica

            End If

120         If IsFeatureEnabled("mr-magic-bonus-damage") Then
122             MR = MR + .Stats.UserSkills(Resistencia) * MRSkillProtectionModifier

            End If

124         GetUserMR = MR + 100 * ModClase(.clase).ResistenciaMagica

        End With

End Function
