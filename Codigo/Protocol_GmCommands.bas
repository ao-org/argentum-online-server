Attribute VB_Name = "Protocol_GmCommands"
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

' Handles the "Online" message.
Public Sub HandleOnline(ByVal UserIndex As Integer)
    On Error GoTo HandleOnline_Err
    'Ladder 17/12/20 : Envio records de usuarios y uptime
    Dim i     As Long
    Dim count As Long
    With UserList(UserIndex)
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged Then
                count = count + 1
            End If
        Next i
        If EsGM(UserIndex) Then
            Call WriteLocaleMsg(UserIndex, "1468", e_FontTypeNames.FONTTYPE_INFOIAO, count) ' Msg1468=Número de usuarios: ¬1 conectados.
            Call WriteLocaleMsg(UserIndex, "1469", e_FontTypeNames.FONTTYPE_INFOIAO, RecordUsuarios) ' Msg1469=Record de usuarios en simultaneo: ¬1.
        Else
            ' Msg526=Comando deshabilitado para tu rango.
            Call WriteLocaleMsg(UserIndex, "526", e_FontTypeNames.FONTTYPE_INFOIAO)
        End If
    End With
    Exit Sub
HandleOnline_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnline", Erl)
End Sub

' Handles the "Help" message.
Public Sub HandleHelp(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo HandleHelp_Err
    Call SendHelp(UserIndex)
    Exit Sub
HandleHelp_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleHelp", Erl)
End Sub

' Handles the "RequestStats" message.
Public Sub HandleRequestStats(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo HandleRequestStats_Err
    Call SendUserStatsTxt(UserIndex, UserIndex)
    Exit Sub
HandleRequestStats_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestStats", Erl)
End Sub

' Handles the "RequestMOTD" message.
Public Sub HandleRequestMOTD(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo HandleRequestMOTD_Err
    Call SendMOTD(UserIndex)
    Exit Sub
HandleRequestMOTD_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestMOTD", Erl)
End Sub

' Handles the "UpTime" message.
Public Sub HandleUpTime(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo HandleUpTime_Err
    Dim Time      As Long
    Dim UpTimeStr As String
    'Get total time in seconds
    Time = TicksElapsed(tInicioServer, GetTickCountRaw()) \ 1000
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (Time Mod 60) & " segundos."
    Time = Time \ 60
    UpTimeStr = (Time Mod 60) & " minutos, " & UpTimeStr
    Time = Time \ 60
    UpTimeStr = (Time Mod 24) & " horas, " & UpTimeStr
    Time = Time \ 24
    If Time = 1 Then
        UpTimeStr = Time & " día, " & UpTimeStr
    Else
        UpTimeStr = Time & " días, " & UpTimeStr
    End If
    Call WriteLocaleMsg(UserIndex, "1470", e_FontTypeNames.FONTTYPE_INFO)  ' Msg1470=Server Online: ¬1
    Exit Sub
HandleUpTime_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleUpTime", Erl)
End Sub

' Handles the "RoleMasterRequest" message.
Public Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim request As String
        request = reader.ReadString8()
        If LenB(request) <> 0 Then
            ' Msg527=Su solicitud ha sido enviada.
            Call WriteLocaleMsg(UserIndex, "527", e_FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageLocaleMsg(1825, .name & "¬" & request, e_FontTypeNames.FONTTYPE_GUILDMSG)) ' Msg1825=¬1 PREGUNTA ROL: ¬2
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRoleMasterRequest", Erl)
End Sub

Public Sub HandlePunishments(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim name As String
        name = reader.ReadString8()
        ' Si un GM usa este comando, me fijo que me haya dado el nick del PJ a analizar.
        If LenB(name) = 0 Then Exit Sub
        If UserList(UserIndex).name <> name Then
            If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
                ' Msg528=Servidor » Comando deshabilitado para tu cargo.
                Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        Dim count As Integer
        If (InStrB(name, "\") <> 0) Then
            name = Replace(name, "\", vbNullString)
        End If
        If (InStrB(name, "/") <> 0) Then
            name = Replace(name, "/", vbNullString)
        End If
        If (InStrB(name, ":") <> 0) Then
            name = Replace(name, ":", vbNullString)
        End If
        If (InStrB(name, "|") <> 0) Then
            name = Replace(name, "|", vbNullString)
        End If
        Dim TargetUserName As String
        If EsGM(UserIndex) Then
            If PersonajeExiste(name) Then
                TargetUserName = name
            Else
                Call WriteLocaleMsg(UserIndex, "1471", e_FontTypeNames.FONTTYPE_INFO)  ' Msg1471=El personaje ¬1 no existe.
                Exit Sub
            End If
        Else
            TargetUserName = .name
        End If
        count = GetUserAmountOfPunishmentsDatabase(TargetUserName)
        If count = 0 Then
            ' Msg529=Sin prontuario..
            Call WriteLocaleMsg(UserIndex, "529", e_FontTypeNames.FONTTYPE_INFO)
        Else
            Call SendUserPunishmentsDatabase(UserIndex, TargetUserName)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandlePunishments", Erl)
End Sub

Public Sub HandleGamble(ByVal UserIndex As Integer)
    On Error GoTo HandleGamble_Err
    With UserList(UserIndex)
        Dim amount As Integer
        amount = reader.ReadInt16()
        Dim NpcIndex  As Integer
        Dim charindex As Integer
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            Call WriteLocaleMsg(UserIndex, "530", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        NpcIndex = .flags.TargetNPC.ArrayIndex
        charindex = NpcList(NpcIndex).Char.charindex
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
        ElseIf Distancia(NpcList(NpcIndex).pos, .pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
        ElseIf NpcList(NpcIndex).npcType <> e_NPCType.Timbero Then
            Call WriteLocaleChatOverHead(UserIndex, 1322, vbNullString, charindex, vbWhite)
        ElseIf amount < 1 Then
            Call WriteLocaleChatOverHead(UserIndex, 1323, vbNullString, charindex, vbWhite)
        ElseIf amount > 5000 Then
            Call WriteLocaleChatOverHead(UserIndex, 1324, vbNullString, charindex, vbWhite)
        ElseIf .Stats.GLD < amount Then
            Call WriteLocaleChatOverHead(UserIndex, 1325, vbNullString, charindex, vbWhite)
        Else
            If RandomNumber(1, 100) <= 10 Then
                ' GANADOR
                .Stats.GLD = .Stats.GLD + amount
                Call WriteLocaleChatOverHead(UserIndex, GetWinnerMsgID(), PonerPuntos(amount), charindex, vbWhite)
                Apuestas.Perdidas = Apuestas.Perdidas + amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                ' PERDEDOR
                .Stats.GLD = .Stats.GLD - amount
                Call WriteLocaleChatOverHead(UserIndex, GetLoserMsgID(), PonerPuntos(amount), charindex, vbRed)
                Apuestas.Ganancias = Apuestas.Ganancias + amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
            End If
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            Call WriteUpdateGold(UserIndex)
        End If
    End With
    Exit Sub
HandleGamble_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGamble", Erl)
End Sub

Public Function GetWinnerMsgID() As Integer
    GetWinnerMsgID = 1328 + Int(Rnd * 4)
End Function

Public Function GetLoserMsgID() As Integer
    GetLoserMsgID = 1332 + Int(Rnd * 4)
End Function

Public Sub HandleMapPriceEntrance(ByVal UserIndex As Integer)
    On Error GoTo HandleMapPriceEntrance_Err
    With UserList(UserIndex)
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            Call WriteLocaleMsg(UserIndex, "530", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim entryPrice     As Integer
        Dim NpcIndex       As Integer
        Dim charindex      As Integer
        Dim arenaMap       As Integer
        Dim MapX           As Integer
        Dim MapY           As Integer
        Dim isArenaEnabled As Boolean
        NpcIndex = .flags.TargetNPC.ArrayIndex
        charindex = NpcList(NpcIndex).Char.charindex
        isArenaEnabled = NpcList(NpcIndex).flags.ArenaEnabled
        entryPrice = NpcList(NpcIndex).flags.MapEntryPrice
        arenaMap = NpcList(NpcIndex).flags.MapTargetEntry
        MapX = NpcList(NpcIndex).flags.MapTargetEntryX
        MapY = NpcList(NpcIndex).flags.MapTargetEntryY
        If Not isArenaEnabled Then Exit Sub
        If entryPrice = 0 Then Exit Sub
        If .flags.Muerto = 1 Then
            Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
        ElseIf Distancia(NpcList(NpcIndex).pos, .pos) > 10 Then
            Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
        ElseIf .Stats.GLD < entryPrice Then
            Call WriteLocaleChatOverHead(UserIndex, 2040, vbNullString, charindex, vbWhite)
        ElseIf NpcList(NpcIndex).npcType <> e_NPCType.ArenaGuard Then
            Call WriteLocaleChatOverHead(UserIndex, 1322, vbNullString, charindex, vbWhite)
        Else
            .Stats.GLD = .Stats.GLD - entryPrice
            Call WriteUpdateGold(UserIndex)
            Call WarpToLegalPos(UserIndex, arenaMap, MapX, MapY, True) 'Teleports user to the arena map
        End If
    End With
    Exit Sub
HandleMapPriceEntrance_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleMapPriceEntrance", Erl)
End Sub

Public Sub HandleDenounce(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim name As String
        name = reader.ReadString8()
        If LenB(name) = 0 Then Exit Sub
        If EsGmChar(name) Then
            ' Msg531=No podés denunciar a un administrador.
            Call WriteLocaleMsg(UserIndex, "531", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim tUser As t_UserReference
        tUser = NameIndex(name)
        If Not IsValidUserRef(tUser) Then 'user abuse of this system to know if someone is online
            ' Msg532=Tu denuncia fue recibida por el equipo de soporte.
            Call WriteLocaleMsg(UserIndex, "532", e_FontTypeNames.FONTTYPE_INFOIAO)
            Exit Sub
        End If
        Dim Denuncia As String, HayChat As Boolean
        Denuncia = "[Últimos mensajes de " & UserList(tUser.ArrayIndex).name & "]" & vbNewLine
        Dim i As Integer
        For i = 1 To UBound(UserList(tUser.ArrayIndex).flags.ChatHistory)
            If LenB(UserList(tUser.ArrayIndex).flags.ChatHistory(i)) <> 0 Then
                Denuncia = Denuncia & UserList(tUser.ArrayIndex).flags.ChatHistory(i) & vbNewLine
                HayChat = True
            End If
        Next
        If Not HayChat Then
            ' Msg533=El usuario no ha escrito nada. Recordá que las denuncias inválidas pueden ser motivo de advertencia.
            Call WriteLocaleMsg(UserIndex, "533", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call Ayuda.Push(.name, Denuncia, "Denuncia a " & UserList(tUser.ArrayIndex).name)
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageLocaleMsg(1826, .name, e_FontTypeNames.FONTTYPE_SERVER)) ' Msg1826=Se ha recibido una nueva denuncia de parte de ¬1.
        ' Msg534=Tu denuncia fue recibida por el equipo de soporte.
        Call WriteLocaleMsg(UserIndex, "534", e_FontTypeNames.FONTTYPE_INFOIAO)
        Call LogConsulta(.name & " (Denuncia a " & UserList(tUser.ArrayIndex).name & ")" & vbNewLine & Denuncia)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleDenounce", Erl)
End Sub

Public Sub HandleGMMessage(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Message As String
        Message = reader.ReadString8()
        If EsGM(UserIndex) Then
            Call LogGM(.name, "Mensaje a Gms: " & Message)
            If LenB(Message) <> 0 Then
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " » " & Message, e_FontTypeNames.FONTTYPE_GMMSG))
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGMMessage", Erl)
End Sub

Public Sub HandleShowName(ByVal UserIndex As Integer)
    On Error GoTo HandleShowName_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
            .showName = Not .showName 'Show / Hide the name
            Call RefreshCharStatus(UserIndex)
        End If
    End With
    Exit Sub
HandleShowName_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleShowName", Erl)
End Sub

Public Sub HandleGoNearby(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        username = reader.ReadString8()
        Dim tUser As t_UserReference
        Dim x     As Long
        Dim y     As Long
        Dim i     As Long
        Dim Found As Boolean
        If Not EsGM(UserIndex) Then Exit Sub
        'Check the user has enough powers
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Or Ayuda.Existe(username) Then
            tUser = NameIndex(username)
            If Not IsValidUserRef(tUser) Then
                ' Si está offline, comparamos privilegios offline, para no revelar si está el gm conectado
                If CompararPrivilegios(.flags.Privilegios, UserDarPrivilegioLevel(username)) >= 0 Then
                    ' Msg535=Usuario offline.
                    Call WriteLocaleMsg(UserIndex, "535", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    ' Msg536=No podés ir cerca de un GM de mayor jerarquía.
                    Call WriteLocaleMsg(UserIndex, "536", e_FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                If CompararPrivilegiosUser(UserIndex, tUser.ArrayIndex) >= 0 Then
                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For x = UserList(tUser.ArrayIndex).pos.x - i To UserList(tUser.ArrayIndex).pos.x + i
                            For y = UserList(tUser.ArrayIndex).pos.y - i To UserList(tUser.ArrayIndex).pos.y + i
                                If MapData(UserList(tUser.ArrayIndex).pos.Map, x, y).UserIndex = 0 Then
                                    If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) And MapInfo(UserList(tUser.ArrayIndex).pos.Map).Seguro = 0 Then
                                        ' Msg1319=No puedes ir en este momento al Usuario esta en zona insegura. Intenta mas tarde, puedes responderle con un mensaje.
                                        Call WriteLocaleMsg(UserIndex, "1319", e_FontTypeNames.FONTTYPE_INFO)
                                        Exit Sub
                                    End If
                                    If LegalPos(UserList(tUser.ArrayIndex).pos.Map, x, y, True, True) Then
                                        Call WriteLocaleMsg(UserIndex, "1472", e_FontTypeNames.FONTTYPE_INFO, UserList(tUser.ArrayIndex).name) ' Msg1472=Te teletransportaste cerca de ¬1.
                                        Call WarpUserChar(UserIndex, UserList(tUser.ArrayIndex).pos.Map, x, y, True)
                                        Found = True
                                        Exit For
                                    End If
                                End If
                            Next y
                            If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next x
                        If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next i
                    'No space found??
                    If Not Found Then
                        ' Msg537=Todos los lugares están ocupados.
                        Call WriteLocaleMsg(UserIndex, "537", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    ' Msg538=No podés ir cerca de un GM de mayor jerarquía.
                    Call WriteLocaleMsg(UserIndex, "538", e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        Else
            ' Msg539=Servidor » No podés ir cerca de ningun Usuario si no pidio SOS.
            Call WriteLocaleMsg(UserIndex, "539", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGoNearby", Erl)
End Sub

Public Sub HandleWhere(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.RoleMaster Or e_PlayerType.Consejero Or e_PlayerType.User Or e_PlayerType.SemiDios)) = 0 Then
            tUser = NameIndex(username)
            If Not IsValidUserRef(tUser) Then
                ' Msg540=Usuario offline.
                Call WriteLocaleMsg(UserIndex, "540", e_FontTypeNames.FONTTYPE_INFO)
            Else
                If CompararPrivilegiosUser(UserIndex, tUser.ArrayIndex) >= 0 Then
                    Call WriteLocaleMsg(UserIndex, "1090", e_FontTypeNames.FONTTYPE_FIGHT, username & ": " & UserList(tUser.ArrayIndex).pos.Map & ", " & UserList( _
                            tUser.ArrayIndex).pos.x & ", " & UserList(tUser.ArrayIndex).pos.y & ".")
                    Call LogGM(.name, "/Donde " & username)
                End If
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleWhere", Erl)
End Sub

Public Sub HandleCreaturesInMap(ByVal UserIndex As Integer)
    On Error GoTo HandleCreaturesInMap_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        Dim Map As Integer
        Dim i, j As Long
        Dim NPCcount1, NPCcount2 As Integer
        Dim NPCcant1() As Integer
        Dim NPCcant2() As Integer
        Dim List1()    As String
        Dim List2()    As String
        Map = reader.ReadInt16()
        If .flags.Privilegios And e_PlayerType.User Then Exit Sub
        If MapaValido(Map) Then
            For i = 1 To LastNPC
                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If NpcList(i).pos.Map = Map Then
                    'íesta vivo?
                    If NpcList(i).flags.NPCActive And NpcList(i).Hostile = 1 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = NpcList(i).name & ": (" & NpcList(i).pos.x & "," & NpcList(i).pos.y & ")"
                            NPCcant1(0) = 1
                        Else
                            For j = 0 To NPCcount1 - 1
                                If Left$(List1(j), Len(NpcList(i).name)) = NpcList(i).name Then
                                    List1(j) = List1(j) & ", (" & NpcList(i).pos.x & "," & NpcList(i).pos.y & ")"
                                    NPCcant1(j) = NPCcant1(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(j) = NpcList(i).name & ": (" & NpcList(i).pos.x & "," & NpcList(i).pos.y & ")"
                                NPCcant1(j) = 1
                            End If
                        End If
                    Else
                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = NpcList(i).name & ": (" & NpcList(i).pos.x & "," & NpcList(i).pos.y & ")"
                            NPCcant2(0) = 1
                        Else
                            For j = 0 To NPCcount2 - 1
                                If Left$(List2(j), Len(NpcList(i).name)) = NpcList(i).name Then
                                    List2(j) = List2(j) & ", (" & NpcList(i).pos.x & "," & NpcList(i).pos.y & ")"
                                    NPCcant2(j) = NPCcant2(j) + 1
                                    Exit For
                                End If
                            Next j
                            If j = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(j) = NpcList(i).name & ": (" & NpcList(i).pos.x & "," & NpcList(i).pos.y & ")"
                                NPCcant2(j) = 1
                            End If
                        End If
                    End If
                End If
            Next i
            ' Msg736=Npcs Hostiles en mapa:
            Call WriteLocaleMsg(UserIndex, "736", e_FontTypeNames.FONTTYPE_WARNING)
            If NPCcount1 = 0 Then
                ' Msg737=No hay NPCS Hostiles
                Call WriteLocaleMsg(UserIndex, "737", e_FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), e_FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            ' Msg738=Otros Npcs en mapa:
            Call WriteLocaleMsg(UserIndex, "738", e_FontTypeNames.FONTTYPE_WARNING)
            If NPCcount2 = 0 Then
                ' Msg739=No hay más NPCS
                Call WriteLocaleMsg(UserIndex, "739", e_FontTypeNames.FONTTYPE_INFO)
            Else
                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), e_FontTypeNames.FONTTYPE_INFO)
                Next j
            End If
            Call LogGM(.name, "Numero enemigos en mapa " & Map)
        End If
    End With
    Exit Sub
HandleCreaturesInMap_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreaturesInMap", Erl)
End Sub

Public Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)
    On Error GoTo HandleWarpMeToTarget_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        If .flags.Privilegios And e_PlayerType.User Then Exit Sub
        Call WarpUserChar(UserIndex, .flags.TargetMap, .flags.TargetX, .flags.TargetY, True)
        Call LogGM(.name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .pos.Map)
    End With
    Exit Sub
HandleWarpMeToTarget_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleWarpMeToTarget", Erl)
End Sub

Public Sub HandleWarpChar(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim Map      As Integer
        Dim x        As Byte
        Dim y        As Byte
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        Map = reader.ReadInt16()
        x = reader.ReadInt8()
        y = reader.ReadInt8()
        If .flags.Privilegios And e_PlayerType.User Then Exit Sub
        If .flags.Privilegios And e_PlayerType.Consejero Then
            If MapInfo(Map).Seguro = 0 Then
                ' Msg741=Solo puedes transportarte a ciudades.
                Call WriteLocaleMsg(UserIndex, "741", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
                'Si manda yo o su propio nombre
            ElseIf LCase$(username) <> LCase$(UserList(UserIndex).name) And UCase$(username) <> "YO" Then
                ' Msg742=Solo puedes transportarte a ti mismo.
                Call WriteLocaleMsg(UserIndex, "742", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        '¿Para que te vas a transportar a la misma posicion?
        If .pos.Map = Map And .pos.x = x And .pos.y = y Then Exit Sub
        If MapaValido(Map) And LenB(username) <> 0 Then
            If UCase$(username) <> "YO" Then
                tUser = NameIndex(username)
            Else
                Call SetUserRef(tUser, UserIndex)
            End If
            If Not IsValidUserRef(tUser) Then
                ' Msg743=Usuario offline.
                Call WriteLocaleMsg(UserIndex, "743", e_FontTypeNames.FONTTYPE_INFO)
            ElseIf InMapBounds(Map, x, y) Then
                'no permitimos que se use el telep para llevas User a casas privadas.
                If UCase$(username) <> "YO" Then
                    If .flags.Privilegios And e_PlayerType.Consejero Or e_PlayerType.SemiDios Then
                        If Not EsMapaEvento(Map) Then
                            'Msg948= Solamente puedes teletransportar gente a mapas de evento.
                            Call WriteLocaleMsg(UserIndex, "948", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                    End If
                    Call WarpToLegalPos(tUser.ArrayIndex, Map, x, y, True, True)
                Else
                    Call FindLegalPos(tUser.ArrayIndex, Map, x, y)
                    Call WarpUserChar(tUser.ArrayIndex, Map, x, y, True)
                End If
                If tUser.ArrayIndex <> UserIndex Then
                    Call LogGM(.name, "Transportó a " & UserList(tUser.ArrayIndex).name & " hacia " & "Mapa" & Map & " X:" & x & " Y:" & y)
                End If
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleWarpChar", Erl)
End Sub

Public Sub HandleSilence(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim minutos  As Integer
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        minutos = reader.ReadInt16()
        If EsGM(UserIndex) Then
            tUser = NameIndex(username)
            If Not IsValidUserRef(tUser) Then
                If PersonajeExiste(username) Then
                    If CompararPrivilegios(.flags.Privilegios, UserDarPrivilegioLevel(username)) > 0 Then
                        If minutos > 0 Then
                            Call SilenciarUserDatabase(username, minutos)
                            Call SavePenaDatabase(username, .name & ": silencio por " & Time & " minutos. " & Date & " " & Time)
                            Call SendData(SendTarget.ToGM, 0, PrepareMessageLocaleMsg(1827, .name & "¬" & username & "¬" & minutos, e_FontTypeNames.FONTTYPE_GM)) ' Msg1827=Administración » ¬1 ha silenciado a ¬2 (offline) por ¬3 minutos.
                            Call LogGM(.name, "Silenciar a " & UserList(tUser.ArrayIndex).name & " por " & minutos & " minutos.")
                        Else
                            Call DesilenciarUserDatabase(username)
                            Call SendData(SendTarget.ToGM, 0, PrepareMessageLocaleMsg(1828, .name & "¬" & username, e_FontTypeNames.FONTTYPE_GM)) ' Msg1828=Administración » ¬1 ha desilenciado a ¬2 (offline).
                            Call LogGM(.name, "Desilenciar a " & UserList(tUser.ArrayIndex).name & ".")
                        End If
                    Else
                        ' Msg561=No puedes silenciar a un administrador de mayor o igual rango.
                        Call WriteLocaleMsg(UserIndex, "561", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    ' Msg560=El personaje no existe.
                    Call WriteLocaleMsg(UserIndex, "560", e_FontTypeNames.FONTTYPE_INFO)
                End If
            ElseIf CompararPrivilegiosUser(UserIndex, tUser.ArrayIndex) > 0 Then
                If minutos > 0 Then
                    UserList(tUser.ArrayIndex).flags.Silenciado = 1
                    UserList(tUser.ArrayIndex).flags.MinutosRestantes = minutos
                    UserList(tUser.ArrayIndex).flags.SegundosPasados = 0
                    Call SavePenaDatabase(username, .name & ": silencio por " & Time & " minutos. " & Date & " " & Time)
                    Call SendData(SendTarget.ToGM, 0, PrepareMessageLocaleMsg(1829, .name & "¬" & UserList(tUser.ArrayIndex).name & "¬" & minutos, e_FontTypeNames.FONTTYPE_GM)) ' Msg1829=Administración » ¬1 ha silenciado a ¬2 por ¬3 minutos.
                    'Msg949= Has sido silenciado por los administradores, no podrás hablar con otros usuarios. Utilice /GM para pedir ayuda.
                    Call WriteLocaleMsg(tUser.ArrayIndex, "949", e_FontTypeNames.FONTTYPE_GM)
                    Call LogGM(.name, "Silenciar a " & UserList(tUser.ArrayIndex).name & " por " & minutos & " minutos.")
                Else
                    UserList(tUser.ArrayIndex).flags.Silenciado = 1
                    Call SendData(SendTarget.ToGM, 0, PrepareMessageLocaleMsg(1830, .name & "¬" & UserList(tUser.ArrayIndex).name, e_FontTypeNames.FONTTYPE_GM)) ' Msg1830=Administración » ¬1 ha desilenciado a ¬2.
                    'Msg950= Has sido desilenciado.
                    Call WriteLocaleMsg(tUser.ArrayIndex, "950", e_FontTypeNames.FONTTYPE_GM)
                    Call LogGM(.name, "Desilenciar a " & UserList(tUser.ArrayIndex).name & ".")
                End If
            Else
                ' Msg561=No puedes silenciar a un administrador de mayor o igual rango.
                Call WriteLocaleMsg(UserIndex, "561", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSilence", Erl)
End Sub

Public Sub HandleSOSShowList(ByVal UserIndex As Integer)
    On Error GoTo HandleSOSShowList_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        If .flags.Privilegios And e_PlayerType.User Then Exit Sub
        Call WriteShowSOSForm(UserIndex)
    End With
    Exit Sub
HandleSOSShowList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSOSShowList", Erl)
End Sub

Public Sub HandleSOSRemove(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        username = reader.ReadString8()
        If Not .flags.Privilegios And e_PlayerType.User Then Call Ayuda.Quitar(username)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSOSRemove", Erl)
End Sub

Public Sub HandleGoToChar(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        Dim x        As Byte
        Dim y        As Byte
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            If LenB(username) <> 0 Then
                tUser = NameIndex(username)
                If Not IsValidUserRef(tUser) Then
                    ' Msg562=El jugador no está online.
                    Call WriteLocaleMsg(UserIndex, "562", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Else
                tUser = .flags.TargetUser
                If Not IsValidUserRef(tUser) Then Exit Sub
            End If
            If CompararPrivilegiosUser(tUser.ArrayIndex, UserIndex) > 0 Then
                Call WriteLocaleMsg(UserIndex, "1473", e_FontTypeNames.FONTTYPE_INFO)  ' Msg1473=Se le ha avisado a ¬1 que quieres ir a su posición.
                Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1850, .name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1850=¬1 quiere transportarse a tu ubicación. Escribe /sum ¬1 para traerlo.
                Exit Sub
            End If
            x = UserList(tUser.ArrayIndex).pos.x
            y = UserList(tUser.ArrayIndex).pos.y + 1
            Call FindLegalPos(UserIndex, UserList(tUser.ArrayIndex).pos.Map, x, y)
            Call WarpUserChar(UserIndex, UserList(tUser.ArrayIndex).pos.Map, x, y, True)
            If .flags.AdminInvisible = 0 Then
                Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1851, .name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1851=¬1 se ha transportado hacia donde te encuentras.
            End If
            Call WriteLocaleMsg(UserIndex, "1474", e_FontTypeNames.FONTTYPE_INFO, UserList(tUser.ArrayIndex).name) ' Msg1474=Te has transportado hacia ¬1.
            Call LogGM(.name, "/IRA " & username & " Mapa:" & UserList(tUser.ArrayIndex).pos.Map & " X:" & UserList(tUser.ArrayIndex).pos.x & " Y:" & UserList( _
                    tUser.ArrayIndex).pos.y)
        Else
            'Msg951= Servidor » Comando deshabilitado para tu cargo. solo puedes ir a Usuarios que piden SOS.
            Call WriteLocaleMsg(UserIndex, "951", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGoToChar", Erl)
End Sub

Public Sub HandleInvisible(ByVal UserIndex As Integer)
    On Error GoTo HandleInvisible_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        If .flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero) Then Exit Sub
        Call DoAdminInvisible(UserIndex)
    End With
    Exit Sub
HandleInvisible_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleInvisible", Erl)
End Sub

Public Sub HandleGMPanel(ByVal UserIndex As Integer)
    On Error GoTo HandleGMPanel_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        If .flags.Privilegios And e_PlayerType.User Then Exit Sub
        Call WriteShowGMPanelForm(UserIndex)
    End With
    Exit Sub
HandleGMPanel_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGMPanel", Erl)
End Sub

Public Sub HandleRequestUserList(ByVal UserIndex As Integer)
    On Error GoTo HandleRequestUserList_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/09/07
    'Last modified by: Lucas Tavolaro Ortiz (Tavo)
    Dim i       As Long
    Dim names() As String
    Dim count   As Long
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.RoleMaster)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        ReDim names(1 To LastUser) As String
        count = 1
        For i = 1 To LastUser
            If (LenB(UserList(i).name) <> 0) Then
                names(count) = UserList(i).name
                count = count + 1
            End If
        Next i
        If count > 1 Then Call WriteUserNameList(UserIndex, names(), count - 1)
    End With
    Exit Sub
HandleRequestUserList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestUserList", Erl)
End Sub

Public Sub HandleWorking(ByVal UserIndex As Integer)
    On Error GoTo HandleWorking_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    Dim i     As Long
    Dim Users As String
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster Or e_PlayerType.SemiDios)) Then
            'Msg952= Servidor » /TRABAJANDO es un comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "952", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                Users = Users & ", " & UserList(i).name
            End If
        Next i
        If LenB(Users) <> 0 Then
            Users = Right$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & Users, e_FontTypeNames.FONTTYPE_INFO)
        Else
            'Msg953= No hay usuarios trabajando
            Call WriteLocaleMsg(UserIndex, "953", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleWorking_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleWorking", Erl)
End Sub

Public Sub HandleHiding(ByVal UserIndex As Integer)
    On Error GoTo HandleHiding_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    Dim i     As Long
    Dim Users As String
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        For i = 1 To LastUser
            If (LenB(UserList(i).name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                Users = Users & UserList(i).name & ", "
            End If
        Next i
        If LenB(Users) <> 0 Then
            Users = Left$(Users, Len(Users) - 2)
            Call WriteLocaleMsg(UserIndex, "1476", e_FontTypeNames.FONTTYPE_INFO, Users)   ' Msg1476=Usuarios ocultandose: ¬1
        Else
            'Msg954= No hay usuarios ocultandose
            Call WriteLocaleMsg(UserIndex, "954", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleHiding_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleHiding", Erl)
End Sub

Public Sub HandleJail(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim Reason   As String
        Dim jailTime As Integer
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        Reason = reader.ReadString8()
        jailTime = reader.ReadInt16()
        If InStr(1, username, "+") Then
            username = Replace(username, "+", " ")
        End If
        '/carcel nick@motivo@<tiempo>
        If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios) Then
            If LenB(username) = 0 Or LenB(Reason) = 0 Then
                'Msg955= Utilice /carcel nick@motivo@tiempo
                Call WriteLocaleMsg(UserIndex, "955", e_FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(username)
                If Not IsValidUserRef(tUser) Then
                    'Msg956= El usuario no está online.
                    Call WriteLocaleMsg(UserIndex, "956", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    If EsGM(tUser.ArrayIndex) Then
                        'Msg957= No podés encarcelar a administradores.
                        Call WriteLocaleMsg(UserIndex, "957", e_FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > SvrConfig.GetValue("MaxJailTime") Then
                        'Msg958= No podés encarcelar por más de ¬1 minutos.
                        Call WriteLocaleMsg(UserIndex, "958", e_FontTypeNames.FONTTYPE_INFO, SvrConfig.GetValue("MaxJailTime"))
                    Else
                        If (InStrB(username, "\") <> 0) Then
                            username = Replace(username, "\", "")
                        End If
                        If (InStrB(username, "/") <> 0) Then
                            username = Replace(username, "/", "")
                        End If
                        If PersonajeExiste(username) Then
                            Call SavePenaDatabase(username, .name & ": CARCEL " & jailTime & "m, MOTIVO: " & Reason & " " & Date & " " & Time)
                        End If
                        Call Encarcelar(tUser.ArrayIndex, jailTime, .name)
                        Call LogGM(.name, " encarceló a " & username)
                    End If
                End If
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleHiding", Erl)
End Sub

Public Sub HandleKillNPC(ByVal UserIndex As Integer)
    On Error GoTo HandleKillNPC_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        If .flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero) Then
            'Msg959= Solo los Administradores y Dioses pueden usar este comando.
            Call WriteLocaleMsg(UserIndex, "959", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If IsValidNpcRef(.flags.TargetNPC) Then
            Dim tNPC As Integer
            tNPC = .flags.TargetNPC.ArrayIndex
            Call WriteLocaleMsg(UserIndex, "1477", e_FontTypeNames.FONTTYPE_INFO, tNPC) ' Msg1477=RMatas (con posible respawn) a: ¬1
            Dim auxNPC As t_Npc
            auxNPC = NpcList(tNPC)
            Call QuitarNPC(tNPC, eGMCommand)
            Call ReSpawnNpc(auxNPC)
        Else
            'Msg960= Debes hacer click sobre el NPC antes
            Call WriteLocaleMsg(UserIndex, "960", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleKillNPC_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleKillNPC", Erl)
End Sub

Public Sub HandleWarnUser(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim Reason   As String
        username = reader.ReadString8()
        Reason = reader.ReadString8()
        ' Tenes que ser Admin, Dios o Semi-Dios
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) = 0 Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        ' Me fijo que esten todos los parametros.
        If Len(username) = 0 Or Len(Trim$(Reason)) = 0 Then
            'Msg961= Formato inválido. /advertencia nick@motivo
            Call WriteLocaleMsg(UserIndex, "961", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim tUser As t_UserReference
        tUser = NameIndex(username)
        ' No advertir a GM's
        If EsGM(tUser.ArrayIndex) Then
            ' Msg541=No podes advertir a Game Masters.
            Call WriteLocaleMsg(UserIndex, "541", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If (InStrB(username, "\") <> 0) Then
            username = Replace(username, "\", "")
        End If
        If (InStrB(username, "/") <> 0) Then
            username = Replace(username, "/", "")
        End If
        If PersonajeExiste(username) Then
            Call SaveWarnDatabase(username, "ADVERTENCIA: " & Reason & " " & Date & " " & Time, .name)
            ' Para el GM
            Call WriteLocaleMsg(UserIndex, "1478", e_FontTypeNames.FONTTYPE_CENTINELA, username) ' Msg1478=Has advertido a ¬1
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageLocaleMsg(1831, .name & "¬" & username & "¬" & Reason, e_FontTypeNames.FONTTYPE_GM)) ' Msg1831=¬1 ha advertido a ¬2 por ¬3.
            Call LogGM(.name, " advirtio a " & username & " por " & Reason)
            ' Si esta online...
            If IsValidUserRef(tUser) Then
                ' Actualizo el valor en la memoria.
                UserList(tUser.ArrayIndex).Stats.Advertencias = UserList(tUser.ArrayIndex).Stats.Advertencias + 1
                ' Para el usuario advertido
                Call WriteLocaleMsg(tUser.ArrayIndex, "1479", e_FontTypeNames.FONTTYPE_CENTINELA, .name)  ' Msg1479=Has sido advertido por ¬1
                Call WriteLocaleMsg(tUser.ArrayIndex, "1480", e_FontTypeNames.FONTTYPE_CENTINELA, UserList(tUser.ArrayIndex).Stats.Advertencias)   ' Msg1480=Tenés ¬1 advertencias actualmente.
                ' Cuando acumulas cierta cantidad de advertencias...
                Select Case UserList(tUser.ArrayIndex).Stats.Advertencias
                    Case 3
                        Call Encarcelar(tUser.ArrayIndex, 30, "Servidor")
                    Case 5
                        ' TODO: Banear PJ alv.
                End Select
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleWarnUser", Erl)
End Sub

Public Sub HandleEditChar(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username      As String
        Dim tUser         As t_UserReference
        Dim opcion        As Byte
        Dim Arg1          As String
        Dim Arg2          As String
        Dim LoopC         As Byte
        Dim commandString As String
        Dim tmpLong       As Long
        username = Replace(reader.ReadString8(), "+", " ")
        If UCase$(username) = "YO" Then
            Call SetUserRef(tUser, UserIndex)
        Else
            tUser = NameIndex(username)
        End If
        opcion = reader.ReadInt8()
        Arg1 = reader.ReadString8()
        Arg2 = reader.ReadString8()
        ' Si no es GM, no hacemos nada.
        If Not EsGM(UserIndex) Then Exit Sub
        ' From now on we disabled character edition in production environment
        ' This should have never been enabled in the first place
        ' If we require similar functionality for adhoc quests we will have to
        ' create new in-game tools focusing in TRANSPARENCY to prevent future abuse
        #If DEBUGGING = 0 Then
            Exit Sub
        #End If
        ' Si NO sos Dios o Admin,
        If (.flags.Privilegios And e_PlayerType.Admin) = 0 Then
            ' Si te editas a vos mismo esta bien ;)
            If UserIndex <> tUser.ArrayIndex Then Exit Sub
        End If
        If Not IsValidUserRef(tUser) Then
            Call WriteLocaleMsg(UserIndex, "1481", e_FontTypeNames.FONTTYPE_INFO, username) ' Msg1481=Usuario offline: ¬1
            Exit Sub
        End If
        Select Case opcion
            Case e_EditOptions.eo_Gold
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                UserList(tUser.ArrayIndex).Stats.GLD = val(Arg1)
                Call WriteUpdateGold(tUser.ArrayIndex)
            Case e_EditOptions.eo_Experience
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                If UserList(tUser.ArrayIndex).Stats.ELV < STAT_MAXELV Then
                    UserList(tUser.ArrayIndex).Stats.Exp = UserList(tUser.ArrayIndex).Stats.Exp + val(Arg1)
                    Call CheckUserLevel(tUser.ArrayIndex)
                    Call WriteUpdateExp(tUser.ArrayIndex)
                Else
                    ' Msg542=El usuario es nivel máximo.
                    Call WriteLocaleMsg(UserIndex, "542", e_FontTypeNames.FONTTYPE_INFO)
                End If
            Case e_EditOptions.eo_Body
                If Not IsValidUserRef(tUser) Then
                    Call SaveUserBodyDatabase(username, val(Arg1))
                    Call WriteLocaleMsg(UserIndex, "1482", e_FontTypeNames.FONTTYPE_INFO, username)  ' Msg1482=Usuario Offline Alterado: ¬1
                Else
                    Call ChangeUserChar(tUser.ArrayIndex, val(Arg1), UserList(tUser.ArrayIndex).Char.head, UserList(tUser.ArrayIndex).Char.Heading, UserList( _
                            tUser.ArrayIndex).Char.WeaponAnim, UserList(tUser.ArrayIndex).Char.ShieldAnim, UserList(tUser.ArrayIndex).Char.CascoAnim, UserList( _
                            UserIndex).Char.CartAnim, UserList(UserIndex).Char.BackpackAnim)
                End If
            Case e_EditOptions.eo_Arma
                Call ChangeUserChar(tUser.ArrayIndex, UserList(tUser.ArrayIndex).Char.body, UserList(tUser.ArrayIndex).Char.head, UserList(tUser.ArrayIndex).Char.Heading, val( _
                        Arg1), UserList(tUser.ArrayIndex).Char.ShieldAnim, UserList(tUser.ArrayIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList( _
                        UserIndex).Char.CartAnim)
            Case e_EditOptions.eo_Escudo
                Call ChangeUserChar(tUser.ArrayIndex, UserList(tUser.ArrayIndex).Char.body, UserList(tUser.ArrayIndex).Char.head, UserList(tUser.ArrayIndex).Char.Heading, _
                        UserList(tUser.ArrayIndex).Char.WeaponAnim, val(Arg1), UserList(tUser.ArrayIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim, UserList( _
                        UserIndex).Char.CartAnim)
            Case e_EditOptions.eo_CASCO
                Call ChangeUserChar(tUser.ArrayIndex, UserList(tUser.ArrayIndex).Char.body, UserList(tUser.ArrayIndex).Char.head, UserList(tUser.ArrayIndex).Char.Heading, _
                        UserList(tUser.ArrayIndex).Char.WeaponAnim, UserList(tUser.ArrayIndex).Char.ShieldAnim, val(Arg1), UserList(UserIndex).Char.CartAnim, UserList( _
                        UserIndex).Char.CartAnim)
            Case e_EditOptions.eo_Particula
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                If Not .flags.Privilegios = Consejero Then
                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, val(Arg1), 9999, False))
                    .Char.ParticulaFx = val(Arg1)
                    .Char.loops = 9999
                End If
            Case e_EditOptions.eo_Head
                If tUser.ArrayIndex <= 0 Then
                    Call SaveUserHeadDatabase(username, val(Arg1))
                    Call WriteLocaleMsg(UserIndex, "1483", e_FontTypeNames.FONTTYPE_INFO, username) ' Msg1483=Usuario Offline Alterado: ¬1
                Else
                    Call ChangeUserChar(tUser.ArrayIndex, UserList(tUser.ArrayIndex).Char.body, val(Arg1), UserList(tUser.ArrayIndex).Char.Heading, UserList( _
                            tUser.ArrayIndex).Char.WeaponAnim, UserList(tUser.ArrayIndex).Char.ShieldAnim, UserList(tUser.ArrayIndex).Char.CascoAnim, UserList( _
                            UserIndex).Char.CartAnim, UserList(UserIndex).Char.CartAnim)
                End If
            Case e_EditOptions.eo_CriminalsKilled
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                If val(Arg1) > MAXUSERMATADOS Then
                    UserList(tUser.ArrayIndex).Faccion.CriminalesMatados = MAXUSERMATADOS
                Else
                    UserList(tUser.ArrayIndex).Faccion.CriminalesMatados = val(Arg1)
                End If
            Case e_EditOptions.eo_CiticensKilled
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                If val(Arg1) > MAXUSERMATADOS Then
                    UserList(tUser.ArrayIndex).Faccion.ciudadanosMatados = MAXUSERMATADOS
                Else
                    UserList(tUser.ArrayIndex).Faccion.ciudadanosMatados = val(Arg1)
                End If
            Case e_EditOptions.eo_Level
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero)) Then Exit Sub
                If val(Arg1) > STAT_MAXELV Then
                    Arg1 = CStr(STAT_MAXELV)
                    Call WriteLocaleMsg(UserIndex, "1484", FONTTYPE_INFO, STAT_MAXELV)  ' Msg1484=No podés tener un nivel superior a ¬1.
                End If
                UserList(tUser.ArrayIndex).Stats.ELV = val(Arg1)
                Call WriteUpdateUserStats(UserIndex)
            Case e_EditOptions.eo_Class
                For LoopC = 1 To NUMCLASES
                    If Tilde(ListaClases(LoopC)) = Tilde(Arg1) Then Exit For
                Next LoopC
                If LoopC > NUMCLASES Then
                    ' Msg543=Clase desconocida. Intente nuevamente.
                    Call WriteLocaleMsg(UserIndex, "543", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    UserList(tUser.ArrayIndex).clase = LoopC
                End If
            Case e_EditOptions.eo_Skills
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero)) Then Exit Sub
                For LoopC = 1 To NUMSKILLS
                    If Tilde(Replace$(SkillsNames(LoopC), " ", "+")) = Tilde(Arg1) Then Exit For
                Next LoopC
                If LoopC > NUMSKILLS Then
                    ' Msg544=Skill Inexistente!
                    Call WriteLocaleMsg(UserIndex, "544", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    If Not IsValidUserRef(tUser) Then
                        Call SaveUserSkillDatabase(username, LoopC, val(Arg2))
                        Call WriteLocaleMsg(UserIndex, "1485", e_FontTypeNames.FONTTYPE_INFO, username) ' Msg1485=Usuario Offline Alterado: ¬1
                    Else
                        UserList(tUser.ArrayIndex).Stats.UserSkills(LoopC) = val(Arg2)
                    End If
                End If
            Case e_EditOptions.eo_SkillPointsLeft
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero)) Then Exit Sub
                If Not IsValidUserRef(tUser) Then
                    Call SaveUserSkillsLibres(username, val(Arg1))
                    Call WriteLocaleMsg(UserIndex, "1486", e_FontTypeNames.FONTTYPE_INFO, username) ' Msg1486=Usuario Offline Alterado: ¬1
                Else
                    UserList(tUser.ArrayIndex).Stats.SkillPts = val(Arg1)
                End If
            Case e_EditOptions.eo_Sex
                If Not IsValidUserRef(tUser) Then
                    Call WriteLocaleMsg(UserIndex, "1487", e_FontTypeNames.FONTTYPE_INFO, username) ' Msg1487=Usuario offline: ¬1
                Else
                    Arg1 = UCase$(Arg1)
                    If (Arg1 = "MUJER") Then
                        UserList(tUser.ArrayIndex).genero = e_Genero.Mujer
                    ElseIf (Arg1 = "HOMBRE") Then
                        UserList(tUser.ArrayIndex).genero = e_Genero.Hombre
                    End If
                End If
            Case e_EditOptions.eo_Raza
                If Not IsValidUserRef(tUser) Then
                    Call WriteLocaleMsg(UserIndex, "1488", e_FontTypeNames.FONTTYPE_INFO, username) ' Msg1488=Usuario offline: ¬1
                Else
                    Arg1 = UCase$(Arg1)
                    If (Arg1 = "HUMANO") Then
                        UserList(tUser.ArrayIndex).raza = e_Raza.Humano
                    ElseIf (Arg1 = "ELFO") Then
                        UserList(tUser.ArrayIndex).raza = e_Raza.Elfo
                    ElseIf (Arg1 = "DROW") Then
                        UserList(tUser.ArrayIndex).raza = e_Raza.Drow
                    ElseIf (Arg1 = "ENANO") Then
                        UserList(tUser.ArrayIndex).raza = e_Raza.Enano
                    ElseIf (Arg1 = "GNOMO") Then
                        UserList(tUser.ArrayIndex).raza = e_Raza.Gnomo
                    ElseIf (Arg1 = "ORCO") Then
                        UserList(tUser.ArrayIndex).raza = e_Raza.Orco
                    End If
                End If
            Case e_EditOptions.eo_Vida
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                If Not IsValidUserRef(tUser) Then
                    Call WriteLocaleMsg(UserIndex, "1489", e_FontTypeNames.FONTTYPE_INFO, username) ' Msg1489=Usuario offline: ¬1
                Else
                    tmpLong = val(Arg1)
                    If tmpLong > 0 Then
                        UserList(tUser.ArrayIndex).Stats.MaxHp = Min(tmpLong, STAT_MAXHP)
                        UserList(tUser.ArrayIndex).Stats.MinHp = UserList(tUser.ArrayIndex).Stats.MaxHp
                        Call WriteUpdateUserStats(tUser.ArrayIndex)
                    End If
                End If
            Case e_EditOptions.eo_Mana
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                tmpLong = val(Arg1)
                If tmpLong > 0 Then
                    UserList(tUser.ArrayIndex).Stats.MaxMAN = Min(tmpLong, STAT_MAXMP)
                    UserList(tUser.ArrayIndex).Stats.MinMAN = UserList(tUser.ArrayIndex).Stats.MaxMAN
                    Call WriteUpdateUserStats(tUser.ArrayIndex)
                End If
            Case e_EditOptions.eo_Energia
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                tmpLong = val(Arg1)
                If tmpLong > 0 Then
                    UserList(tUser.ArrayIndex).Stats.MaxSta = Min(tmpLong, STAT_MAXSTA)
                    UserList(tUser.ArrayIndex).Stats.MinSta = UserList(tUser.ArrayIndex).Stats.MaxSta
                    Call WriteUpdateUserStats(tUser.ArrayIndex)
                End If
            Case e_EditOptions.eo_MinHP
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                tmpLong = val(Arg1)
                If tmpLong >= 0 Then
                    UserList(tUser.ArrayIndex).Stats.MinHp = Min(tmpLong, STAT_MAXHP)
                    Call WriteUpdateHP(tUser.ArrayIndex)
                End If
            Case e_EditOptions.eo_MinMP
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                tmpLong = val(Arg1)
                If tmpLong >= 0 Then
                    UserList(tUser.ArrayIndex).Stats.MinMAN = Min(tmpLong, STAT_MAXMP)
                    Call WriteUpdateMana(tUser.ArrayIndex)
                End If
            Case e_EditOptions.eo_Hit
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                tmpLong = val(Arg1)
                If tmpLong >= 0 Then
                    UserList(tUser.ArrayIndex).Stats.MaxHit = Min(tmpLong, STAT_MAXHIT)
                    UserList(tUser.ArrayIndex).Stats.MinHIT = UserList(tUser.ArrayIndex).Stats.MaxHit
                End If
            Case e_EditOptions.eo_MinHit
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                tmpLong = val(Arg1)
                If tmpLong >= 0 Then
                    UserList(tUser.ArrayIndex).Stats.MinHIT = Min(tmpLong, STAT_MAXHIT)
                End If
            Case e_EditOptions.eo_MaxHit
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                tmpLong = val(Arg1)
                If tmpLong >= 0 Then
                    UserList(tUser.ArrayIndex).Stats.MaxHit = Min(tmpLong, STAT_MAXHIT)
                End If
            Case e_EditOptions.eo_Desc
                If ValidDescription(Arg1) Then
                    UserList(tUser.ArrayIndex).Desc = Arg1
                Else
                    ' Msg545=Caracteres inválidos en la descripción.
                    Call WriteLocaleMsg(UserIndex, "545", e_FontTypeNames.FONTTYPE_INFO)
                End If
            Case e_EditOptions.eo_Intervalo
                If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
                Arg1 = UCase$(Arg1)
                tmpLong = val(Arg2)
                If tmpLong >= 0 Then
                    Select Case Arg1
                        Case "USAR"
                            UserList(tUser.ArrayIndex).Intervals.UsarClic = tmpLong
                            UserList(tUser.ArrayIndex).Intervals.UsarU = tmpLong
                        Case "USAR_U", "USAR+U", "USAR-U"
                            UserList(tUser.ArrayIndex).Intervals.UsarU = tmpLong
                        Case "USAR_CLIC", "USAR+CLIC", "USAR-CLIC", "USAR_CLICK", "USAR+CLICK", "USAR-CLICK"
                            UserList(tUser.ArrayIndex).Intervals.UsarClic = tmpLong
                        Case "ARCO", "PROYECTILES"
                            UserList(tUser.ArrayIndex).Intervals.Arco = tmpLong
                        Case "GOLPE", "GOLPES", "GOLPEAR"
                            UserList(tUser.ArrayIndex).Intervals.Golpe = tmpLong
                        Case "MAGIA", "HECHIZO", "HECHIZOS", "LANZAR"
                            UserList(tUser.ArrayIndex).Intervals.Magia = tmpLong
                        Case "COMBO"
                            UserList(tUser.ArrayIndex).Intervals.GolpeMagia = tmpLong
                            UserList(tUser.ArrayIndex).Intervals.MagiaGolpe = tmpLong
                        Case "GOLPE-MAGIA", "GOLPE-HECHIZO"
                            UserList(tUser.ArrayIndex).Intervals.GolpeMagia = tmpLong
                        Case "MAGIA-GOLPE", "HECHIZO-GOLPE"
                            UserList(tUser.ArrayIndex).Intervals.MagiaGolpe = tmpLong
                        Case "GOLPE-USAR"
                            UserList(tUser.ArrayIndex).Intervals.GolpeUsar = tmpLong
                        Case "TRABAJAR", "WORK", "TRABAJO"
                            UserList(tUser.ArrayIndex).Intervals.TrabajarConstruir = tmpLong
                            UserList(tUser.ArrayIndex).Intervals.TrabajarExtraer = tmpLong
                        Case "TRABAJAR_EXTRAER", "EXTRAER", "TRABAJO_EXTRAER"
                            UserList(tUser.ArrayIndex).Intervals.TrabajarExtraer = tmpLong
                        Case "TRABAJAR_CONSTRUIR", "CONSTRUIR", "TRABAJO_CONSTRUIR"
                            UserList(tUser.ArrayIndex).Intervals.TrabajarConstruir = tmpLong
                        Case Else
                            Exit Sub
                    End Select
                    Call WriteIntervals(tUser.ArrayIndex)
                End If
            Case e_EditOptions.eo_Hogar
                Arg1 = UCase$(Arg1)
                Select Case Arg1
                    Case "NIX"
                        UserList(tUser.ArrayIndex).Hogar = e_Ciudad.cNix
                    Case "ULLA", "ULLATHORPE"
                        UserList(tUser.ArrayIndex).Hogar = e_Ciudad.cUllathorpe
                    Case "BANDER", "BANDERBILL"
                        UserList(tUser.ArrayIndex).Hogar = e_Ciudad.cBanderbill
                    Case "LINDOS"
                        UserList(tUser.ArrayIndex).Hogar = e_Ciudad.cLindos
                    Case "ARGHAL"
                        UserList(tUser.ArrayIndex).Hogar = e_Ciudad.cArghal
                    Case "ARKHEIN"
                        UserList(tUser.ArrayIndex).Hogar = e_Ciudad.cArkhein
                End Select
            Case Else
                ' Msg546=Comando no permitido.
                Call WriteLocaleMsg(UserIndex, "546", e_FontTypeNames.FONTTYPE_INFO)
        End Select
        'Log it!
        commandString = "/MOD "
        Select Case opcion
            Case e_EditOptions.eo_Gold
                commandString = commandString & "ORO "
            Case e_EditOptions.eo_Experience
                commandString = commandString & "EXP "
            Case e_EditOptions.eo_Body
                commandString = vbNullString
            Case e_EditOptions.eo_Head
                commandString = vbNullString
            Case e_EditOptions.eo_CriminalsKilled
                commandString = commandString & "CRI "
            Case e_EditOptions.eo_CiticensKilled
                commandString = commandString & "CIU "
            Case e_EditOptions.eo_Level
                commandString = commandString & "LEVEL "
            Case e_EditOptions.eo_Class
                commandString = commandString & "CLASE "
            Case e_EditOptions.eo_Skills
                commandString = commandString & "SKILLS "
            Case e_EditOptions.eo_SkillPointsLeft
                commandString = commandString & "SKILLSLIBRES "
            Case e_EditOptions.eo_Sex
                commandString = commandString & "SEX "
            Case e_EditOptions.eo_Raza
                commandString = commandString & "RAZA "
            Case e_EditOptions.eo_Vida
                commandString = commandString & "VIDA "
            Case e_EditOptions.eo_Mana
                commandString = commandString & "MANA "
            Case e_EditOptions.eo_Energia
                commandString = commandString & "ENERGIA "
            Case e_EditOptions.eo_MinHP
                commandString = commandString & "MINHP "
            Case e_EditOptions.eo_MinMP
                commandString = commandString & "MINMP "
            Case e_EditOptions.eo_Hit
                commandString = commandString & "HIT "
            Case e_EditOptions.eo_MinHit
                commandString = commandString & "MINHIT "
            Case e_EditOptions.eo_MaxHit
                commandString = commandString & "MAXHIT "
            Case e_EditOptions.eo_Desc
                commandString = commandString & "DESC "
            Case e_EditOptions.eo_Intervalo
                commandString = commandString & "INTERVALO "
            Case e_EditOptions.eo_Hogar
                commandString = commandString & "HOGAR "
            Case e_EditOptions.eo_CASCO
                commandString = vbNullString
            Case e_EditOptions.eo_Arma
                commandString = vbNullString
            Case e_EditOptions.eo_Escudo
                commandString = vbNullString
            Case Else
                commandString = commandString & "UNKOWN "
        End Select
        If commandString <> vbNullString Then
            Call LogGM(.name, commandString & Arg1 & " " & Arg2 & " " & username)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleEditChar", Erl)
End Sub

Public Sub HandleRequestCharInfo(ByVal UserIndex As Integer)
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim targetName As String
        Dim tUser      As t_UserReference
        targetName = Replace$(reader.ReadString8(), "+", " ")
        tUser = NameIndex(targetName)
        If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios) Then
            'is the player offline?
            If IsValidUserRef(tUser) Then
                'don't allow to retrieve administrator's info
                If UserList(tUser.ArrayIndex).flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(UserIndex, tUser.ArrayIndex)
                End If
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharInfo", Erl)
End Sub

' Handles the "RequestCharStats" message.
Public Sub HandleRequestCharStats(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        If (Not .flags.Privilegios And e_PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) <> 0 Then
            Call LogGM(.name, "/STAT " & username)
            tUser = NameIndex(username)
            If IsValidUserRef(tUser) Then
                Call SendUserMiniStatsTxt(UserIndex, tUser.ArrayIndex)
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharStats", Erl)
End Sub

' Handles the "RequestCharGold" message.
Public Sub HandleRequestCharGold(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        tUser = NameIndex(username)
        ' Verificar si el usuario tiene privilegios suficientes
        If (.flags.Privilegios And (e_PlayerType.Admin)) Then
            Call LogGM(.name, "/BAL " & username)
            If IsValidUserRef(tUser) Then
                Call WriteLocaleMsg(UserIndex, 1612, e_FontTypeNames.FONTTYPE_TALK, username & "¬" & UserList(tUser.ArrayIndex).Stats.Banco) 'Msg1612=El usuario ¬1 tiene ¬2 en el banco
            End If
        Else
            ' Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharGold", Erl)
End Sub

' Handles the "RequestCharInventory" message.
Public Sub HandleRequestCharInventory(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        tUser = NameIndex(username)
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            Call LogGM(.name, "/INV " & username)
            If IsValidUserRef(tUser) Then
                Call SendUserInvTxt(UserIndex, tUser.ArrayIndex)
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharInventory", Erl)
End Sub

Public Sub HandleRequestCharBank(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        tUser = NameIndex(username)
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            Call LogGM(.name, "/BOV " & username)
            If Not IsValidUserRef(tUser) Then
                ' Msg547=Usuario offline.
                Call WriteLocaleMsg(UserIndex, "547", e_FontTypeNames.FONTTYPE_TALK)
            Else
                Call SendUserBovedaTxt(UserIndex, tUser.ArrayIndex)
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharBank", Erl)
End Sub

Public Sub HandleRequestCharSkills(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        tUser = NameIndex(username)
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            Call LogGM(.name, "/STATS " & username)
            If Not IsValidUserRef(tUser) Then
                If (InStrB(username, "\") <> 0) Then
                    username = Replace(username, "\", "")
                End If
                If (InStrB(username, "/") <> 0) Then
                    username = Replace(username, "/", "")
                End If
            Else
                Call SendUserSkillsTxt(UserIndex, tUser.ArrayIndex)
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharSkills", Erl)
End Sub

Public Sub HandleReviveChar(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            If UCase$(username) <> "YO" Then
                tUser = NameIndex(username)
            Else
                Call SetUserRef(tUser, UserIndex)
            End If
            If Not IsValidUserRef(tUser) Then
                ' Msg743=Usuario offline.
                Call WriteLocaleMsg(UserIndex, "743", e_FontTypeNames.FONTTYPE_INFO)
            Else
                With UserList(tUser.ArrayIndex)
                    If .flags.Muerto = 1 Then
                        If UserList(UserIndex).flags.Privilegios And e_PlayerType.SemiDios Then
                            If MapInfo(.pos.Map).Seguro = 0 Or EsMapaEvento(.pos.Map) = False Then
                                'Msg962= Servidor » No puedes revivir en una zona insegura.
                                Call WriteLocaleMsg(UserIndex, "962", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        End If
                        .flags.Muerto = 0
                        Call RevivirUsuario(tUser.ArrayIndex)
                        Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1853, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1853=¬1 te ha resucitado.
                    Else
                        Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1854, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1854=¬1 te ha curado.
                    End If
                    .Stats.MinHp = .Stats.MaxHp
                    .Stats.MinMAN = .Stats.MaxMAN
                    .Stats.MinSta = .Stats.MaxSta
                End With
                Call WriteUpdateHP(tUser.ArrayIndex)
                Call WriteUpdateMana(tUser.ArrayIndex)
                Call WriteUpdateSta(tUser.ArrayIndex)
                Call ActualizarVelocidadDeUsuario(tUser.ArrayIndex)
                Call LogGM(.name, "Resucito a " & username)
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleReviveChar", Erl)
End Sub

Public Sub HandleOnlineGM(ByVal UserIndex As Integer)
    On Error GoTo HandleOnlineGM_Err
    'Author: Fredy Horacio Treboux (liquid)
    Dim i    As Long
    Dim list As String
    Dim priv As e_PlayerType
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        priv = e_PlayerType.Consejero Or e_PlayerType.SemiDios
        If .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then priv = priv Or e_PlayerType.Dios Or e_PlayerType.Admin
        For i = 1 To LastUser
            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then list = list & UserList(i).name & ", "
            End If
        Next i
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(UserIndex, list & ".", e_FontTypeNames.FONTTYPE_INFO)
        Else
            'Msg963= No hay GMs Online.
            Call WriteLocaleMsg(UserIndex, "963", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleOnlineGM_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnlineGM", Erl)
End Sub

Public Sub HandleOnlineMap(ByVal UserIndex As Integer)
    On Error GoTo HandleOnlineMap_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim LoopC As Long
        Dim list  As String
        Dim priv  As e_PlayerType
        priv = e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios
        If .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then priv = priv + (e_PlayerType.Dios Or e_PlayerType.Admin)
        For LoopC = 1 To LastUser
            If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).pos.Map = .pos.Map Then
                If UserList(LoopC).flags.Privilegios And priv Then list = list & UserList(LoopC).name & ", "
            End If
        Next LoopC
        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        Call WriteLocaleMsg(UserIndex, "1491", e_FontTypeNames.FONTTYPE_INFO, list) ' Msg1491=Usuarios en el mapa: ¬1
    End With
    Exit Sub
HandleOnlineMap_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnlineMap", Erl)
End Sub

Public Sub HandleForgive(ByVal UserIndex As Integer)
    On Error GoTo HandleForgive_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        'Se asegura que el target es un npc
        If Not IsValidNpcRef(.flags.TargetNPC) Then
            'Msg964= Primero tenés que seleccionar al sacerdote.
            Call WriteLocaleMsg(UserIndex, "964", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Validate NPC and make sure player is not dead
        If (NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Revividor And (NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.ResucitadorNewbie Or Not EsNewbie( _
                UserIndex))) Or .flags.Muerto = 1 Then Exit Sub
        Dim priest As t_Npc
        priest = NpcList(.flags.TargetNPC.ArrayIndex)
        'Make sure it's close enough
        If Distancia(.pos, priest.pos) > 3 Then
            'Msg965= El sacerdote no puede escuchar tus pecados debido a que estás demasiado lejos.
            Call WriteLocaleMsg(UserIndex, "965", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .Faccion.Status = e_Facciones.Ciudadano Or .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Then
            Call WriteLocaleChatOverHead(UserIndex, "1342", vbNullString, priest.Char.charindex, vbWhite) ' Msg1342=Tu alma ya esta libre de pecados hijo mio.
            Exit Sub
        End If
        If .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.consejo Then
            Call WriteLocaleChatOverHead(UserIndex, "1343", vbNullString, priest.Char.charindex, vbWhite) ' Msg1343=¡¡Dios no te perdonará mientras seas fiel al Demonio!!
            Exit Sub
        End If
        If .GuildIndex <> 0 Then
            If modGuilds.Alineacion(.GuildIndex) = 1 Then
                Call WriteLocaleChatOverHead(UserIndex, "1344", vbNullString, priest.Char.charindex, vbWhite) ' Msg1344=Te encuentras en un clan criminal... debes retirarte para que pueda perdonarte.
                Exit Sub
            End If
        End If
        If .Faccion.ciudadanosMatados > 0 Then
            Dim Donacion As Long
            Donacion = .Faccion.ciudadanosMatados * SvrConfig.GetValue("GoldMult") * SvrConfig.GetValue("CostoPerdonPorCiudadano")
            Call WriteLocaleChatOverHead(UserIndex, "1345", Donacion, priest.Char.charindex, vbWhite) ' Msg1345=Has matado a ciudadanos inocentes, Dios no puede perdonarte lo que has hecho. Pero si haces una generosa donación de, digamos, ¬1 monedas de oro, tal vez cambie de opinión...
            Exit Sub
        Else
            Donacion = SvrConfig.GetValue("CostoPerdonPorCiudadano") / 2
            Call WriteLocaleChatOverHead(UserIndex, 1346, Donacion, priest.Char.charindex, vbWhite)  ' Msg1346=Para volver a ser un ciudadano deberás Donar ¬1 monedas de oro.
            Exit Sub
        End If
        Dim permitePerdon As Boolean
        permitePerdon = False
        If .GuildIndex > 0 And (GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CAOTICA Or GuildAlignmentIndex(.GuildIndex) = _
                e_ALINEACION_GUILD.ALINEACION_CRIMINAL) Then
            permitePerdon = False
        Else
            permitePerdon = True
        End If
        If Not permitePerdon Then
            Call WriteLocaleChatOverHead(UserIndex, "1347", "", priest.Char.charindex, vbYellow) ' Msg1347=No podrás ser perdonado perteneciendo a un clan de alineación Criminal o de Alineación Oscura.
            Exit Sub
        End If
        Call WriteLocaleChatOverHead(UserIndex, "1348", "", priest.Char.charindex, vbYellow) ' Msg1348=Con estas palabras, te libero de todo tipo de pecados. ¡Que Dios te acompañe hijo mío!
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, "80", 100, False))
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave("100", UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        Call VolverCiudadano(UserIndex)
    End With
    Exit Sub
HandleForgive_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleForgive", Erl)
End Sub

Public Sub HandleKick(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        Dim rank     As Integer
        rank = e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            tUser = NameIndex(username)
            If Not IsValidUserRef(tUser) Then
                ' Msg557=El usuario no esta online.
                Call WriteLocaleMsg(UserIndex, "557", e_FontTypeNames.FONTTYPE_INFO)
            Else
                If (UserList(tUser.ArrayIndex).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                    ' Msg558=No podes echar a alguien con jerarquia mayor a la tuya.
                    Call WriteLocaleMsg(UserIndex, "558", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageLocaleMsg(1702, .name & "¬" & username, e_FontTypeNames.FONTTYPE_INFO)) 'Msg1702=Servidor » ¬1 echo a ¬2.
                    Call CloseSocket(tUser.ArrayIndex)
                    Call LogGM(.name, "Echo a " & username)
                End If
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleKick", Erl)
End Sub

Public Sub HandleExecute(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            tUser = NameIndex(username)
            If IsValidUserRef(tUser) Then
                Call CustomScenarios.UserDie(UserIndex)
                Call UserMod.UserDie(tUser.ArrayIndex)
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageLocaleMsg(1832, .name & "¬" & UserList(tUser.ArrayIndex).name, e_FontTypeNames.FONTTYPE_EJECUCION)) ' Msg1832=¬1 ha ejecutado a ¬2.
                Call LogGM(.name, " ejecuto a " & username)
            Else
                ' Msg559=No está online
                Call WriteLocaleMsg(UserIndex, "559", e_FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleExecute", Erl)
End Sub

Public Sub HandleBanChar(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim Reason   As String
        username = reader.ReadString8()
        Reason = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            Call BanPJ(UserIndex, username, Reason)
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleBanChar", Erl)
End Sub

Public Sub HandleUnbanChar(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
            If Not PersonajeExiste(username) Then
                ' Msg560=El personaje no existe.
                Call WriteLocaleMsg(UserIndex, "560", e_FontTypeNames.FONTTYPE_INFO)
            Else
                If BANCheck(username) Then
                    Call SavePenaDatabase(username, .name & ": UNBAN. " & Date & " " & Time)
                    Call UnBanDatabase(username)
                    Call LogGM(.name, "/UNBAN a " & username)
                    Call WriteLocaleMsg(UserIndex, 1613, e_FontTypeNames.FONTTYPE_INFO, username) 'Msg1613=¬1 desbaneado.
                Else
                    Call WriteLocaleMsg(UserIndex, 1614, e_FontTypeNames.FONTTYPE_INFO, username) 'Msg1614=¬1 no esta baneado. Imposible desbanear.
                End If
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleUnbanChar", Erl)
End Sub

Public Sub HandleNPCFollow(ByVal UserIndex As Integer)
    On Error GoTo HandleNPCFollow_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If IsValidNpcRef(.flags.TargetNPC) Then
            Call DoFollow(.flags.TargetNPC.ArrayIndex, .name)
            NpcList(.flags.TargetNPC.ArrayIndex).flags.Inmovilizado = 0
            NpcList(.flags.TargetNPC.ArrayIndex).flags.Paralizado = 0
            NpcList(.flags.TargetNPC.ArrayIndex).Contadores.Paralisis = 0
        End If
    End With
    Exit Sub
HandleNPCFollow_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleNPCFollow", Erl)
End Sub

Public Sub HandleSummonChar(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios) Then
            If LenB(username) <> 0 Then
                tUser = NameIndex(username)
                If Not IsValidUserRef(tUser) Then
                    ' Msg562=El jugador no está online.
                    Call WriteLocaleMsg(UserIndex, "562", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            ElseIf IsValidUserRef(.flags.TargetUser) Then
                tUser = .flags.TargetUser
                ' Mover NPCs
            ElseIf IsValidNpcRef(.flags.TargetNPC) Then
                If NpcList(.flags.TargetNPC.ArrayIndex).pos.Map = .pos.Map Then
                    Call WarpNpcChar(.flags.TargetNPC.ArrayIndex, .pos.Map, .pos.x, .pos.y + 1, True)
                    ' Msg563=Has desplazado a la criatura.
                    Call WriteLocaleMsg(UserIndex, "563", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    'Msg966= Sólo puedes mover NPCs dentro del mismo mapa.
                    Call WriteLocaleMsg(UserIndex, "966", e_FontTypeNames.FONTTYPE_INFO)
                End If
                Exit Sub
            Else
                Exit Sub
            End If
            If CompararPrivilegiosUser(tUser.ArrayIndex, UserIndex) > 0 Then
                Call WriteLocaleMsg(UserIndex, "1492", e_FontTypeNames.FONTTYPE_INFO, (UserList(tUser.ArrayIndex).name)) ' Msg1492=Se le ha avisado a ¬1 que quieres traerlo a tu posición.
                Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1855, .name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1855=¬1 quiere transportarte a su ubicación. Escribe /ira ¬1 para ir.
                Exit Sub
            End If
            Dim NotConsejero As Boolean
            NotConsejero = (.flags.Privilegios And e_PlayerType.Consejero) = 0
            ' Consejeros sólo pueden traer en el mismo mapa
            If NotConsejero Or .pos.Map = UserList(tUser.ArrayIndex).pos.Map Then
                If .flags.Privilegios And (e_PlayerType.SemiDios) Then
                    If Not EsMapaEvento(.pos.Map) Then 'HarThaos ReyarB hay que sacar eso de mapa seguro y poner mapas de eventos [MapasEventos]
                        'Msg967= Solamente puedes traer usuarios a mapas de evento.
                        Call WriteLocaleMsg(UserIndex, "967", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    If UserList(tUser.ArrayIndex).flags.Muerto = 1 Then
                        'Msg968= No puedes transportar a un muerto. ellos tienen /hogar
                        Call WriteLocaleMsg(UserIndex, "968", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                ' Si el admin está invisible no mostramos el nombre
                If NotConsejero And .flags.AdminInvisible = 1 Then
                    'Msg969= Te han trasportado.
                    Call WriteLocaleMsg(tUser.ArrayIndex, "969", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(tUser.ArrayIndex, PrepareMessageLocaleMsg(1856, .name, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1856=¬1 te ha transportado.
                End If
                'Si lo sumonean a un mapa interdimensional desde uno no interdimensional me guardo la posición de donde viene.
                If EsMapaInterdimensional(.pos.Map) And Not EsMapaInterdimensional(UserList(tUser.ArrayIndex).pos.Map) Then
                    UserList(tUser.ArrayIndex).flags.ReturnPos = UserList(tUser.ArrayIndex).pos
                End If
                Call WarpToLegalPos(tUser.ArrayIndex, .pos.Map, .pos.x, .pos.y + 1, True, True)
                Call WriteLocaleMsg(UserIndex, "1493", e_FontTypeNames.FONTTYPE_INFO, UserList(tUser.ArrayIndex).name) ' Msg1493=Has traído a ¬1.
                Call LogGM(.name, "/SUM " & username & " Map:" & .pos.Map & " X:" & .pos.x & " Y:" & .pos.y)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1833, .name & "¬" & username & "¬" & .pos.Map, e_FontTypeNames.FONTTYPE_INFO)) ' Msg1833=¬1 ha trasladado al usuario ¬2 al Mapa ¬3.
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSummonChar", Erl)
End Sub

Public Sub HandleSpawnListRequest(ByVal UserIndex As Integer)
    On Error GoTo HandleSpawnListRequest_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        If .flags.Privilegios And e_PlayerType.User Then
            Exit Sub
        ElseIf .flags.Privilegios And e_PlayerType.Consejero Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        ElseIf .flags.Privilegios And (e_PlayerType.SemiDios) Then
            'Msg970= Servidor » La cantidad de NPCs disponible para tu rango está limitada.
            Call WriteLocaleMsg(UserIndex, "970", e_FontTypeNames.FONTTYPE_INFO)
        End If
        Call WriteSpawnList(UserIndex, UserList(UserIndex).flags.Privilegios And e_PlayerType.Admin Or e_PlayerType.Dios)
    End With
    Exit Sub
HandleSpawnListRequest_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSpawnListRequest", Erl)
End Sub

Public Sub HandleSpawnCreature(ByVal UserIndex As Integer)
    On Error GoTo HandleSpawnCreature_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        Dim Npc As Integer
        Npc = reader.ReadInt16()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            If Npc > 0 And Npc <= UBound(Declaraciones.SpawnList()) Then
                If Declaraciones.SpawnList(Npc).NpcName <> "Nada" And (Declaraciones.SpawnList(Npc).PuedeInvocar Or (.flags.Privilegios And (e_PlayerType.Dios Or _
                        e_PlayerType.Admin)) <> 0) Then
                    Call SpawnNpc(Declaraciones.SpawnList(Npc).NpcIndex, .pos, True, False)
                End If
            End If
            Call LogGM(.name, "Sumoneo " & Declaraciones.SpawnList(Npc).NpcName)
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleSpawnCreature_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSpawnCreature", Erl)
End Sub

Public Sub HandleResetNPCInventory(ByVal UserIndex As Integer)
    On Error GoTo HandleResetNPCInventory_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
        Call ResetNpcInv(.flags.TargetNPC.ArrayIndex)
        Call LogGM(.name, "/RESETINV " & NpcList(.flags.TargetNPC.ArrayIndex).name)
    End With
    Exit Sub
HandleResetNPCInventory_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleResetNPCInventory", Erl)
End Sub

Public Sub HandleCleanWorld(ByVal UserIndex As Integer)
    On Error GoTo HandleCleanWorld_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        'Msg971= Se han limpiado los items del suelo.
        Call WriteLocaleMsg(UserIndex, "971", e_FontTypeNames.FONTTYPE_INFO)
    End With
    Exit Sub
HandleCleanWorld_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCleanWorld", Erl)
End Sub

Public Sub HandleServerMessage(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Message As String
        Message = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) Then
            If LenB(Message) <> 0 Then
                Call LogGM(.name, "Mensaje Broadcast:" & Message)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & "> " & Message, e_FontTypeNames.FONTTYPE_CENTINELA))
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleServerMessage", Erl)
End Sub

Public Sub HandleNickToIP(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Pablo (ToxicWaste): Agrego para uqe el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        Dim priv     As e_PlayerType
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            tUser = NameIndex(username)
            Call LogGM(.name, "NICK2IP Solicito la IP de " & username)
            If .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then
                priv = e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.Admin
            Else
                priv = e_PlayerType.User
            End If
            If IsValidUserRef(tUser) Then
                If UserList(tUser.ArrayIndex).flags.Privilegios And priv Then
                    Call WriteLocaleMsg(UserIndex, "1494", e_FontTypeNames.FONTTYPE_INFO, username & "¬" & UserList(tUser.ArrayIndex).ConnectionDetails.IP)  ' Msg1494=El ip de ¬1 es ¬2
                    Dim IP    As String
                    Dim lista As String
                    Dim LoopC As Long
                    IP = UserList(tUser.ArrayIndex).ConnectionDetails.IP
                    For LoopC = 1 To LastUser
                        If UserList(LoopC).ConnectionDetails.IP = IP Then
                            If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).name & ", "
                                End If
                            End If
                        End If
                    Next LoopC
                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteLocaleMsg(UserIndex, "1495", e_FontTypeNames.FONTTYPE_INFO, IP & "¬" & lista)   ' Msg1495=Los personajes con ip ¬1 son: ¬2
                End If
            Else
                'Msg972= No hay ningun personaje con ese nick
                Call WriteLocaleMsg(UserIndex, "972", e_FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleNickToIP", Erl)
End Sub

Public Sub HandleIPToNick(ByVal UserIndex As Integer)
    On Error GoTo HandleIPToNick_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        Dim IP    As String
        Dim LoopC As Long
        Dim lista As String
        Dim priv  As e_PlayerType
        IP = reader.ReadInt8() & "."
        IP = IP & reader.ReadInt8() & "."
        IP = IP & reader.ReadInt8() & "."
        IP = IP & reader.ReadInt8()
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call LogGM(.name, "IP2NICK Solicito los Nicks de IP " & IP)
        If .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then
            priv = e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.Admin
        Else
            priv = e_PlayerType.User
        End If
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnectionDetails.IP = IP Then
                If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).name & ", "
                    End If
                End If
            End If
        Next LoopC
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteLocaleMsg(UserIndex, "1496", e_FontTypeNames.FONTTYPE_INFO, IP & "¬" & lista)  ' Msg1496=Los personajes con ip ¬1 son: ¬2
    End With
    Exit Sub
HandleIPToNick_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleIPToNick", Erl)
End Sub

Public Sub HandleTeleportCreate(ByVal UserIndex As Integer)
    On Error GoTo HandleTeleportCreate_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        Dim Mapa   As Integer
        Dim x      As Byte
        Dim y      As Byte
        Dim Radio  As Byte
        Dim Motivo As String
        Mapa = reader.ReadInt16()
        x = reader.ReadInt8()
        y = reader.ReadInt8()
        Radio = reader.ReadInt8()
        Motivo = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call LogGM(.name, "/CT " & Mapa & "," & x & "," & y & "," & Motivo)
        If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, x, y) Then Exit Sub
        If MapData(.pos.Map, .pos.x, .pos.y - 1).ObjInfo.ObjIndex > 0 Then Exit Sub
        If MapData(.pos.Map, .pos.x, .pos.y - 1).TileExit.Map > 0 Then Exit Sub
        If MapData(Mapa, x, y).ObjInfo.ObjIndex > 0 Then
            'Msg973= Hay un objeto en el piso en ese lugar
            Call WriteLocaleMsg(UserIndex, "973", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If MapData(Mapa, x, y).TileExit.Map > 0 Then
            'Msg974= No podés crear un teleport que apunte a la entrada de otro.
            Call WriteLocaleMsg(UserIndex, "974", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim Objeto As t_Obj
        Objeto.amount = 1
        Select Case Radio
            Case 1
                Objeto.ObjIndex = 370
            Case 2
                Objeto.ObjIndex = 371
            Case Else
                Objeto.ObjIndex = 378
        End Select
        Call MakeObj(Objeto, .pos.Map, .pos.x, .pos.y - 1)
        With MapData(.pos.Map, .pos.x, .pos.y - 1)
            .TileExit.Map = Mapa
            .TileExit.x = x
            .TileExit.y = y
        End With
    End With
    Exit Sub
HandleTeleportCreate_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleTeleportCreate", Erl)
End Sub

Public Sub HandleTeleportDestroy(ByVal UserIndex As Integer)
    On Error GoTo HandleTeleportDestroy_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        Dim Mapa As Integer
        Dim x    As Byte
        Dim y    As Byte
        '/dt
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Mapa = .flags.TargetMap
        x = .flags.TargetX
        y = .flags.TargetY
        If Not InMapBounds(Mapa, x, y) Then Exit Sub
        With MapData(Mapa, x, y)
            'Si no tengo objeto y no tengo traslado
            If .ObjInfo.ObjIndex = 0 And .TileExit.Map = 0 Then Exit Sub
            'Si no tengo objeto pero tengo traslado
            If .ObjInfo.ObjIndex = 0 And .TileExit.Map > 0 Then
                Call LogGM(UserList(UserIndex).name, "/DT: " & Mapa & "," & x & "," & y)
                .TileExit.Map = 0
                .TileExit.x = 0
                .TileExit.y = 0
                'si tengo objeto y traslado
            ElseIf .ObjInfo.ObjIndex > 0 And ObjData(.ObjInfo.ObjIndex).OBJType = e_OBJType.otTeleport Then
                Call LogGM(UserList(UserIndex).name, "/DT: " & Mapa & "," & x & "," & y)
                Call EraseObj(.ObjInfo.amount, Mapa, x, y)
                If MapData(.TileExit.Map, .TileExit.x, .TileExit.y).ObjInfo.ObjIndex = 651 Then
                    Call EraseObj(1, .TileExit.Map, .TileExit.x, .TileExit.y)
                End If
                .TileExit.Map = 0
                .TileExit.x = 0
                .TileExit.y = 0
            End If
        End With
    End With
    Exit Sub
HandleTeleportDestroy_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleTeleportDestroy", Erl)
End Sub

Public Sub HandleRainToggle(ByVal UserIndex As Integer)
    On Error GoTo HandleRainToggle_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call LogGM(.name, "/LLUVIA")
        Lloviendo = Not Lloviendo
        Nebando = Not Nebando
        Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
        Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())
        If Lloviendo Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(404, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
            Call SendData(SendTarget.ToAll, 0, PrepareMessageFlashScreen(&HF5D3F3, 250)) 'Rayo
            Call ApagarFogatas
        End If
    End With
    Exit Sub
HandleRainToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRainToggle", Erl)
End Sub

Public Sub HandleSetCharDescription(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim tUser As t_UserReference
        Dim Desc  As String
        Desc = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
            tUser = .flags.TargetUser
            If IsValidUserRef(tUser) Then
                UserList(tUser.ArrayIndex).DescRM = Desc
            Else
                'Msg975= Haz click sobre un personaje antes!
                Call WriteLocaleMsg(UserIndex, "975", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSetCharDescription", Erl)
End Sub

Public Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)
    On Error GoTo HanldeForceMIDIToMap_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        Dim midiID As Byte
        Dim Mapa   As Integer
        midiID = reader.ReadInt8
        Mapa = reader.ReadInt16
        'Solo dioses, admins y RMS
        If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, 50, 50) Then
                Mapa = .pos.Map
            End If
            If midiID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(MapInfo(.pos.Map).music_numberLow))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(midiID))
            End If
        End If
    End With
    Exit Sub
HanldeForceMIDIToMap_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HanldeForceMIDIToMap", Erl)
End Sub

Public Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)
    On Error GoTo HandleForceWAVEToMap_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        Dim waveID As Byte
        Dim Mapa   As Integer
        Dim x      As Byte
        Dim y      As Byte
        waveID = reader.ReadInt8()
        Mapa = reader.ReadInt16()
        x = reader.ReadInt8()
        y = reader.ReadInt8()
        'Solo dioses, admins y RMS
        If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, x, y) Then
                Mapa = .pos.Map
                x = .pos.x
                y = .pos.y
            End If
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayWave(waveID, x, y))
        End If
    End With
    Exit Sub
HandleForceWAVEToMap_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleForceWAVEToMap", Erl)
End Sub

Public Sub HandleTalkAsNPC(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Message As String
        Message = reader.ReadString8()
        'Solo dioses, admins y RMS
        If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.SemiDios)) Then
            'Asegurarse haya un NPC seleccionado
            If IsValidNpcRef(.flags.TargetNPC) Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC.ArrayIndex, PrepareMessageChatOverHead(Message, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite))
            Else
                'Msg976= Debes seleccionar el NPC por el que quieres hablar antes de usar este comando
                Call WriteLocaleMsg(UserIndex, "976", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleTalkAsNPC", Erl)
End Sub

Public Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)
    On Error GoTo HandleDestroyAllItemsInArea_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim x As Long
        Dim y As Long
        For y = .pos.y - MinYBorder + 1 To .pos.y + MinYBorder - 1
            For x = .pos.x - MinXBorder + 1 To .pos.x + MinXBorder - 1
                If x > 0 And y > 0 And x < 101 And y < 101 Then
                    If MapData(.pos.Map, x, y).ObjInfo.ObjIndex > 0 Then
                        If ItemNoEsDeMapa(MapData(.pos.Map, x, y).ObjInfo.ObjIndex) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .pos.Map, x, y)
                        End If
                    End If
                End If
            Next x
        Next y
        Call LogGM(UserList(UserIndex).name, "/MASSDEST")
    End With
    Exit Sub
HandleDestroyAllItemsInArea_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleDestroyAllItemsInArea", Erl)
End Sub

Public Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)
    On Error GoTo HandleItemsInTheFloor_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim tObj As Integer
        Dim x    As Long
        Dim y    As Long
        For x = 5 To 95
            For y = 5 To 95
                tObj = MapData(.pos.Map, x, y).ObjInfo.ObjIndex
                If tObj > 0 Then
                    If ObjData(tObj).OBJType <> e_OBJType.otTrees Then
                        Call WriteConsoleMsg(UserIndex, "(" & x & "," & y & ") " & ObjData(tObj).name, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Next y
        Next x
    End With
    Exit Sub
HandleItemsInTheFloor_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleItemsInTheFloor", Erl)
End Sub

Public Sub HandleMakeDumb(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
            tUser = NameIndex(username)
            'para deteccion de aoice
            If Not IsValidUserRef(tUser) Then
                'Msg977= Offline
                Call WriteLocaleMsg(UserIndex, "977", e_FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser.ArrayIndex)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleMakeDumb", Erl)
End Sub

Public Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)
    'Last Modification: 12/30/06
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
            tUser = NameIndex(username)
            'para deteccion de aoice
            If Not IsValidUserRef(tUser) Then
                'Msg978= Offline
                Call WriteLocaleMsg(UserIndex, "978", e_FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser.ArrayIndex)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleMakeDumbNoMore", Erl)
End Sub

Public Sub HandleSetTrigger(ByVal UserIndex As Integer)
    On Error GoTo HandleSetTrigger_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        Dim tTrigger As Byte
        Dim tLog     As String
        tTrigger = reader.ReadInt8()
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        If tTrigger >= 0 Then
            MapData(.pos.Map, .pos.x, .pos.y).trigger = tTrigger
            tLog = "Trigger " & tTrigger & " on the map " & .pos.Map & " " & .pos.x & "," & .pos.y
            Call LogGM(.name, tLog)
            Call WriteConsoleMsg(UserIndex, tLog, e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleSetTrigger_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSetTrigger", Erl)
End Sub

Public Sub HandleAskTrigger(ByVal UserIndex As Integer)
    On Error GoTo HandleAskTrigger_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    Dim tTrigger As Byte
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        tTrigger = MapData(.pos.Map, .pos.x, .pos.y).trigger
        Call LogGM(.name, "Miro el trigger en " & .pos.Map & "," & .pos.x & "," & .pos.y & ". Era " & tTrigger)
        Call WriteLocaleMsg(UserIndex, "1498", e_FontTypeNames.FONTTYPE_INFO, "MAP " & .pos.Map & "," & .pos.x & "," & .pos.y & ". = " & tTrigger) ' Msg1498=Trigger ¬1
    End With
    Exit Sub
HandleAskTrigger_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleAskTrigger", Erl)
End Sub

Public Sub HandleCreateItem(ByVal UserIndex As Integer)
    On Error GoTo HandleCreateItem_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        Dim tObj    As Integer
        Dim Cuantos As Integer
        tObj = reader.ReadInt16()
        Cuantos = reader.ReadInt16()
        ' Si es usuario, lo sacamos cagando.
        If Not EsGM(UserIndex) Or (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios)) Then
            'Msg979= Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "979", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        ' Si es Dios, dejamos crear un item siempre y cuando pueda estar en el inventario.
        If (.flags.Privilegios And e_PlayerType.Dios) <> 0 And ObjData(tObj).Agarrable = 1 Then Exit Sub
        ' Si hace mas de 10000, lo sacamos cagando.
        If Cuantos > MAX_INVENTORY_OBJS Then
            Call WriteLocaleMsg(UserIndex, "1499", e_FontTypeNames.FONTTYPE_TALK, CStr(MAX_INVENTORY_OBJS)) ' Msg1499=Solo podés crear hasta ¬1 unidades
            Exit Sub
        End If
        ' El indice proporcionado supera la cantidad minima o total de items existentes en el juego?
        If tObj < 1 Or tObj > NumObjDatas Then Exit Sub
        ' El nombre del objeto es nulo?
        If LenB(ObjData(tObj).name) = 0 Then Exit Sub
        Dim Objeto As t_Obj
        Objeto.amount = Cuantos
        Objeto.ObjIndex = tObj
        ' Chequeo si el objeto es AGARRABLE(para las puertas, arboles y demAs t_Objs. que no deberian estar en el inventario)
        '   0 = SI
        '   1 = NO
        If ObjData(tObj).Agarrable = 0 Then
            ' Trato de meterlo en el inventario.
            If MeterItemEnInventario(UserIndex, Objeto) Then
                'Call WriteLocaleMsg(UserIndex, "1500", e_FontTypeNames.FONTTYPE_INFO, LenB(ObjData(tObj).name)) ' Msg1500=Has creado ¬1 unidades de ¬2.
                Call WriteLocaleMsg(UserIndex, "1500", e_FontTypeNames.FONTTYPE_INFO, (ObjData(tObj).name & "¬" & CStr(Cuantos)))
            Else
                'Msg980= No tenes espacio en tu inventario para crear el item.
                Call WriteLocaleMsg(UserIndex, "980", e_FontTypeNames.FONTTYPE_INFO)
                ' Si no hay espacio y es Admin, lo tiro al piso.
                If (.flags.Privilegios And e_PlayerType.Admin) <> 0 Then
                    Call TirarItemAlPiso(.pos, Objeto)
                    Call WriteLocaleMsg(UserIndex, "1501", e_FontTypeNames.FONTTYPE_GUILD, Cuantos & "¬" & _
                            " ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!")    ' Msg1501=ATENCION: CREASTE [¬1¬2 ver ReyarB
                End If
            End If
        Else
            ' Crear el item NO AGARRARBLE y tirarlo al piso.
            ' Si no hay espacio y es Admin, lo tiro al piso.
            If (.flags.Privilegios And e_PlayerType.Admin) <> 0 Then
                Call TirarItemAlPiso(.pos, Objeto)
                Call WriteLocaleMsg(UserIndex, "1502", e_FontTypeNames.FONTTYPE_GUILD, Cuantos & "¬" & _
                        " ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!")  ' Msg1502=ATENCION: CREASTE [¬1¬2 ver ReyarB
            End If
        End If
        Call LogGM(.name, "/CI: " & tObj & " Cantidad : " & Cuantos)
    End With
    Exit Sub
HandleCreateItem_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateItem", Erl)
End Sub

Public Sub HandleDestroyItems(ByVal UserIndex As Integer)
    On Error GoTo HandleDestroyItems_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If MapData(.pos.Map, .pos.x, .pos.y).ObjInfo.ObjIndex = 0 Then Exit Sub
        Call LogGM(.name, "/DEST")
        Call EraseObj(MAX_INVENTORY_OBJS, .pos.Map, .pos.x, .pos.y)
    End With
    Exit Sub
HandleDestroyItems_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleDestroyItems", Erl)
End Sub

Public Sub HandleForceMIDIAll(ByVal UserIndex As Integer)
    On Error GoTo HandleForceMIDIAll_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        Dim midiID As Byte
        midiID = reader.ReadInt8()
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " broadcast musica: " & midiID, e_FontTypeNames.FONTTYPE_SERVER))
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))
    End With
    Exit Sub
HandleForceMIDIAll_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleForceMIDIAll", Erl)
End Sub

Public Sub HandleForceWAVEAll(ByVal UserIndex As Integer)
    On Error GoTo HandleForceWAVEAll_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        Dim waveID As Byte
        waveID = reader.ReadInt8()
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
    End With
    Exit Sub
HandleForceWAVEAll_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleForceWAVEAll", Erl)
End Sub

Public Sub HandleRemovePunishment(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username   As String
        Dim punishment As Byte
        Dim NewText    As String
        username = reader.ReadString8()
        punishment = reader.ReadInt8
        NewText = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
            If LenB(username) = 0 Then
                'Msg981= Utilice /borrarpena Nick@NumeroDePena@NuevaPena
                Call WriteLocaleMsg(UserIndex, "981", e_FontTypeNames.FONTTYPE_INFO)
            Else
                If (InStrB(username, "\") <> 0) Then
                    username = Replace(username, "\", "")
                End If
                If (InStrB(username, "/") <> 0) Then
                    username = Replace(username, "/", "")
                End If
                If PersonajeExiste(username) Then
                    Call LogGM(.name, "Borro la pena " & punishment & " de " & username & " y la cambió por: " & NewText)
                    Call CambiarPenaDatabase(username, punishment, .name & ": <" & NewText & "> " & Date & " " & Time)
                    'Msg982= Pena Modificada.
                    Call WriteLocaleMsg(UserIndex, "982", e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleRemovePunishment", Erl)
End Sub

Public Sub HandleTile_BlockedToggle(ByVal UserIndex As Integer)
    On Error GoTo HandleTile_BlockedToggle_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call LogGM(.name, "/BLOQ")
        If MapData(.pos.Map, .pos.x, .pos.y).Blocked = 0 Then
            MapData(.pos.Map, .pos.x, .pos.y).Blocked = e_Block.ALL_SIDES Or e_Block.GM
        Else
            MapData(.pos.Map, .pos.x, .pos.y).Blocked = 0
        End If
        Call Bloquear(True, .pos.Map, .pos.x, .pos.y, IIf(MapData(.pos.Map, .pos.x, .pos.y).Blocked > 0, e_Block.ALL_SIDES, 0))
    End With
    Exit Sub
HandleTile_BlockedToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleTile_BlockedToggle", Erl)
End Sub

Public Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)
    On Error GoTo HandleKillNPCNoRespawn_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    With UserList(UserIndex)
        If Not EsGM(UserIndex) Then Exit Sub
        If .flags.Privilegios And e_PlayerType.Consejero Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
        Call QuitarNPC(.flags.TargetNPC.ArrayIndex, eGMCommand)
        Call LogGM(.name, "/MATA " & NpcList(.flags.TargetNPC.ArrayIndex).name)
    End With
    Exit Sub
HandleKillNPCNoRespawn_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleKillNPCNoRespawn", Erl)
End Sub

Public Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)
    On Error GoTo HandleKillAllNearbyNPCs_Err
    'Author: Nicolas Matias Gonzalez (NIGO)
    'ReyarB
    With UserList(UserIndex)
        If Not EsGM(UserIndex) Then Exit Sub
        If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim x As Long
        Dim y As Long
        For y = .pos.y - MinYBorder + 1 To .pos.y + MinYBorder - 1
            For x = .pos.x - MinXBorder + 1 To .pos.x + MinXBorder - 1
                If x > 0 And y > 0 And x < 101 And y < 101 Then
                    If MapData(.pos.Map, x, y).NpcIndex > 0 Then
                        Call QuitarNPC(MapData(.pos.Map, x, y).NpcIndex, eGMCommand)
                    End If
                End If
            Next x
        Next y
        Call LogGM(.name, "/MASSKILL")
    End With
    Exit Sub
HandleKillAllNearbyNPCs_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleKillAllNearbyNPCs", Erl)
End Sub

Public Sub HandleLastIP(ByVal UserIndex As Integer)
    'Author: Martín Trionfetti (HarThaoS) - Fernando Quinteros (Lord Fers)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim LoopC    As Byte
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
            'Handle special chars
            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "\", "")
            End If
            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "/", "")
            End If
            If (InStrB(username, "+") <> 0) Then
                username = Replace(username, "+", " ")
            End If
            Dim tUser As t_UserReference
            tUser = NameIndex(username)
            Dim RS    As ADODB.Recordset
            Dim ipStr As String
            If IsValidUserRef(tUser) Then
                Call LogGM(.name, "/LASTIP " & username)
                Set RS = Query("SELECT last_ip FROM account WHERE id = ?", UserList(tUser.ArrayIndex).AccountID)
                'Revisamos si recibio un resultado
                If RS Is Nothing Then Exit Sub
                If RS.BOF Or RS.EOF Then Exit Sub
                'Obtenemos la variable
                ipStr = RS.Fields(0).value
            Else
                Dim account_id As String
                Set RS = Query("SELECT u.account_id FROM user u WHERE LOWER(u.name) = LOWER(?)", username)
                'Revisamos si recibio un resultado
                If RS Is Nothing Then Exit Sub
                If RS.BOF Or RS.EOF Then Exit Sub
                'Obtenemos la variable
                account_id = RS.Fields(0).value
                Set RS = Query("SELECT last_ip FROM account WHERE id = ?", account_id)
                'Revisamos si recibio un resultado
                If RS Is Nothing Then Exit Sub
                If RS.BOF Or RS.EOF Then Exit Sub
                'Obtenemos la variable
                ipStr = RS.Fields(0).value
            End If
            Dim countIps As Long
            countIps = UBound(Split(ipStr, ";"))
            If countIps <= 0 Then Exit Sub
            ReDim ip_list(0 To (countIps - 1)) As String
            ip_list = Split(ipStr, ";")
            'Msg983= Las últimas ips para el personaje son:
            Call WriteLocaleMsg(UserIndex, "983", e_FontTypeNames.FONTTYPE_INFO)
            For LoopC = 0 To (countIps - 1)
                Call WriteConsoleMsg(UserIndex, ip_list(LoopC), e_FontTypeNames.FONTTYPE_INFO)
            Next LoopC
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleLastIP", Erl)
End Sub

Public Sub HandleChangeMOTD(ByVal UserIndex As Integer)
    On Error GoTo HandleChangeMOTD_Err
    'Author: Juan Martín sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.RoleMaster Or e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
        Dim auxiliaryString As String
        Dim LoopC           As Long
        For LoopC = LBound(MOTD()) To UBound(MOTD())
            auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
        Next LoopC
        If Len(auxiliaryString) >= 2 Then
            If Right$(auxiliaryString, 2) = vbCrLf Then
                auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)
            End If
        End If
        Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)
    End With
    Exit Sub
HandleChangeMOTD_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMOTD", Erl)
End Sub

Public Sub HandleSetMOTD(ByVal UserIndex As Integer)
    'Author: Lucas Tavolaro Ortiz (Tavo)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim newMOTD           As String
        Dim auxiliaryString() As String
        Dim LoopC             As Long
        newMOTD = reader.ReadString8()
        auxiliaryString = Split(newMOTD, vbCrLf)
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
            Call LogGM(.name, "Ha fijado un nuevo MOTD")
            MaxLines = UBound(auxiliaryString()) + 1
            If MaxLines > 0 Then
                ReDim MOTD(1 To MaxLines)
                Call WriteVar(DatPath & "Motd.ini", "INIT", "NumLines", CStr(MaxLines))
                For LoopC = 1 To MaxLines
                    Call WriteVar(DatPath & "Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                    MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
                Next LoopC
                'Msg984= Se ha cambiado el MOTD con exito
                Call WriteLocaleMsg(UserIndex, "984", e_FontTypeNames.FONTTYPE_INFO)
            Else
                ReDim MOTD(0 To 0)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSetMOTD", Erl)
End Sub

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)
    On Error GoTo HandleCreateNPC_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        Dim NpcIndex As Integer
        NpcIndex = reader.ReadInt16()
        If Not EsGM(UserIndex) Then Exit Sub
        If .flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        NpcIndex = SpawnNpc(NpcIndex, .pos, True, False)
        If NpcIndex <> 0 Then
            Call LogGM(.name, "Sumoneo a " & NpcList(NpcIndex).name & " en mapa " & .pos.Map)
        End If
    End With
    Exit Sub
HandleCreateNPC_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateNPC", Erl)
End Sub

Public Sub HandleCreateNPCWithRespawn(ByVal UserIndex As Integer)
    On Error GoTo HandleCreateNPCWithRespawn_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        Dim NpcIndex As Integer
        NpcIndex = reader.ReadInt16()
        If Not EsGM(UserIndex) Then Exit Sub
        If .flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios) Then
            'Msg985= Servidor » Comando deshabilitado para tu cargo. Si el motivo es probar algo ya saben ir a Test
            Call WriteLocaleMsg(UserIndex, "985", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        NpcIndex = SpawnNpc(NpcIndex, .pos, True, True)
        If NpcIndex <> 0 Then
            Call LogGM(.name, "Sumoneo con respawn " & NpcList(NpcIndex).name & " en mapa " & .pos.Map)
        End If
    End With
    Exit Sub
HandleCreateNPCWithRespawn_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateNPCWithRespawn", Erl)
End Sub

Public Sub HandleImperialArmour(ByVal UserIndex As Integer)
    On Error GoTo HandleImperialArmour_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        Dim Index    As Byte
        Dim ObjIndex As Integer
        Index = reader.ReadInt8()
        ObjIndex = reader.ReadInt16()
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        Select Case Index
            Case 1
                ' ArmaduraImperial1 = objindex
            Case 2
                ' ArmaduraImperial2 = objindex
            Case 3
                ' ArmaduraImperial3 = objindex
            Case 4
                ' TunicaMagoImperial = objindex
        End Select
    End With
    Exit Sub
HandleImperialArmour_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleImperialArmour", Erl)
End Sub

Public Sub HandleChaosArmour(ByVal UserIndex As Integer)
    On Error GoTo HandleChaosArmour_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        Dim Index    As Byte
        Dim ObjIndex As Integer
        Index = reader.ReadInt8()
        ObjIndex = reader.ReadInt16()
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        Select Case Index
            Case 1
                '   ArmaduraCaos1 = objindex
            Case 2
                '   ArmaduraCaos2 = objindex
            Case 3
                '   ArmaduraCaos3 = objindex
            Case 4
                '  TunicaMagoCaos = objindex
        End Select
    End With
    Exit Sub
HandleChaosArmour_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChaosArmour", Erl)
End Sub

Public Sub HandleTurnCriminal(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
            Call LogGM(.name, "/CONDEN " & username)
            tUser = NameIndex(username)
            If IsValidUserRef(tUser) Then Call VolverCriminal(tUser.ArrayIndex)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleTurnCriminal", Erl)
End Sub

Public Sub HandleAlterName(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        'Reads the userName and newUser Packets
        Dim username   As String
        Dim NewName    As String
        Dim tUser      As t_UserReference
        Dim GuildIndex As Integer
        username = UCase$(reader.ReadString8())
        NewName = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then Exit Sub
        If LenB(username) = 0 Or LenB(NewName) = 0 Then
            'Msg986= Usar: /ANAME origen@destino
            Call WriteLocaleMsg(UserIndex, "986", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        tUser = NameIndex(username)
        If IsValidUserRef(tUser) Then
            If UserList(tUser.ArrayIndex).GuildIndex > 0 Then
                Call WriteLocaleMsg(UserIndex, "1503", e_FontTypeNames.FONTTYPE_INFO, username)   ' Msg1503=El personaje ¬1¬2 ver ReyarB
                Exit Sub
            End If
        Else
            If Not PersonajeExiste(username) Then
                Call WriteLocaleMsg(UserIndex, "1504", e_FontTypeNames.FONTTYPE_INFO, username) ' Msg1504=El personaje ¬1 es inexistente.
                Exit Sub
            End If
            GuildIndex = GetUserGuildIndexDatabase(username)
            If GuildIndex > 0 Then
                Call WriteLocaleMsg(UserIndex, "1505", e_FontTypeNames.FONTTYPE_INFO, username)    ' Msg1505=El personaje ¬1¬2 ver ReyarB
                Exit Sub
            End If
        End If
        If PersonajeExiste(NewName) Then
            'Msg987= El nick solicitado ya existe.
            Call WriteLocaleMsg(UserIndex, "987", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call ChangeNameDatabase(username, NewName)
        'Msg988= Transferencia exitosa
        Call WriteLocaleMsg(UserIndex, "988", e_FontTypeNames.FONTTYPE_INFO)
        Call SavePenaDatabase(username, .name & ": nombre cambiado de """ & username & """ a """ & NewName & """. " & Date & " " & Time)
        Call SendData(SendTarget.ToGM, 0, PrepareMessageLocaleMsg(1834, .name & "¬" & username & "¬" & NewName, e_FontTypeNames.FONTTYPE_GM)) ' Msg1834=Administración » ¬1 cambió el nombre del usuario "¬2" por "¬3".
        Call LogGM(.name, "Ha cambiado de nombre al usuario """ & username & """. Ahora se llama """ & NewName & """.")
        If IsValidUserRef(tUser) Then
            UserList(tUser.ArrayIndex).name = NewName
            Call RefreshCharStatus(tUser.ArrayIndex)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleAlterName", Erl)
End Sub

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
    On Error GoTo HandleDoBackUp_Err
    'Author: Lucas Tavolaro Ortiz (Tavo)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        Call LogGM(.name, .name & " ha hecho un backup")
        Call ES.DoBackUp 'Sino lo confunde con la id del paquete
    End With
    Exit Sub
HandleDoBackUp_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleDoBackUp", Erl)
End Sub

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)
    On Error GoTo HandleChangeMapInfoPK_Err
    'Author: Lucas Tavolaro Ortiz (Tavo)
    With UserList(UserIndex)
        Dim isMapPk As Boolean
        isMapPk = reader.ReadBool()
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call LogGM(.name, .name & " ha cambiado la informacion sobre si es seguro el mapa.")
        MapInfo(.pos.Map).Seguro = IIf(isMapPk, 1, 0)
        Call WriteLocaleMsg(UserIndex, "1506", e_FontTypeNames.FONTTYPE_INFO, CStr(.pos.Map) & "¬" & IIf(isMapPk, "No", "Sí"))
    End With
    Exit Sub
HandleChangeMapInfoPK_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoPK", Erl)
End Sub

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)
    On Error GoTo HandleChangeMapInfoBackup_Err
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        Dim doTheBackUp As Boolean
        doTheBackUp = reader.ReadBool()
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        Call LogGM(.name, .name & " ha cambiado la información sobre el BackUp")
        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapInfo(.pos.Map).backup_mode = 1
        Else
            MapInfo(.pos.Map).backup_mode = 0
        End If
        'Change the boolean to string in a fast way
        Call WriteVar(MapPath & "mapa" & .pos.Map & ".dat", "Mapa" & .pos.Map, "backup", MapInfo(.pos.Map).backup_mode)
        Call WriteLocaleMsg(UserIndex, "1507", e_FontTypeNames.FONTTYPE_INFO, CStr(.pos.Map) & "¬" & IIf(doTheBackUp, "Sí", "No"))
    End With
    Exit Sub
HandleChangeMapInfoBackup_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoBackup", Erl)
End Sub

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)
    'Author: Pablo (ToxicWaste)
    'Restringido -> Options: "NEWBIE", "SINMAGIA", "SININVI", "NOPKS", "NOCIUD".
    On Error GoTo ErrHandler
    Dim tStr As String
    With UserList(UserIndex)
        tStr = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) <> 0 Then
            Select Case UCase$(tStr)
                Case "NEWBIE"
                    MapInfo(.pos.Map).Newbie = Not MapInfo(.pos.Map).Newbie
                    Call WriteLocaleMsg(UserIndex, "1508", e_FontTypeNames.FONTTYPE_INFO, CStr(.pos.Map) & "¬" & IIf(MapInfo(.pos.Map).Newbie, "Sí", "No"))  ' Msg1508=Mapa ¬1: Newbie = ¬2
                    Call LogGM(.name, .name & " ha cambiado la restricción del mapa " & .pos.Map & ": Newbie = " & MapInfo(.pos.Map).Newbie)
                Case "SINMAGIA"
                    MapInfo(.pos.Map).SinMagia = Not MapInfo(.pos.Map).SinMagia
                    Call WriteLocaleMsg(UserIndex, "1509", e_FontTypeNames.FONTTYPE_INFO, CStr(.pos.Map) & "¬" & IIf(MapInfo(.pos.Map).SinMagia, "Sí", "No"))  ' Msg1509=Mapa ¬1: SinMagia = ¬2
                    Call LogGM(.name, .name & " ha cambiado la restricción del mapa " & .pos.Map & ": SinMagia = " & MapInfo(.pos.Map).SinMagia)
                Case "NOPKS"
                    MapInfo(.pos.Map).NoPKs = Not MapInfo(.pos.Map).NoPKs
                    Call WriteLocaleMsg(UserIndex, "1510", e_FontTypeNames.FONTTYPE_INFO, CStr(.pos.Map) & "¬" & IIf(MapInfo(.pos.Map).NoPKs, "Sí", "No"))  ' Msg1510=Mapa ¬1: NoPKs = ¬2
                    Call LogGM(.name, .name & " ha cambiado la restricción del mapa " & .pos.Map & ": NoPKs = " & MapInfo(.pos.Map).NoPKs)
                Case "NOCIUD"
                    MapInfo(.pos.Map).NoCiudadanos = Not MapInfo(.pos.Map).NoCiudadanos
                    Call WriteLocaleMsg(UserIndex, "1511", e_FontTypeNames.FONTTYPE_INFO, CStr(.pos.Map) & "¬" & IIf(MapInfo(.pos.Map).NoCiudadanos, "Sí", "No"))  ' Msg1511=Mapa ¬1: NoCiudadanos = ¬2
                    Call LogGM(.name, .name & " ha cambiado la restricción del mapa " & .pos.Map & ": NoCiudadanos = " & MapInfo(.pos.Map).NoCiudadanos)
                Case "SININVI"
                    MapInfo(.pos.Map).SinInviOcul = Not MapInfo(.pos.Map).SinInviOcul
                    Call WriteLocaleMsg(UserIndex, "1512", e_FontTypeNames.FONTTYPE_INFO, CStr(.pos.Map) & "¬" & IIf(MapInfo(.pos.Map).SinInviOcul, "Sí", "No"))  ' Msg1512=Mapa ¬1: SinInvi = ¬2
                    Call LogGM(.name, .name & " ha cambiado la restricción del mapa " & .pos.Map & ": SinInvi = " & MapInfo(.pos.Map).SinInviOcul)
                Case Else
                    'Msg989= Opciones para restringir: 'NEWBIE', 'SINMAGIA', 'SININVI', 'NOPKS', 'NOCIUD'
                    Call WriteLocaleMsg(UserIndex, "989", e_FontTypeNames.FONTTYPE_INFO)
            End Select
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoRestricted", Erl)
End Sub

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)
    On Error GoTo HandleChangeMapInfoNoMagic_Err
    'Author: Pablo (ToxicWaste)
    'MagiaSinEfecto -> Options: "1" , "0".
    Dim nomagic As Boolean
    With UserList(UserIndex)
        nomagic = reader.ReadBool
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar la Magia el mapa.")
        End If
    End With
    Exit Sub
HandleChangeMapInfoNoMagic_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoNoMagic", Erl)
End Sub

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)
    On Error GoTo HandleChangeMapInfoNoInvi_Err
    'Author: Pablo (ToxicWaste)
    'InviSinEfecto -> Options: "1", "0"
    Dim noinvi As Boolean
    With UserList(UserIndex)
        noinvi = reader.ReadBool()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Invisibilidad el mapa.")
        End If
    End With
    Exit Sub
HandleChangeMapInfoNoInvi_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoNoInvi", Erl)
End Sub

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)
    On Error GoTo HandleChangeMapInfoNoResu_Err
    Dim noresu As Boolean
    With UserList(UserIndex)
        noresu = reader.ReadBool()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
            Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Resucitar el mapa.")
        End If
    End With
    Exit Sub
HandleChangeMapInfoNoResu_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoNoResu", Erl)
End Sub

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)
    'Author: Pablo (ToxicWaste)
    'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    On Error GoTo ErrHandler
    Dim tStr As String
    With UserList(UserIndex)
        tStr = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la informacion del Terreno del mapa.")
                MapInfo(UserList(UserIndex).pos.Map).terrain = tStr
                Call WriteVar(MapPath & "mapa" & UserList(UserIndex).pos.Map & ".dat", "Mapa" & UserList(UserIndex).pos.Map, "Terreno", tStr)
                Call WriteLocaleMsg(UserIndex, "1513", e_FontTypeNames.FONTTYPE_INFO, CStr(.pos.Map) & "¬" & tStr)  ' Msg1513=Mapa ¬1 Terreno: ¬2
            Else
                'Msg990= Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'
                Call WriteLocaleMsg(UserIndex, "990", e_FontTypeNames.FONTTYPE_INFO)
                'Msg991= Igualmente, el ínico ítil es 'NIEVE' ya que al ingresarlo, la gente muere de frio en el Mapa
                Call WriteLocaleMsg(UserIndex, "991", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)
    'Author: Pablo (ToxicWaste)
    'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    On Error GoTo ErrHandler
    Dim tStr As String
    With UserList(UserIndex)
        tStr = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.name, .name & " ha cambiado la informacion de la Zona del mapa.")
                MapInfo(UserList(UserIndex).pos.Map).zone = tStr
                Call WriteVar(MapPath & "mapa" & UserList(UserIndex).pos.Map & ".dat", "Mapa" & UserList(UserIndex).pos.Map, "Zona", tStr)
                Call WriteLocaleMsg(UserIndex, "1514", e_FontTypeNames.FONTTYPE_INFO, CStr(.pos.Map) & "¬" & tStr)  ' Msg1514=Mapa ¬1 Zona: ¬2
            Else
                'Msg992= Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'
                Call WriteLocaleMsg(UserIndex, "992", e_FontTypeNames.FONTTYPE_INFO)
                'Msg993= Igualmente, el ínico ítil es 'DUNGEON' ya que al ingresarlo, NO se sentirí el efecto de la lluvia en este mapa.
                Call WriteLocaleMsg(UserIndex, "993", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Public Sub HandleChangeMapSetting(ByVal UserIndex As Integer)
    On Error GoTo ChangeMapSetting_Err
    Dim SettingType As Byte
    SettingType = reader.ReadInt8()
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
            Select Case SettingType
                Case e_MapSetting.e_DropItems
                    MapInfo(UserList(UserIndex).pos.Map).DropItems = reader.ReadInt8()
                    Call LogGM(.name, .name & " ha cambiado la configuracion el dropeo de items en el mapa" & UserList(UserIndex).pos.Map & " a " & MapInfo(UserList( _
                            UserIndex).pos.Map).DropItems)
                    'Msg994= Mapa actualizado correctamente
                    Call WriteLocaleMsg(UserIndex, "994", e_FontTypeNames.FONTTYPE_INFO)
                Case e_MapSetting.e_SafeFight
                    MapInfo(UserList(UserIndex).pos.Map).SafeFightMap = reader.ReadInt8()
                    Call LogGM(.name, .name & " ha cambiado la configuracion el pelea segura del mapa" & UserList(UserIndex).pos.Map & " a " & MapInfo(UserList( _
                            UserIndex).pos.Map).DropItems)
                    'Msg995= Mapa actualizado correctamente
                    Call WriteLocaleMsg(UserIndex, "995", e_FontTypeNames.FONTTYPE_INFO)
                Case e_MapSetting.e_FriendlyFire
                    MapInfo(UserList(UserIndex).pos.Map).FriendlyFire = reader.ReadInt8()
                    Call LogGM(.name, .name & " ha cambiado la configuracion el friendly fire del mapa" & UserList(UserIndex).pos.Map & " a " & MapInfo(UserList( _
                            UserIndex).pos.Map).DropItems)
                    'Msg996= Mapa actualizado correctamente
                    Call WriteLocaleMsg(UserIndex, "996", e_FontTypeNames.FONTTYPE_INFO)
                Case Else
                    'Msg997= Opcion no disponible
                    Call WriteLocaleMsg(UserIndex, "997", e_FontTypeNames.FONTTYPE_INFO)
            End Select
        End If
    End With
    Exit Sub
ChangeMapSetting_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapSetting", Erl)
End Sub

Public Sub HandleSaveChars(ByVal UserIndex As Integer)
    On Error GoTo HandleSaveChars_Err
    'Author: Lucas Tavolaro Ortiz (Tavo)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call LogGM(.name, .name & " ha guardado todos los chars")
        Call GuardarUsuarios
    End With
    Exit Sub
HandleSaveChars_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSaveChars", Erl)
End Sub

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)
    On Error GoTo HandleCleanSOS_Err
    'Author: Lucas Tavolaro Ortiz (Tavo)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        Call LogGM(.name, .name & " ha borrado los SOS")
        Call Ayuda.Reset
    End With
    Exit Sub
HandleCleanSOS_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCleanSOS", Erl)
End Sub

Public Sub HandleShowServerForm(ByVal UserIndex As Integer)
    On Error GoTo HandleShowServerForm_Err
    'Author: Lucas Tavolaro Ortiz (Tavo)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        Call LogGM(.name, .name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click
    End With
    Exit Sub
HandleShowServerForm_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleShowServerForm", Erl)
End Sub

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)
    On Error GoTo HandleKickAllChars_Err
    'Author: Lucas Tavolaro Ortiz (Tavo)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        Call LogGM(.name, .name & " ha echado a todos los personajes.")
        Call EcharPjsNoPrivilegiados
    End With
    Exit Sub
HandleKickAllChars_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleKickAllChars", Erl)
End Sub

Public Sub HandleReloadNPCs(ByVal UserIndex As Integer)
    On Error GoTo HandleReloadNPCs_Err
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Reload the Server`s NPC
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        Call LogGM(.name, .name & " ha recargado los NPCs.")
        Call CargaNpcsDat
        'Msg998= Npcs.dat recargado exitosamente.
        Call WriteLocaleMsg(UserIndex, "998", e_FontTypeNames.FONTTYPE_SERVER)
    End With
    Exit Sub
HandleReloadNPCs_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleReloadNPCs", Erl)
End Sub

Public Sub HandleReloadServerIni(ByVal UserIndex As Integer)
    On Error GoTo HandleReloadServerIni_Err
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Reload the Server`s INI
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        Call LogGM(.name, .name & " ha recargado los INITs.")
        Call LoadSini
        Call LoadMD5
        Call LoadPrivateKey
    End With
    Exit Sub
HandleReloadServerIni_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleReloadServerIni", Erl)
End Sub

Public Sub HandleReloadSpells(ByVal UserIndex As Integer)
    On Error GoTo HandleReloadSpells_Err
    'Author: Lucas Tavolaro Ortiz (Tavo)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        Call LogGM(.name, .name & " ha recargado los hechizos.")
        Call CargarHechizos
    End With
    Exit Sub
HandleReloadSpells_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleReloadSpells", Erl)
End Sub

Public Sub HandleReloadObjects(ByVal UserIndex As Integer)
    On Error GoTo HandleReloadObjects_Err
    'Author: Lucas Tavolaro Ortiz (Tavo)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        Call LogGM(.name, .name & " ha recargado a los objetos.")
        Call LoadOBJData
        Call LoadPesca
        Call LoadRecursosEspeciales
        'Msg999= Obj.dat recargado exitosamente.
        Call WriteLocaleMsg(UserIndex, "999", e_FontTypeNames.FONTTYPE_SERVER)
    End With
    Exit Sub
HandleReloadObjects_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleReloadObjects", Erl)
End Sub

Public Sub HandleIgnored(ByVal UserIndex As Integer)
    'Author: Lucas Tavolaro Ortiz (Tavo)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible
        End If
    End With
    Exit Sub
HandleIgnored_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleIgnored", Erl)
End Sub

Public Sub HandleCheckSlot(ByVal UserIndex As Integer)
    'Author: Pablo (ToxicWaste)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        'Reads the UserName and Slot Packets
        Dim username As String
        Dim Slot     As Byte
        Dim tUser    As t_UserReference
        username = reader.ReadString8() 'Que UserName?
        Slot = reader.ReadInt8() 'Que Slot?
        tUser = NameIndex(username)
        If Not EsGM(UserIndex) Then Exit Sub
        Call LogGM(.name, .name & " Checkeo el slot " & Slot & " de " & username)
        If IsValidUserRef(tUser) Then
            If Slot > 0 And Slot <= UserList(UserIndex).CurrentInventorySlots Then
                If UserList(tUser.ArrayIndex).invent.Object(Slot).ObjIndex > 0 Then
                    Call WriteLocaleMsg(UserIndex, "1515", e_FontTypeNames.FONTTYPE_INFO, CStr(Slot) & "¬" & username & "¬" & UserList(tUser.ArrayIndex).invent.Object( _
                            Slot).amount)  ' Msg1515= Objeto ¬1) ¬2 Cantidad:¬3
                Else
                    'Msg1000= No hay Objeto en slot seleccionado
                    Call WriteLocaleMsg(UserIndex, "1000", e_FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                'Msg1001= Slot Invílido.
                Call WriteLocaleMsg(UserIndex, "1001", e_FontTypeNames.FONTTYPE_TALK)
            End If
        Else
            ' Msg547=Usuario offline.
            Call WriteLocaleMsg(UserIndex, "547", e_FontTypeNames.FONTTYPE_TALK)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCheckSlot", Erl)
End Sub

Public Sub HandleSetSpeed(ByVal UserIndex As Integer)
    On Error GoTo HandleGlobalOnOff_Err
    Dim Speed As Single
    Speed = reader.ReadReal32()
    'Author: Pablo Mercavides
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) = 0 Then Exit Sub
        UserList(UserIndex).Char.speeding = Speed
        Call WriteVelocidadToggle(Speed)
    End With
    Exit Sub
HandleGlobalOnOff_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGlobalOnOff", Erl)
End Sub

Public Sub HandleGlobalMessage(ByVal UserIndex As Integer)
    Dim nowRaw    As Long
    Dim elapsedMs As Double
    nowRaw = GetTickCountRaw()
    elapsedMs = TicksElapsed(UserList(UserIndex).Counters.MensajeGlobal, nowRaw)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim chat As String
        chat = reader.ReadString8()
        If .flags.Silenciado = 1 Then
            Call WriteLocaleMsg(UserIndex, "110", e_FontTypeNames.FONTTYPE_VENENO, .flags.MinutosRestantes)
        ElseIf elapsedMs < IntervaloMensajeGlobal Then
            ' Msg548=No puedes escribir mensajes globales tan rápido.
            Call WriteLocaleMsg(UserIndex, "548", e_FontTypeNames.FONTTYPE_WARNING)
        Else
            UserList(UserIndex).Counters.MensajeGlobal = nowRaw
            If SvrConfig.GetValue("ChatGlobal") = 1 Then
                If LenB(chat) <> 0 Then
                    Dim i As Integer
                    For i = 1 To UBound(.flags.ChatHistory) - 1
                        .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                    Next
                    .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
                    Dim MessageChat As String
                    MessageChat = "[" & .name & "] " & chat
                    Call modSendData.SendData(SendTarget.ToAll, 0, MessageChat, PrepareMessageConsoleMsg(MessageChat, e_FontTypeNames.FONTTYPE_GLOBAL))
                    Call LogThis(0, MessageChat, vbLogEventTypeInformation)
                    'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
                End If
            Else
                ' Msg549=El global se encuentra Desactivado.
                Call WriteLocaleMsg(UserIndex, "549", e_FontTypeNames.FONTTYPE_GLOBAL)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGlobalMessage", Erl)
End Sub

Public Sub HandleGlobalOnOff(ByVal UserIndex As Integer)
    On Error GoTo HandleGlobalOnOff_Err
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then Exit Sub
        Call LogGM(.name, " activo al Chat Global a las " & Now)
        If SvrConfig.GetValue("ChatGlobal") = 0 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1835, vbNullString, e_FontTypeNames.FONTTYPE_SERVER)) ' Msg1835=Servidor » Chat general habilitado. Escribe "/CONSOLA" o ";" y su mensaje para utilizarlo.
        Else
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1734, vbNullString, e_FontTypeNames.FONTTYPE_SERVER))
        End If
    End With
    Exit Sub
HandleGlobalOnOff_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGlobalOnOff", Erl)
End Sub

Public Sub HandleGiveItem(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim ObjIndex As Integer
        Dim Cantidad As Integer
        Dim Motivo   As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        ObjIndex = reader.ReadInt16()
        Cantidad = reader.ReadInt16()
        Motivo = reader.ReadString8()
        If (.flags.Privilegios And e_PlayerType.Admin) Then
            If ObjData(ObjIndex).Agarrable = 1 Then Exit Sub
            If Cantidad > MAX_INVENTORY_OBJS Then Cantidad = MAX_INVENTORY_OBJS
            ' El indice proporcionado supera la cantidad minima o total de items existentes en el juego?
            If ObjIndex < 1 Or ObjIndex > NumObjDatas Then Exit Sub
            ' El nombre del objeto es nulo?
            If LenB(ObjData(ObjIndex).name) = 0 Then Exit Sub
            ' Está online?
            tUser = NameIndex(username)
            If Not IsValidUserRef(tUser) Then
                Call WriteLocaleMsg(UserIndex, "1516", e_FontTypeNames.FONTTYPE_INFO)  ' Msg1516=El usuario ¬1 no está conectado.
                Exit Sub
            End If
            Dim Objeto As t_Obj
            Objeto.amount = Cantidad
            Objeto.ObjIndex = ObjIndex
            ' Trato de meterlo en el inventario.
            If MeterItemEnInventario(tUser.ArrayIndex, Objeto) Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1743, .name & "¬" & UserList(tUser.ArrayIndex).name & "¬" & Cantidad & "¬" & ObjData(ObjIndex).name & _
                        "¬" & Motivo, e_FontTypeNames.FONTTYPE_ROSA)) 'Msg1743=¬1 ha otorgado a ¬2 ¬3 ¬4: ¬5
            Else
                Call WriteLocaleMsg(UserIndex, "550", e_FontTypeNames.FONTTYPE_INFO) ' Msg550=El usuario no tiene espacio en el inventario.
            End If
            ' Lo registro en los logs.
            Call LogGM(.name, "/DAR " & username & " - Item: " & ObjData(ObjIndex).name & "(" & ObjIndex & ") Cantidad : " & Cantidad)
            Call LogPremios(ObjIndex, Cantidad, Motivo)
        Else
            ' Msg551=Servidor » Comando deshabilitado para tu cargo, debes pedir a un Admin que lo de.
            Call WriteLocaleMsg(UserIndex, "551", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGiveItem", Erl)
End Sub

Public Sub HandleQuestionGM(ByVal UserIndex As Integer)
    Dim nowRaw    As Long
    Dim elapsedMs As Double
    nowRaw = GetTickCountRaw()
    elapsedMs = TicksElapsed(UserList(UserIndex).Counters.LastGmMessage, nowRaw)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Consulta       As String
        Dim TipoDeConsulta As String
        Consulta = reader.ReadString8()
        TipoDeConsulta = reader.ReadString8()
        Dim PacketCounter As Long
        PacketCounter = reader.ReadInt32
        Dim Packet_ID As Long
        Packet_ID = PacketNames.QuestionGM
        If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "QuestionGM", PacketTimerThreshold( _
                Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        .Counters.CounterGmMessages = .Counters.CounterGmMessages + 1
        If .Counters.CounterGmMessages >= 20 Then
            Dim bannedip As String
            bannedip = UserList(UserIndex).ConnectionDetails.IP
            'Find every player with that ip and ban him!
            Dim i As Long
            For i = 1 To LastUser
                If UserList(i).ConnectionDetails.ConnIDValida Then
                    If UserList(i).ConnectionDetails.IP = bannedip Then
                        Call WriteCerrarleCliente(i)
                        Call CloseSocket(i)
                    End If
                End If
            Next i
        End If
        If elapsedMs < IntervaloConsultaGM Then
            ' Msg552=Solo puedes enviar una consulta cada 5 minutos.
            Call WriteLocaleMsg(UserIndex, "552", e_FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        UserList(UserIndex).Counters.LastGmMessage = nowRaw
        Call Ayuda.Push(.name, Consulta, TipoDeConsulta)
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageLocaleMsg(1836, UserList(UserIndex).name, e_FontTypeNames.FONTTYPE_SERVER)) ' Msg1836=Se ha recibido un nuevo mensaje de soporte de ¬1.
        .Counters.CounterGmMessages = 0
        ' Msg553=Tu mensaje fue recibido por el equipo de soporte.
        Call WriteLocaleMsg(UserIndex, "553", e_FontTypeNames.FONTTYPE_INFOIAO)
        Call LogConsulta(.name & " (" & TipoDeConsulta & ") " & Consulta)
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Public Sub HandleCuentaRegresiva(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Seconds As Byte
        Seconds = reader.ReadInt8()
        If Not .flags.Privilegios And e_PlayerType.User Then
            CuentaRegresivaTimer = Seconds
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1689, Seconds, e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1689=¡Empezando cuenta regresiva desde: ¬1 segundos...!
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCuentaRegresiva", Erl)
End Sub

Public Sub HandlePossUser(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        username = reader.ReadString8()
        If Not InStrB(username, "(") Then Exit Sub
        'Modifico la forma en que se usa el destrabar, ahora solamente lo puedo destrabar si está online.
        If (.flags.Privilegios And (e_PlayerType.User)) = 0 Then
            Dim tUser As t_UserReference
            username = Split(username, "(")(0)
            tUser = NameIndex(username)
            If IsValidUserRef(tUser) Then
                If LegalPosDestrabar(UserList(tUser.ArrayIndex).pos.Map, UserList(tUser.ArrayIndex).pos.x, UserList(tUser.ArrayIndex).pos.y, .flags.Navegando = 1, _
                        .flags.Navegando = 0) Then Exit Sub
                Dim nPos As t_WorldPos
                Call ClosestLegalPos(UserList(tUser.ArrayIndex).pos, nPos, False, True)
                Call FindLegalPos(tUser.ArrayIndex, UserList(tUser.ArrayIndex).pos.Map, CByte(UserList(tUser.ArrayIndex).pos.x), CByte(UserList(tUser.ArrayIndex).pos.y))
                Call WarpUserChar(tUser.ArrayIndex, nPos.Map, nPos.x, nPos.y, True)
                Call WriteLocaleMsg(UserIndex, "1517", e_FontTypeNames.FONTTYPE_INFO, username & "¬" & nPos.Map & "¬" & nPos.x & "¬" & nPos.y) ' Msg1517=Servidor » Acción realizada con exito! La nueva posicion de ¬1 es: ¬2-¬3-¬4.
                'ver porque si el usuario esta online lo dice igual
            Else
                ' Msg554=Servidor » El usuario debe estar deslogueado para dicha solicitud!
                Call WriteLocaleMsg(UserIndex, "554", e_FontTypeNames.FONTTYPE_INFO)
                'avisar al usuario que un gm lo destrabara pero debe desloguear
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandlePossUser", Erl)
End Sub

Public Sub HandleNieveToggle(ByVal UserIndex As Integer)
    On Error GoTo HandleNieveToggle_Err
    'Author: Pablo Mercavides
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call LogGM(.name, "/NIEVE")
        Nebando = Not Nebando
        Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())
    End With
    Exit Sub
HandleNieveToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleNieveToggle", Erl)
End Sub

Public Sub HandleNieblaToggle(ByVal UserIndex As Integer)
    On Error GoTo HandleNieblaToggle_Err
    'Author: Pablo Mercavides
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call LogGM(.name, "/NIEBLA")
        Call ResetMeteo
    End With
    Exit Sub
HandleNieblaToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleNieblaToggle", Erl)
End Sub

Public Sub HandleGenio(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo HandleGenio_Err
    With UserList(UserIndex)
        'Si no es GM, no pasara nada :P
        If (.flags.Privilegios And e_PlayerType.User) Then Exit Sub
        Dim i As Byte
        For i = 1 To NUMSKILLS
            .Stats.UserSkills(i) = 100
        Next i
        ' Msg555=Tus skills fueron editados.
        Call WriteLocaleMsg(UserIndex, "555", e_FontTypeNames.FONTTYPE_INFOIAO)
    End With
    Exit Sub
HandleGenio_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleGenio", Erl)
End Sub

Public Sub HandleBanCuenta(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim Reason   As String
        username = reader.ReadString8()
        Reason = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
            Call BanearCuenta(UserIndex, username, Reason)
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleBanCuenta", Erl)
End Sub

Public Sub HandleUnBanCuenta(ByVal UserIndex As Integer)
    ' /unbancuenta namepj
    ' /unbancuenta email
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim UserNameOEmail As String
        UserNameOEmail = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
            If DesbanearCuenta(UserNameOEmail) Then
                Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageLocaleMsg(1703, .name & "¬" & UserNameOEmail, e_FontTypeNames.FONTTYPE_SERVER)) 'Msg1703=Servidor » ¬1 ha desbaneado la cuenta de ¬2.
            Else
                ' Msg556=No se ha podido desbanear la cuenta.
                Call WriteLocaleMsg(UserIndex, "556", e_FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleUnBanCuenta", Erl)
End Sub

Public Sub HandleCerrarCliente(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        username = reader.ReadString8()
        ' Solo administradores pueden cerrar clientes ajenos
        If (.flags.Privilegios And e_PlayerType.Admin) Then
            tUser = NameIndex(username)
            If Not IsValidUserRef(tUser) Then
                ' Msg557=El usuario no esta online.
                Call WriteLocaleMsg(UserIndex, "557", e_FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageLocaleMsg(1704, .name & "¬" & username, e_FontTypeNames.FONTTYPE_INFO)) 'Msg1704=¬1 cerro el cliente de ¬2.
                Call WriteCerrarleCliente(tUser.ArrayIndex)
                Call LogGM(.name, "Cerro el cliene de:" & username)
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCerrarCliente", Erl)
End Sub

Public Sub HandleEventoInfo(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo HandleEventoInfo_Err
    With UserList(UserIndex)
        If EventoActivo Then
            Call WriteLocaleMsg(UserIndex, 1615, e_FontTypeNames.FONTTYPE_New_Eventos, PublicidadEvento & "¬" & TiempoRestanteEvento) 'Msg1615=¬1. Tiempo restante: ¬2 minuto(s).
        Else
            ' Msg729=Eventos> Actualmente no hay ningún evento en curso.
            Call WriteLocaleMsg(UserIndex, "729", e_FontTypeNames.FONTTYPE_New_Eventos)
        End If
        Dim i           As Byte
        Dim encontre    As Boolean
        Dim HoraProximo As Byte
        If Not HoraEvento + 1 >= 24 Then
            For i = HoraEvento + 1 To 23
                If Evento(i).Tipo <> 0 Then
                    encontre = True
                    HoraProximo = i
                    Exit For
                End If
            Next i
        End If
        If encontre = False Then
            For i = 0 To HoraEvento
                If Evento(i).Tipo <> 0 Then
                    encontre = True
                    HoraProximo = i
                    Exit For
                End If
            Next i
        End If
        If encontre Then
            Call WriteLocaleMsg(UserIndex, "1518", e_FontTypeNames.FONTTYPE_New_Eventos, PublicidadEvento & "¬" & HoraProximo)  ' Msg1518=Eventos> El proximo evento ¬1 iniciara a las ¬2:00 horas.
        Else
            ' Msg730=Eventos> No hay eventos próximos.
            Call WriteLocaleMsg(UserIndex, "730", e_FontTypeNames.FONTTYPE_New_Eventos)
        End If
    End With
    Exit Sub
HandleEventoInfo_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleEventoInfo", Erl)
End Sub

Public Sub HandleCrearEvento(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Tipo           As Byte
        Dim Duracion       As Byte
        Dim multiplicacion As Byte
        Tipo = reader.ReadInt8()
        Duracion = reader.ReadInt8()
        multiplicacion = reader.ReadInt8()
        If multiplicacion > 5 Then 'no superar este multiplicador
            multiplicacion = 2
        End If
        '/ dejar solo Administradores
        If .flags.Privilegios >= e_PlayerType.Admin Then
            If EventoActivo = False Then
                If LenB(Tipo) = 0 Or LenB(Duracion) = 0 Or LenB(multiplicacion) = 0 Then
                    ' Msg731=Utilice /CREAREVENTO TIPO@DURACION@MULTIPLICACION.
                    Call WriteLocaleMsg(UserIndex, "731", e_FontTypeNames.FONTTYPE_New_Eventos)
                Else
                    Call ForzarEvento(Tipo, Duracion, multiplicacion, UserList(UserIndex).name)
                End If
            Else
                ' Msg732=Ya hay un evento en curso. Finalicelo con /FINEVENTO primero.
                Call WriteLocaleMsg(UserIndex, "732", e_FontTypeNames.FONTTYPE_New_Eventos)
            End If
        Else
            ' Msg733=Servidor » Solo Administradores pueder crear estos eventos.
            Call WriteLocaleMsg(UserIndex, "733", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Public Sub HandleBanTemporal(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim Reason   As String
        Dim dias     As Byte
        username = reader.ReadString8()
        Reason = reader.ReadString8()
        dias = reader.ReadInt8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
            Call Admin.BanTemporal(username, dias, Reason, UserList(UserIndex).name)
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Public Sub HandleCancelarExit(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo HandleCancelarExit_Err
    Call CancelExit(UserIndex)
    Exit Sub
HandleCancelarExit_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCancelarExit", Erl)
End Sub

Public Sub HandleCrearTorneo(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim NivelMinimo As Byte
        Dim NivelMaximo As Byte
        Dim cupos       As Byte
        Dim costo       As Long
        Dim mago        As Byte
        Dim clerico     As Byte
        Dim guerrero    As Byte
        Dim asesino     As Byte
        Dim bardo       As Byte
        Dim druido      As Byte
        Dim Paladin     As Byte
        Dim cazador     As Byte
        Dim Trabajador  As Byte
        Dim Pirata      As Byte
        Dim Ladron      As Byte
        Dim Bandido     As Byte
        Dim Mapa        As Integer
        Dim x           As Byte
        Dim y           As Byte
        Dim nombre      As String
        Dim reglas      As String
        NivelMinimo = reader.ReadInt8
        NivelMaximo = reader.ReadInt8
        cupos = reader.ReadInt8
        costo = reader.ReadInt32
        mago = reader.ReadInt8
        clerico = reader.ReadInt8
        guerrero = reader.ReadInt8
        asesino = reader.ReadInt8
        bardo = reader.ReadInt8
        druido = reader.ReadInt8
        Paladin = reader.ReadInt8
        cazador = reader.ReadInt8
        Trabajador = reader.ReadInt8
        Pirata = reader.ReadInt8
        Ladron = reader.ReadInt8
        Bandido = reader.ReadInt8
        Mapa = reader.ReadInt16
        x = reader.ReadInt8
        y = reader.ReadInt8
        nombre = reader.ReadString8
        reglas = reader.ReadString8
        If EsGM(UserIndex) And ((.flags.Privilegios And e_PlayerType.Consejero) = 0) Then
            Torneo.NivelMinimo = NivelMinimo
            Torneo.NivelMaximo = NivelMaximo
            Torneo.cupos = cupos
            Torneo.costo = costo
            Torneo.mago = mago
            Torneo.clerico = clerico
            Torneo.guerrero = guerrero
            Torneo.asesino = asesino
            Torneo.bardo = bardo
            Torneo.druido = druido
            Torneo.Paladin = Paladin
            Torneo.cazador = cazador
            Torneo.Trabajador = Trabajador
            Torneo.Pirata = Pirata
            Torneo.Ladron = Ladron
            Torneo.Bandido = Bandido
            Torneo.Mapa = Mapa
            Torneo.x = x
            Torneo.y = y
            Torneo.nombre = nombre
            Torneo.reglas = reglas
            Call IniciarTorneo
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCrearTorneo", Erl)
End Sub

Public Sub HandleComenzarTorneo(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        If EsGM(UserIndex) Then
            Call ComenzarTorneoOk
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleComenzarTorneo", Erl)
End Sub

Public Sub HandleCancelarTorneo(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        If EsGM(UserIndex) Then
            Call ResetearTorneo
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleComenzarTorneo", Erl)
End Sub

Public Sub HandleBusquedaTesoro(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim Tipo As Byte
        Tipo = reader.ReadInt8()
        If (.flags.Privilegios And Not (e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.User)) Then
            Select Case Tipo
                Case 0
                    If Not BusquedaTesoroActiva And BusquedaRegaloActiva = False And BusquedaNpcActiva = False Then
                        Call PerderTesoro
                    Else
                        If BusquedaTesoroActiva Then
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1735, get_map_name(TesoroNumMapa) & "¬" & TesoroNumMapa, e_FontTypeNames.FONTTYPE_TALK))
                            Call WriteLocaleMsg(UserIndex, "1519", e_FontTypeNames.FONTTYPE_INFO)  ' Msg1519=Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: ¬1-¬2-¬3
                        Else
                            Call WriteLocaleMsg(UserIndex, "734", e_FontTypeNames.FONTTYPE_INFO) ' Msg734=Ya hay una busqueda del tesoro activa.
                        End If
                    End If
                Case 1
                    If Not BusquedaRegaloActiva And BusquedaTesoroActiva = False And BusquedaNpcActiva = False Then
                        Call PerderRegalo
                    Else
                        If BusquedaRegaloActiva Then
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1736, get_map_name(RegaloNumMapa) & "¬" & RegaloNumMapa, e_FontTypeNames.FONTTYPE_TALK))
                            Call WriteLocaleMsg(UserIndex, "1520", e_FontTypeNames.FONTTYPE_INFO)  ' Msg1520=Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: ¬1-¬2-¬3
                        Else
                            Call WriteLocaleMsg(UserIndex, "734", e_FontTypeNames.FONTTYPE_INFO) ' Msg734=Ya hay una busqueda del tesoro activa.
                        End If
                    End If
                Case 2
                    If Not BusquedaNpcActiva And BusquedaTesoroActiva = False And BusquedaRegaloActiva = False Then
                        Dim pos As t_WorldPos
                        pos.Map = TesoroNPCMapa(RandomNumber(1, UBound(TesoroNPCMapa)))
                        pos.y = 50
                        pos.x = 50
                        npc_index_evento = SpawnNpc(TesoroNPC(RandomNumber(1, UBound(TesoroNPC))), pos, True, False, True)
                        BusquedaNpcActiva = True
                    Else
                        If BusquedaNpcActiva Then
                            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1737, NpcList(npc_index_evento).pos.Map, e_FontTypeNames.FONTTYPE_TALK))
                            Call WriteLocaleMsg(UserIndex, "1521", e_FontTypeNames.FONTTYPE_INFO)  ' Msg1521=Ya hay una busqueda de npc activo. El tesoro se encuentra en: ¬1-¬2-¬3
                        Else
                            Call WriteLocaleMsg(UserIndex, "734", e_FontTypeNames.FONTTYPE_INFO) ' Msg734=Ya hay una busqueda del tesoro activa.
                        End If
                    End If
            End Select
        Else
            Call WriteLocaleMsg(UserIndex, "735", e_FontTypeNames.FONTTYPE_INFO) ' Msg735=Servidor » No estas habilitado para hacer Eventos.
        End If
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleBusquedaTesoro", Erl)
End Sub

Public Sub HandleMarcaDeGM(ByVal UserIndex As Integer)
    'Author: Pablo Mercavides
    On Error GoTo HandleMarcaDeGM_Err
    Call WriteWorkRequestTarget(UserIndex, e_Skill.MarcaDeGM)
    Exit Sub
HandleMarcaDeGM_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleMarcaDeGM", Erl)
End Sub

Public Sub HandleFinEvento(ByVal UserIndex As Integer)
    On Error GoTo HandleDenounce_Err
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.User Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If EventoActivo Then
            Call FinalizarEvento
        Else
            'Msg1002= No hay ningun evento activo.
            Call WriteLocaleMsg(UserIndex, "1002", e_FontTypeNames.FONTTYPE_New_Eventos)
        End If
    End With
    Exit Sub
HandleDenounce_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleDenounce", Erl)
End Sub

Public Sub HandleCreateEvent(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim name As String
        name = reader.ReadString8()
        If LenB(name) = 0 Then Exit Sub
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Select Case UCase$(name)
            Case "INVASION BANDER"
                Call IniciarEvento(TipoEvento.Invasion, 1)
                Call LogGM(.name, "Forzó el evento Invasión en Banderbille.")
            Case "INVASION CARCEL"
                Call IniciarEvento(TipoEvento.Invasion, 2)
                Call LogGM(.name, "Forzó el evento Invasión en Carcel.")
            Case Else
                Call WriteLocaleMsg(UserIndex, 1616, e_FontTypeNames.FONTTYPE_INFO, name) 'Msg1616=No existe el evento "¬1".
        End Select
    End With
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateEvent", Erl)
End Sub

Public Sub HandleStartEvent(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    Dim eventType As Byte
    eventType = reader.ReadInt8()
    Dim LobbySettings As t_NewScenearioSettings
    LobbySettings.ScenearioType = reader.ReadInt8
    LobbySettings.MinLevel = reader.ReadInt8
    LobbySettings.MaxLevel = reader.ReadInt8
    LobbySettings.MinPlayers = reader.ReadInt8
    LobbySettings.MaxPlayers = reader.ReadInt8
    LobbySettings.TeamSize = reader.ReadInt8
    LobbySettings.TeamType = reader.ReadInt8
    LobbySettings.RoundNumber = reader.ReadInt8
    LobbySettings.InscriptionFee = reader.ReadInt32
    LobbySettings.Description = reader.ReadString8
    LobbySettings.Password = reader.ReadString8
    CurrentActiveEventType = LobbySettings.ScenearioType
    Call initEventLobby(UserIndex, eventType, LobbySettings)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleStartEvent", Erl)
End Sub

Public Sub HandleCancelarEvento(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    If (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) = 0 Then
        'Msg528=Servidor » Comando deshabilitado para tu cargo.
        Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If CurrentActiveEventType = CaptureTheFlag Then
        If InstanciaCaptura Is Nothing Then
            'Msg1003= Eventos » No hay ninguna instancia en curso para ese evento.
            Call WriteLocaleMsg(UserIndex, "1003", e_FontTypeNames.FONTTYPE_INFO)
        Else
            Call InstanciaCaptura.finalizarCaptura
        End If
    Else
        If GlobalLobbyIndex >= 0 Then
            Call CancelLobby(LobbyList(GlobalLobbyIndex))
            If LobbyList(GlobalLobbyIndex).Scenario Is Nothing Then Call ReleaseLobby(GlobalLobbyIndex)
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1738, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1738=Eventos» El evento ha sido cancelado.
        Else
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1739, vbNullString, e_FontTypeNames.FONTTYPE_GUILD)) 'Msg1739=No se encontró ningún evento activo.
        End If
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleCancelarCaptura", Erl)
End Sub

Public Sub HandleSeguirMouse(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        Dim tempArea As Long
        username = reader.ReadString8()
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
            If UCase$(username) <> "YO" Then
                tUser = NameIndex(username)
            Else
                Call WriteLocaleMsg(UserIndex, "1004", e_FontTypeNames.FONTTYPE_INFO) 'Msg1004= No puedes seguirte a vos mismo
                Exit Sub
            End If
            Call LogGM(.name, "/SM " & username)
            If username = "" And .flags.SigueUsuario.ArrayIndex > 0 Then
                'Me devuelvo inventario y stats
                UserList(UserIndex).invent = UserList(UserIndex).Invent_bk
                UserList(UserIndex).Stats = UserList(UserIndex).Stats_bk
                'UserList(UserIndex).Char.charindex = UserList(UserIndex).Char.charindex_bk
                Call WriteLocaleMsg(UserIndex, "1522", e_FontTypeNames.FONTTYPE_INFO)  ' Msg1522=Dejas de seguir a ¬1.
                Call WriteCancelarSeguimiento(UserIndex)
                Call WriteNotificarClienteSeguido(.flags.SigueUsuario.ArrayIndex, 0)
                Call SetUserRef(UserList(.flags.SigueUsuario.ArrayIndex).flags.GMMeSigue, 0)
                Call WriteUserCharIndexInServer(UserIndex)
                Call WarpUserChar(UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, False)
                Call SetUserRef(UserList(UserIndex).flags.SigueUsuario, 0)
            ElseIf tUser.ArrayIndex <> .flags.SigueUsuario.ArrayIndex And .flags.SigueUsuario.ArrayIndex > 0 And IsValidUserRef(tUser) Then
                If IsValidUserRef(UserList(tUser.ArrayIndex).flags.GMMeSigue) And UserList(tUser.ArrayIndex).flags.GMMeSigue.ArrayIndex <> UserIndex Then
                    Call WriteLocaleMsg(UserIndex, "1523", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).name)  ' Msg1523=El usuario está siendo seguido por ¬1.
                    Exit Sub
                End If
                'Me devuelvo inventario y stats
                UserList(UserIndex).invent = UserList(UserIndex).Invent_bk
                UserList(UserIndex).Stats = UserList(UserIndex).Stats_bk
                'UserList(UserIndex).Char.charindex = UserList(UserIndex).Char.charindex_bk
                Call WriteLocaleMsg(UserIndex, "1524", e_FontTypeNames.FONTTYPE_INFO, username)  ' Msg1524=Dejas de seguir a ¬1.
                Call WriteCancelarSeguimiento(UserIndex)
                Call WriteNotificarClienteSeguido(.flags.SigueUsuario.ArrayIndex, 0)
                Call SetUserRef(UserList(.flags.SigueUsuario.ArrayIndex).flags.GMMeSigue, 0)
                Call WriteUserCharIndexInServer(UserIndex)
                Call WarpUserChar(UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, False)
                Call SetUserRef(UserList(UserIndex).flags.SigueUsuario, 0)
                Call DoAdminInvisible(UserIndex, 1) 'HACER LO DE ADMIN INVISIBLE
                'Call WarpUserChar(UserIndex, UserList(tUser).Pos.map, UserList(tUser).Pos.X, UserList(tUser).Pos.y, False)
                'Me backupeo el inventario y los hechizos
                UserList(UserIndex).Invent_bk = UserList(UserIndex).invent
                UserList(UserIndex).Stats_bk = UserList(UserIndex).Stats
                'Me pego el inventario y los hechizos del usuario seguido
                UserList(UserIndex).invent = UserList(tUser.ArrayIndex).invent
                UserList(UserIndex).Stats = UserList(tUser.ArrayIndex).Stats
                UserList(UserIndex).flags.SigueUsuario = tUser
                'Actualizo flag en cliente para que empiece a enviar paquetes
                Call WriteNotificarClienteSeguido(tUser.ArrayIndex, 1)
                Call SetUserRef(UserList(tUser.ArrayIndex).flags.GMMeSigue, UserIndex)
                Call WriteLocaleMsg(UserIndex, "1525", e_FontTypeNames.FONTTYPE_INFO, username)  ' Msg1525=Comienzas a seguir a ¬1.
                tempArea = UserList(UserIndex).AreasInfo.AreaID
                Call WarpUserChar(UserIndex, UserList(tUser.ArrayIndex).pos.Map, 15, 15)
                Call WriteSendFollowingCharindex(UserIndex, UserList(tUser.ArrayIndex).Char.charindex)
                If tempArea <> UserList(tUser.ArrayIndex).AreasInfo.AreaID Then
                    Call MakeUserChar(False, UserIndex, tUser.ArrayIndex, UserList(tUser.ArrayIndex).pos.Map, UserList(tUser.ArrayIndex).pos.x, UserList(tUser.ArrayIndex).pos.y)
                End If
                Call WritePosUpdateCharIndex(UserIndex, UserList(tUser.ArrayIndex).pos.x, UserList(tUser.ArrayIndex).pos.y, UserList(tUser.ArrayIndex).Char.charindex)
                Call WriteUserCharIndexInServer(tUser.ArrayIndex)
                UserList(UserIndex).AreasInfo = UserList(tUser.ArrayIndex).AreasInfo
                Call ModAreas.CheckUpdateNeededUser(UserIndex, 5, 0)
            ElseIf Not IsValidUserRef(tUser) Then
                ' Msg743=Usuario offline.
                Call WriteLocaleMsg(UserIndex, "743", e_FontTypeNames.FONTTYPE_INFO)
            Else
                'Si empiezo a seguir a alguien
                If Not IsValidUserRef(UserList(UserIndex).flags.SigueUsuario) Then
                    If IsValidUserRef(UserList(tUser.ArrayIndex).flags.GMMeSigue) And UserList(tUser.ArrayIndex).flags.GMMeSigue.ArrayIndex <> UserIndex Then
                        Call WriteLocaleMsg(UserIndex, "1526", e_FontTypeNames.FONTTYPE_INFO)  ' Msg1526=El usuario está siendo seguido por ¬1.
                        Exit Sub
                    End If
                    Call DoAdminInvisible(UserIndex, 1) 'HACER LO DE ADMIN INVISIBLE
                    'Call WarpUserChar(UserIndex, UserList(tUser).Pos.map, UserList(tUser).Pos.X, UserList(tUser).Pos.y, False)
                    'Me backupeo el inventario y los hechizos
                    UserList(UserIndex).Invent_bk = UserList(UserIndex).invent
                    UserList(UserIndex).Stats_bk = UserList(UserIndex).Stats
                    'Me pego el inventario y los hechizos del usuario seguido
                    UserList(UserIndex).invent = UserList(tUser.ArrayIndex).invent
                    UserList(UserIndex).Stats = UserList(tUser.ArrayIndex).Stats
                    UserList(UserIndex).flags.SigueUsuario = tUser
                    'Actualizo flag en cliente para que empiece a enviar paquetes
                    Call WriteNotificarClienteSeguido(tUser.ArrayIndex, 1)
                    Call SetUserRef(UserList(tUser.ArrayIndex).flags.GMMeSigue, UserIndex)
                    Call WriteLocaleMsg(UserIndex, "1527", e_FontTypeNames.FONTTYPE_INFO, username) ' Msg1527=Comienzas a seguir a ¬1.
                    tempArea = UserList(UserIndex).AreasInfo.AreaID
                    Call WarpUserChar(UserIndex, UserList(tUser.ArrayIndex).pos.Map, 15, 15)
                    Call WriteSendFollowingCharindex(UserIndex, UserList(tUser.ArrayIndex).Char.charindex)
                    If tempArea <> UserList(tUser.ArrayIndex).AreasInfo.AreaID Then
                        Call MakeUserChar(False, UserIndex, tUser.ArrayIndex, UserList(tUser.ArrayIndex).pos.Map, UserList(tUser.ArrayIndex).pos.x, UserList( _
                                tUser.ArrayIndex).pos.y)
                    End If
                    Call WritePosUpdateCharIndex(UserIndex, UserList(tUser.ArrayIndex).pos.x, UserList(tUser.ArrayIndex).pos.y, UserList(tUser.ArrayIndex).Char.charindex)
                    Call WriteUserCharIndexInServer(tUser.ArrayIndex)
                    UserList(UserIndex).AreasInfo = UserList(tUser.ArrayIndex).AreasInfo
                    Call ModAreas.CheckUpdateNeededUser(UserIndex, 5, 0)
                    Call WriteAreaChanged(UserIndex, UserList(tUser.ArrayIndex).pos.x, UserList(tUser.ArrayIndex).pos.y)
                Else
                    If UserList(UserIndex).flags.SigueUsuario.ArrayIndex <> tUser.ArrayIndex Then
                        Call WriteLocaleMsg(UserIndex, "1528", e_FontTypeNames.FONTTYPE_INFO, username) ' Msg1528=Ya te encuentras siguiendo a un usuario, para dejar de seguirlo escribe /SM ¬1.
                        Exit Sub
                    End If
                    'Me devuelvo inventario y stats
                    UserList(UserIndex).invent = UserList(UserIndex).Invent_bk
                    UserList(UserIndex).Stats = UserList(UserIndex).Stats_bk
                    Call SetUserRef(UserList(UserIndex).flags.SigueUsuario, 0)
                    Call WriteLocaleMsg(UserIndex, "1529", e_FontTypeNames.FONTTYPE_INFO, username) ' Msg1529=Dejas de seguir a ¬1.
                    Call WriteCancelarSeguimiento(UserIndex)
                    Call WriteNotificarClienteSeguido(tUser.ArrayIndex, 0)
                    Call SetUserRef(UserList(tUser.ArrayIndex).flags.GMMeSigue, 0)
                    Call WriteUserCharIndexInServer(UserIndex)
                    Call WarpUserChar(UserIndex, UserList(UserIndex).pos.Map, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y, False)
                End If
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Call UpdateUserInv(True, UserIndex, 1)
    Call UpdateUserHechizos(True, UserIndex, 0)
    Call WriteUpdateUserStats(UserIndex)
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleReviveChar", Erl)
End Sub

Public Sub HandleEventoFaccionario(ByVal UserIndex As Integer)
    On Error GoTo HandleEventoFaccionario_Err:
    If (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
        EnEventoFaccionario = Not EnEventoFaccionario
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageLocaleMsg(1837, IIf(EnEventoFaccionario, "iniciado", "finalizado"), e_FontTypeNames.FONTTYPE_SERVER)) ' Msg1837=Servidor » Se ha ¬1 el evento faccionario.
    End If
    Exit Sub
HandleEventoFaccionario_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleEventoFaccionario", Erl)
End Sub

Public Sub HandleDebugRequest(ByVal UserIndex As Integer)
    On Error GoTo HandleDebugRequest_Err:
    Dim debugType As Byte
    Dim Args()    As String
    debugType = reader.ReadInt8()
    If debugType = 0 Then
        ReDim Args(0) As String
    ElseIf debugType = 1 Then
        ReDim Args(1) As String
        Args(0) = reader.ReadString8()
    ElseIf debugType = 2 Then
        ReDim Args(0) As String
    Else
        Exit Sub
    End If
    If (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) Then
        Call WriteDebugLogResponse(UserIndex, debugType, Args)
    End If
    Exit Sub
HandleDebugRequest_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleDebugRequest", Erl)
End Sub

Public Sub HandleLobbyCommand(ByVal UserIndex As Integer)
    On Error GoTo HandleLobbyCommand_err
    Dim Command As Byte
    Dim Params  As String
    Command = reader.ReadInt8()
    Params = reader.ReadString8()
    With UserList(UserIndex)
        If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios) Then
            If Not HandleRemoteLobbyCommand(Command, Params, UserIndex, GlobalLobbyIndex) Then
                'Msg1005= Servidor » No se pudo procesar el comando.
                Call WriteLocaleMsg(UserIndex, "1005", e_FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            'Msg528=Servidor » Comando deshabilitado para tu cargo.
            Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleLobbyCommand_err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleLobbyCommand", Erl)
End Sub

Public Sub HandleFeatureToggle(ByVal UserIndex As Integer)
    On Error GoTo HandleFeatureToggle_Err:
    Dim value    As Byte
    Dim name     As String
    Dim nameSize As Integer
    value = max(Min(1, reader.ReadInt8), 0)
    name = reader.ReadString8
    nameSize = Len(nameSize)
    If nameSize = 0 Or nameSize > 100 Then
        Exit Sub
    End If
    If (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Admin)) Then
        If name = "SGRACEFULLY" Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1740, vbNullString, e_FontTypeNames.FONTTYPE_PROMEDIO_MENOR)) 'Msg1740=Servidor » cerrando ahora.
            Call GuardarUsuarios
            Call EcharPjsNoPrivilegiados
            frmMain.GuardarYCerrar = True
            Unload frmMain
        Else
            Call SetFeatureToggle(name, value > 0)
            Call WriteLocaleMsg(UserIndex, "1006", e_FontTypeNames.FONTTYPE_INFO) 'Msg1006= variable configurada correctamente.
        End If
    Else
        Call WriteLocaleMsg(UserIndex, "1007", e_FontTypeNames.FONTTYPE_INFO) 'Msg1007= no tienes permisos para realizar esta accion.
    End If
    Exit Sub
HandleFeatureToggle_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleFeatureToggle", Erl)
End Sub

'HarThaoS: Iniciar captura de bandera
Public Sub HandleIniciarCaptura(EventSettings As t_NewScenearioSettings)
    On Error GoTo ErrHandler
    If Not InstanciaCaptura Is Nothing Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1008, "", e_FontTypeNames.FONTTYPE_GLOBAL))
        Exit Sub
    Else
        'El precio no puede ser negativo
        If EventSettings.InscriptionFee < 0 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1009, "", e_FontTypeNames.FONTTYPE_GLOBAL))
            Exit Sub
        End If
        'Permito un máximo de 48 participantes
        If EventSettings.MaxPlayers > 48 Then 'Leer de una variable de configuración
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1011, "", e_FontTypeNames.FONTTYPE_GLOBAL))
            Exit Sub
        End If
        If EventSettings.MinLevel < 1 Or EventSettings.MinLevel > 47 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1013, "", e_FontTypeNames.FONTTYPE_GLOBAL))
            Exit Sub
        End If
        If EventSettings.MinLevel > EventSettings.MaxLevel Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageLocaleMsg(1014, "", e_FontTypeNames.FONTTYPE_GLOBAL))
            Exit Sub
        End If
        Set InstanciaCaptura = New clsCaptura
        Call InstanciaCaptura.inicializar(EventSettings.MaxPlayers, EventSettings.RoundNumber, EventSettings.MinLevel, EventSettings.MaxLevel, EventSettings.InscriptionFee)
    End If
    Exit Sub
ErrHandler:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleIniciarCaptura", Erl)
End Sub
