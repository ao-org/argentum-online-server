Attribute VB_Name = "UsUaRiOs"
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Public Function ConnectUser_Check(ByVal UserIndex As Integer, _
                                  ByRef Name As String, _
                                  ByRef UserCuenta As String) As Boolean
                           
    On Error GoTo Check_ConnectUser_Err

    With UserList(UserIndex)
        
        ConnectUser_Check = False
        
        If .flags.UserLogged Then
            Call LogCheating("El usuario " & .Name & " ha intentado loguear a " & Name & " desde la IP " & .IP)
            
            'Kick player ( and leave character inside :D )!
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)
            
            Exit Function

        End If
            
        '¿Ya esta conectado el personaje?
        Dim tIndex As Integer: tIndex = NameIndex(Name)

        If tIndex > 0 And tIndex <> UserIndex Then

            If UserList(tIndex).Counters.Saliendo Then

                Call WriteShowMessageBox(UserIndex, "El personaje está saliendo.")

            Else
                
                Call WriteShowMessageBox(UserIndex, "El personaje ya está conectado. Espere mientras es desconectado.")

                ' Le avisamos al usuario que está jugando, en caso de que haya uno
                Call WriteShowMessageBox(tIndex, "Alguien está ingresando con tu personaje. Si no has sido tú, por favor cambia la contraseña de tu cuenta.")
                Call Cerrar_Usuario(tIndex)

            End If

            Call CloseSocket(UserIndex)
            Exit Function

        End If
        
        '¿Supera el máximo de usuarios por cuenta?
        If MaxUsersPorCuenta > 0 Then

            If ContarUsuariosMismaCuenta(.AccountID) >= MaxUsersPorCuenta Then

                If MaxUsersPorCuenta = 1 Then
                    Call WriteShowMessageBox(UserIndex, "Ya hay un usuario conectado con esta cuenta.")
                Else
                    Call WriteShowMessageBox(UserIndex, "La cuenta ya alcanzó el máximo de " & MaxUsersPorCuenta & " usuarios conectados.")

                End If

                Call CloseSocket(UserIndex)
                Exit Function

            End If

        End If
        
        'Controlamos no pasar el maximo de usuarios
        If NumUsers >= MaxUsers Then
            Call WriteShowMessageBox(UserIndex, "El servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo mas tarde.")
            Call CloseSocket(UserIndex)
            Exit Function

        End If
        
        '¿Este IP ya esta conectado?
        If MaxConexionesIP > 0 Then

            If ContarMismaIP(UserIndex, .IP) >= MaxConexionesIP Then
                Call WriteShowMessageBox(UserIndex, "Has alcanzado el límite de conexiones por IP.")
                Call CloseSocket(UserIndex)
                Exit Function

            End If

        End If
        
        'Le damos los privilegios
        .flags.Privilegios = UserDarPrivilegioLevel(Name)

        'Add RM flag if needed
        If EsRolesMaster(Name) Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.RoleMaster
        End If
        
        If EsGM(UserIndex) Then

            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » " & Name & " se conecto al juego.", FontTypeNames.FONTTYPE_INFOBOLD))
            Call LogGM(.Name, "Se conectó con IP: " & .IP)

        Else

            If ServerSoloGMs > 0 Then
                Call WriteShowMessageBox(UserIndex, "Servidor restringido a administradores. Por favor reintente en unos momentos.")
                Call CloseSocket(UserIndex)
                Exit Function

            End If

        End If
        
        If EnPausa Then

            Call WritePauseToggle(UserIndex)
            Call WriteConsoleMsg(UserIndex, "Servidor » Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
            Call CloseSocket(UserIndex)
            Exit Function

        End If
        
        If UCase$(.Cuenta) <> UCase$(UserCuenta) Then
            Call WriteShowMessageBox(UserIndex, "El personaje no corresponde a su cuenta.")
            Call CloseSocket(UserIndex)
            Exit Function

        End If
    
    End With
    
    ConnectUser_Check = True

    Exit Function

Check_ConnectUser_Err:
    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ConnectUser_Check", Erl)
        
End Function

Public Sub ConnectUser_Prepare(ByVal UserIndex As Integer, ByRef Name As String, ByRef UserCuenta As String)

        On Error GoTo Prepare_ConnectUser_Err

100     With UserList(UserIndex)
    
            'Reseteamos los FLAGS
105         .flags.Escondido = 0
110         .flags.TargetNPC = 0
115         .flags.TargetNpcTipo = eNPCType.Comun
120         .flags.TargetObj = 0
125         .flags.TargetUser = 0
130         .Char.FX = 0
135         .Counters.CuentaRegresiva = -1

            ' Seteamos el nombre
170         .Name = Name
        
            ' Vinculamos el nombre con el UserIndex
175         m_NameIndex(UCase$(Name)) = UserIndex

180         .showName = True
    
        End With
    
        Exit Sub

Prepare_ConnectUser_Err:
        Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ConnectUser_Prepare", Erl)

End Sub

Public Function ConnectUser_Complete(ByVal UserIndex As Integer, _
                                     ByRef Name As String, _
                                     ByRef UserCuenta As String)

        On Error GoTo Complete_ConnectUser_Err
        
        Dim n    As Integer
        Dim tStr As String
        
100     With UserList(UserIndex)
            
            ' -----------------------------------------------------------------------
            '   INFORMACION INICIAL DEL PERSONAJE
            ' -----------------------------------------------------------------------

105         If .flags.Paralizado = 1 Then
110             .Counters.Paralisis = IntervaloParalizado
            End If

115         If .flags.Muerto = 0 Then
120             .Char = .OrigChar
            
125             If .Char.Body = 0 Then
130                 Call DarCuerpoDesnudo(UserIndex)
                End If
            
135             If .Char.Head = 0 Then
140                 .Char.Head = 1
                End If
            Else
145             .Char.Body = iCuerpoMuerto
150             .Char.Head = iCabezaMuerto
155             .Char.WeaponAnim = NingunArma
160             .Char.ShieldAnim = NingunEscudo
165             .Char.CascoAnim = NingunCasco
170             .Char.Heading = eHeading.SOUTH
            End If
        
            'Obtiene el indice-objeto del arma
175         If .Invent.WeaponEqpSlot > 0 Then
180             If .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex > 0 Then
185                 .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex

190                 If .flags.Muerto = 0 Then
195                     .Char.Arma_Aura = ObjData(.Invent.WeaponEqpObjIndex).CreaGRH
                    End If
                Else
200                 .Invent.WeaponEqpSlot = 0
                End If
            End If

            'Obtiene el indice-objeto del armadura
205         If .Invent.ArmourEqpSlot > 0 Then
210             If .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex > 0 Then
215                 .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex

220                 If .flags.Muerto = 0 Then
225                     .Char.Body_Aura = ObjData(.Invent.ArmourEqpObjIndex).CreaGRH
                    End If
                Else
230                 .Invent.ArmourEqpSlot = 0
                End If
235             .flags.Desnudo = 0
            Else
240             .flags.Desnudo = 1
            End If

            'Obtiene el indice-objeto del escudo
245         If .Invent.EscudoEqpSlot > 0 Then
250             If .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex > 0 Then
255                 .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex

260                 If .flags.Muerto = 0 Then
265                     .Char.Escudo_Aura = ObjData(.Invent.EscudoEqpObjIndex).CreaGRH
                    End If
                Else
270                 .Invent.EscudoEqpSlot = 0
                End If
            End If
        
            'Obtiene el indice-objeto del casco
275         If .Invent.CascoEqpSlot > 0 Then
280             If .Invent.Object(.Invent.CascoEqpSlot).ObjIndex > 0 Then
285                 .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).ObjIndex

290                 If .flags.Muerto = 0 Then
295                     .Char.Head_Aura = ObjData(.Invent.CascoEqpObjIndex).CreaGRH
                    End If
                Else
300                 .Invent.CascoEqpSlot = 0
                End If
            End If

            'Obtiene el indice-objeto barco
305         If .Invent.BarcoSlot > 0 Then
310             If .Invent.Object(.Invent.BarcoSlot).ObjIndex > 0 Then
315                 .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).ObjIndex
                Else
320                 .Invent.BarcoSlot = 0
                End If
            End If

            'Obtiene el indice-objeto municion
325         If .Invent.MunicionEqpSlot > 0 Then
330             If .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex > 0 Then
335                 .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex
                Else
340                 .Invent.MunicionEqpSlot = 0
                End If
            End If

            ' DM
345         If .Invent.DañoMagicoEqpSlot > 0 Then
350             If .Invent.Object(.Invent.DañoMagicoEqpSlot).ObjIndex > 0 Then
355                 .Invent.DañoMagicoEqpObjIndex = .Invent.Object(.Invent.DañoMagicoEqpSlot).ObjIndex

360                 If .flags.Muerto = 0 Then
365                     .Char.DM_Aura = ObjData(.Invent.DañoMagicoEqpObjIndex).CreaGRH
                    End If
                Else
370                 .Invent.DañoMagicoEqpSlot = 0
                End If
            End If
            
            ' RM
375         If .Invent.ResistenciaEqpSlot > 0 Then
380             If .Invent.Object(.Invent.ResistenciaEqpSlot).ObjIndex > 0 Then
385                 .Invent.ResistenciaEqpObjIndex = .Invent.Object(.Invent.ResistenciaEqpSlot).ObjIndex

390                 If .flags.Muerto = 0 Then
395                     .Char.RM_Aura = ObjData(.Invent.ResistenciaEqpObjIndex).CreaGRH
                    End If
                Else
400                 .Invent.ResistenciaEqpSlot = 0
                End If
            End If

405         If .Invent.MonturaSlot > 0 Then
410             If .Invent.Object(.Invent.MonturaSlot).ObjIndex > 0 Then
415                 .Invent.MonturaObjIndex = .Invent.Object(.Invent.MonturaSlot).ObjIndex
                Else
420                 .Invent.MonturaSlot = 0
                End If
            End If
        
425         If .Invent.HerramientaEqpSlot > 0 Then
430             If .Invent.Object(.Invent.HerramientaEqpSlot).ObjIndex Then
435                 .Invent.HerramientaEqpObjIndex = .Invent.Object(.Invent.HerramientaEqpSlot).ObjIndex
                Else
440                 .Invent.HerramientaEqpSlot = 0
                End If
            End If
        
445         If .Invent.NudilloSlot > 0 Then
450             If .Invent.Object(.Invent.NudilloSlot).ObjIndex > 0 Then
455                 .Invent.NudilloObjIndex = .Invent.Object(.Invent.NudilloSlot).ObjIndex

460                 If .flags.Muerto = 0 Then
465                     .Char.Arma_Aura = ObjData(.Invent.NudilloObjIndex).CreaGRH
                    End If
                Else
470                 .Invent.NudilloSlot = 0
                End If
            End If
        
475         If .Invent.MagicoSlot > 0 Then
480             If .Invent.Object(.Invent.MagicoSlot).ObjIndex Then
485                 .Invent.MagicoObjIndex = .Invent.Object(.Invent.MagicoSlot).ObjIndex

490                 If .flags.Muerto = 0 Then
495                     .Char.Otra_Aura = ObjData(.Invent.MagicoObjIndex).CreaGRH
                    End If
                Else
500                 .Invent.MagicoSlot = 0
                End If
            End If
            
505         If .Invent.EscudoEqpSlot = 0 Then .Char.ShieldAnim = NingunEscudo
510         If .Invent.CascoEqpSlot = 0 Then .Char.CascoAnim = NingunCasco
515         If .Invent.WeaponEqpSlot = 0 And .Invent.NudilloSlot = 0 And .Invent.HerramientaEqpSlot = 0 Then .Char.WeaponAnim = NingunArma
            ' -----------------------------------------------------------------------
            '   FIN - INFORMACION INICIAL DEL PERSONAJE
            ' -----------------------------------------------------------------------
            
520         If Not ValidateChr(UserIndex) Then
525             Call WriteShowMessageBox(UserIndex, "Error en el personaje. Comuniquese con el staff.")
530             Call CloseSocket(UserIndex)
                Exit Function

            End If
            
535         .flags.SeguroParty = True
540         .flags.SeguroClan = True
545         .flags.SeguroResu = True
        
550         .CurrentInventorySlots = getMaxInventorySlots(UserIndex)
        
555         Call WriteInventoryUnlockSlots(UserIndex)
        
560         Call LoadUserIntervals(UserIndex)
565         Call WriteIntervals(UserIndex)
        
570         Call UpdateUserInv(True, UserIndex, 0)
575         Call UpdateUserHechizos(True, UserIndex, 0)
        
580         Call EnviarLlaves(UserIndex)

590         If .flags.Paralizado Then Call WriteParalizeOK(UserIndex)
        
595         If .flags.Inmovilizado Then Call WriteInmovilizaOK(UserIndex)

            ''
            'TODO : Feo, esto tiene que ser parche cliente
600         If .flags.Estupidez = 0 Then
605             Call WriteDumbNoMore(UserIndex)
            End If
        
            'Ladder Inmunidad
610         .flags.Inmunidad = 1
615         .Counters.TiempoDeInmunidad = IntervaloPuedeSerAtacado
            'Ladder Inmunidad
        
            'Mapa válido
620         If Not MapaValido(.Pos.Map) Then
625             Call WriteErrorMsg(UserIndex, "EL PJ se encuenta en un mapa invalido.")
630             Call CloseSocket(UserIndex)
                Exit Function

            End If
        
            'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
            'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martin Sotuyo Dodero (Maraxus)
635         If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex <> 0 Or MapData(.Pos.Map, .Pos.X, .Pos.Y).NpcIndex <> 0 Then

                Dim FoundPlace As Boolean
                Dim esAgua     As Boolean
                Dim tX         As Long
                Dim tY         As Long
        
640             FoundPlace = False
645             esAgua = (MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked And FLAG_AGUA) <> 0

650             For tY = .Pos.Y - 1 To .Pos.Y + 1
655                 For tX = .Pos.X - 1 To .Pos.X + 1

660                     If esAgua Then

                            'reviso que sea pos legal en agua, que no haya User ni NPC para poder loguear.
665                         If LegalPos(.Pos.Map, tX, tY, True, False, False, False, False) Then
670                             FoundPlace = True
                                Exit For

                            End If

                        Else

                            'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
675                         If LegalPos(.Pos.Map, tX, tY, False, False, False, False, False) Then
680                             FoundPlace = True
                                Exit For

                            End If

                        End If

685                 Next tX
            
690                 If FoundPlace Then Exit For
695             Next tY
        
700             If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi

705                 .Pos.X = tX
710                 .Pos.Y = tY

                Else

                    'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
715                 If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex <> 0 Then

                        'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
720                     If UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu > 0 Then

                            'Le avisamos al que estaba comerciando que se tuvo que ir.
725                         If UserList(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu).flags.UserLogged Then
730                             Call FinComerciarUsu(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu)
735                             Call WriteConsoleMsg(UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_WARNING)

                            End If

                            'Lo sacamos.
740                         If UserList(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex).flags.UserLogged Then
745                             Call FinComerciarUsu(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex)
750                             Call WriteErrorMsg(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconectate...")

                            End If

                        End If
                
755                     Call CloseSocket(MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex)

                    End If

                End If

            End If
        
            'If in the water, and has a boat, equip it!
760         If .Invent.BarcoObjIndex > 0 And (MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked And FLAG_AGUA) <> 0 Then
765             .flags.Navegando = 1
770             Call EquiparBarco(UserIndex)

            End If
            
775         If .Invent.MagicoObjIndex <> 0 Then
780             If ObjData(.Invent.MagicoObjIndex).EfectoMagico = 11 Then .flags.Paraliza = 1
            End If

785         Call WriteUserIndexInServer(UserIndex) 'Enviamos el User index
790         .flags.NecesitaOxigeno = RequiereOxigeno(.Pos.Map)
        
795         Call WriteHora(UserIndex)
800         Call WriteChangeMap(UserIndex, .Pos.Map) 'Carga el mapa
            
805         Select Case .flags.Privilegios
            
                Case PlayerType.Admin
810                 .flags.ChatColor = RGB(217, 164, 32)
                
815             Case PlayerType.Dios
820                 .flags.ChatColor = RGB(217, 164, 32)
                    
825             Case PlayerType.SemiDios
830                 .flags.ChatColor = RGB(2, 161, 38)
                    
835             Case PlayerType.Consejero
840                 .flags.ChatColor = RGB(2, 161, 38)
                
845             Case Else
850                 .flags.ChatColor = vbWhite
                
            End Select
            
            ' Jopi: Te saco de los mapas de retos (si logueas ahi) 324 372 389 390
855         If Not EsGM(UserIndex) And (.Pos.Map = 324 Or .Pos.Map = 372 Or .Pos.Map = 389 Or .Pos.Map = 390) Then
                
                ' Si tiene una posicion a la que volver, lo mando ahi
860             If MapaValido(.flags.ReturnPos.Map) And .flags.ReturnPos.X > 0 And .flags.ReturnPos.X <= XMaxMapSize And .flags.ReturnPos.Y > 0 And .flags.ReturnPos.Y <= YMaxMapSize Then
                    
865                 Call WarpToLegalPos(UserIndex, .flags.ReturnPos.Map, .flags.ReturnPos.X, .flags.ReturnPos.Y, True)
                
                Else ' Lo mando a su hogar
                    
870                 Call WarpToLegalPos(UserIndex, Ciudades(.Hogar).Map, Ciudades(.Hogar).X, Ciudades(.Hogar).Y, True)
                    
                End If
                
            End If
            
            ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
            #If ConUpTime Then
875             .LogOnTime = Now
            #End If
        
            'Crea  el personaje del usuario
880         Call MakeUserChar(True, .Pos.Map, UserIndex, .Pos.Map, .Pos.X, .Pos.Y, 1)

885         Call WriteUserCharIndexInServer(UserIndex)
890         Call ActualizarVelocidadDeUsuario(UserIndex)
        
895         If EsGM(UserIndex) Then Call DoAdminInvisible(UserIndex)
900         Call WriteUpdateUserStats(UserIndex)
905         Call WriteUpdateHungerAndThirst(UserIndex)
910         Call WriteUpdateDM(UserIndex)
915         Call WriteUpdateRM(UserIndex)
        
920         Call SendMOTD(UserIndex)
925         Call SetUserLogged(UserIndex)
        
            'Actualiza el Num de usuarios
930         NumUsers = NumUsers + 1
935         .flags.UserLogged = True
940         .Counters.LastSave = GetTickCount
        
945         MapInfo(.Pos.Map).NumUsers = MapInfo(.Pos.Map).NumUsers + 1
        
950         If .Stats.SkillPts > 0 Then
955             Call WriteSendSkills(UserIndex)
960             Call WriteLevelUp(UserIndex, .Stats.SkillPts)

            End If
        
965         If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers
        
970         If NumUsers > RecordUsuarios Then
975             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultáneamente: " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
980             RecordUsuarios = NumUsers
985             Call SetRecordUsersDatabase(RecordUsuarios)

            End If

990         Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageOnlineUser(NumUsers))

995         Call WriteFYA(UserIndex)
1000         Call WriteBindKeys(UserIndex)
        
1005         If .NroMascotas > 0 And MapInfo(.Pos.Map).Seguro = 0 And .flags.MascotasGuardadas = 0 Then
                 Dim i As Integer

1010             For i = 1 To MAXMASCOTAS

1015                If .MascotasType(i) > 0 Then
1020                    .MascotasIndex(i) = SpawnNpc(.MascotasType(i), .Pos, False, False, False, UserIndex)
                    
1025                    If .MascotasIndex(i) > 0 Then
1030                        NpcList(.MascotasIndex(i)).MaestroUser = UserIndex
1035                        Call FollowAmo(.MascotasIndex(i))
                        
                         Else
1040                        .MascotasIndex(i) = 0

                         End If
                     End If

1045            Next i

             End If
        
1050        If .flags.Navegando = 1 Then
1055            Call WriteNavigateToggle(UserIndex)
1060            Call EquiparBarco(UserIndex)

             End If
        
1065        If .flags.Montado = 1 Then
1070            Call WriteEquiteToggle(UserIndex)

             End If

1075        Call ActualizarVelocidadDeUsuario(UserIndex)
        
1080        If .GuildIndex > 0 Then

                 'welcome to the show baby...
1085            If Not modGuilds.m_ConectarMiembroAClan(UserIndex, .GuildIndex) Then
1090                Call WriteConsoleMsg(UserIndex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
                 End If

             End If
        
1095        tStr = modGuilds.a_ObtenerRechazoDeChar(.Name)
    
1100        If LenB(tStr) <> 0 Then
1105            Call WriteShowMessageBox(UserIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
             End If

1110        If Lloviendo Then Call WriteRainToggle(UserIndex)
        
1115        If ServidorNublado Then Call WriteNubesToggle(UserIndex)

1120        Call WriteLoggedMessage(UserIndex)
        
1125        If .Stats.ELV = 1 Then
1130            Call WriteConsoleMsg(UserIndex, "¡Bienvenido a las tierras de AO20! ¡" & .Name & " que tengas buen viaje y mucha suerte!", FontTypeNames.FONTTYPE_GUILD)

1135        ElseIf .Stats.ELV < 14 Then
1140            Call WriteConsoleMsg(UserIndex, "¡Bienvenido de nuevo " & .Name & "! Actualmente estas en el nivel " & .Stats.ELV & " en " & DarNameMapa(.Pos.Map) & ", ¡buen viaje y mucha suerte!", FontTypeNames.FONTTYPE_GUILD)

             End If

1145        If Status(UserIndex) = Criminal Or Status(UserIndex) = e_Facciones.Caos Then
1150            Call WriteSafeModeOff(UserIndex)
1155            .flags.Seguro = False

             Else
1160            .flags.Seguro = True
1165            Call WriteSafeModeOn(UserIndex)

             End If
        
1170        If LenB(.MENSAJEINFORMACION) > 0 Then
                 Dim Lines() As String
1175            Lines = Split(.MENSAJEINFORMACION, vbNewLine)

1180            For i = 0 To UBound(Lines)

1185                If LenB(Lines(i)) > 0 Then
1190                    Call WriteConsoleMsg(UserIndex, Lines(i), FontTypeNames.FONTTYPE_New_DONADOR)
                     End If

                 Next

1195            .MENSAJEINFORMACION = vbNullString
             End If

1215        If EventoActivo Then
1220            Call WriteConsoleMsg(UserIndex, PublicidadEvento & ". Tiempo restante: " & TiempoRestanteEvento & " minuto(s).", FontTypeNames.FONTTYPE_New_Eventos)
             End If
        
1225        Call WriteContadores(UserIndex)
1230        Call WriteOxigeno(UserIndex)
    
         End With

         Exit Function

Complete_ConnectUser_Err:
1235    Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ConnectUser_Complete", Erl)

End Function

Sub ActStats(ByVal VictimIndex As Integer, ByVal AttackerIndex As Integer)
        
        On Error GoTo ActStats_Err
        

        Dim DaExp       As Integer

        Dim EraCriminal As Byte
    
100     DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)
    
102     If UserList(AttackerIndex).Stats.ELV < STAT_MAXELV Then
104         UserList(AttackerIndex).Stats.Exp = UserList(AttackerIndex).Stats.Exp + DaExp

106         If UserList(AttackerIndex).Stats.Exp > MAXEXP Then UserList(AttackerIndex).Stats.Exp = MAXEXP

108         Call WriteUpdateExp(AttackerIndex)
110         Call CheckUserLevel(AttackerIndex)

        End If
    
        'Lo mata
        'Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).name & "!", FontTypeNames.FONTTYPE_FIGHT)
    
112     Call WriteLocaleMsg(AttackerIndex, "184", FontTypeNames.FONTTYPE_FIGHT, UserList(VictimIndex).Name)
114     Call WriteLocaleMsg(AttackerIndex, "140", FontTypeNames.FONTTYPE_EXP, DaExp)
          
        'Call WriteConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
116     Call WriteLocaleMsg(VictimIndex, "185", FontTypeNames.FONTTYPE_FIGHT, UserList(AttackerIndex).Name)
    
118     If TriggerZonaPelea(VictimIndex, AttackerIndex) <> TRIGGER6_PERMITE Then
120         EraCriminal = Status(AttackerIndex)
        
122         If EraCriminal = 2 And Status(AttackerIndex) < 2 Then
124             Call RefreshCharStatus(AttackerIndex)
126         ElseIf EraCriminal < 2 And Status(AttackerIndex) = 2 Then
128             Call RefreshCharStatus(AttackerIndex)

            End If

        End If
    
130     Call UserDie(VictimIndex)
        
132     If UserList(AttackerIndex).Stats.UsuariosMatados < MAXUSERMATADOS Then
134         UserList(AttackerIndex).Stats.UsuariosMatados = UserList(AttackerIndex).Stats.UsuariosMatados + 1
        End If
        
        Exit Sub

ActStats_Err:
136     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ActStats", Erl)

        
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
            
114             Call WriteUpdateHungerAndThirst(UserIndex)
            End If
        
116         Call WriteUpdateHP(UserIndex)
            
118         If .flags.Navegando = 1 Then
120             Call EquiparBarco(UserIndex)
            Else

122             .Char.Head = .OrigChar.Head
    
124             If .Invent.CascoEqpObjIndex > 0 Then
126                 .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                End If
    
128             If .Invent.EscudoEqpObjIndex > 0 Then
130                 .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
    
                End If
    
132             If .Invent.WeaponEqpObjIndex > 0 Then
134                 .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
        
136                 If ObjData(.Invent.WeaponEqpObjIndex).CreaGRH <> "" Then
138                     .Char.Arma_Aura = ObjData(.Invent.WeaponEqpObjIndex).CreaGRH
140                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
    
                    End If
            
                End If
    
142             If .Invent.ArmourEqpObjIndex > 0 Then
144                 .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
        
146                 If ObjData(.Invent.ArmourEqpObjIndex).CreaGRH <> "" Then
148                     .Char.Body_Aura = ObjData(.Invent.ArmourEqpObjIndex).CreaGRH
150                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Body_Aura, False, 2))
    
                    End If
    
                Else
152                 Call DarCuerpoDesnudo(UserIndex)
            
                End If
    
154             If .Invent.EscudoEqpObjIndex > 0 Then
156                 .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
    
158                 If ObjData(.Invent.EscudoEqpObjIndex).CreaGRH <> "" Then
160                     .Char.Escudo_Aura = ObjData(.Invent.EscudoEqpObjIndex).CreaGRH
162                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Escudo_Aura, False, 3))
    
                    End If
            
                End If
    
164             If .Invent.CascoEqpObjIndex > 0 Then
166                 .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
    
168                 If ObjData(.Invent.CascoEqpObjIndex).CreaGRH <> "" Then
170                     .Char.Head_Aura = ObjData(.Invent.CascoEqpObjIndex).CreaGRH
172                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Head_Aura, False, 4))
    
                    End If
            
                End If
    
174             If .Invent.MagicoObjIndex > 0 Then
176                 If ObjData(.Invent.MagicoObjIndex).CreaGRH <> "" Then
178                     .Char.Otra_Aura = ObjData(.Invent.MagicoObjIndex).CreaGRH
180                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Otra_Aura, False, 5))
    
                    End If
    
                End If
    
182             If .Invent.NudilloObjIndex > 0 Then
184                 If ObjData(.Invent.NudilloObjIndex).CreaGRH <> "" Then
186                     .Char.Arma_Aura = ObjData(.Invent.NudilloObjIndex).CreaGRH
188                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.Arma_Aura, False, 1))
    
                    End If
                End If
                
190             If .Invent.DañoMagicoEqpObjIndex > 0 Then
192                 If ObjData(.Invent.DañoMagicoEqpObjIndex).CreaGRH <> "" Then
194                     .Char.DM_Aura = ObjData(.Invent.DañoMagicoEqpObjIndex).CreaGRH
196                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.DM_Aura, False, 6))
                    End If
                End If
                
198             If .Invent.ResistenciaEqpObjIndex > 0 Then
200                 If ObjData(.Invent.ResistenciaEqpObjIndex).CreaGRH <> "" Then
202                     .Char.RM_Aura = ObjData(.Invent.ResistenciaEqpObjIndex).CreaGRH
204                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAuraToChar(.Char.CharIndex, .Char.RM_Aura, False, 7))
                    End If
                End If
    
            End If
    
206         Call ActualizarVelocidadDeUsuario(UserIndex)
208         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

        End With
        
        Exit Sub

RevivirUsuario_Err:
210     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.RevivirUsuario", Erl)

        
End Sub

Sub ChangeUserChar(ByVal UserIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
        
        On Error GoTo ChangeUserChar_Err
        

100     With UserList(UserIndex).Char
102         .Body = Body
104         .Head = Head
106         .Heading = Heading
108         .WeaponAnim = Arma
110         .ShieldAnim = Escudo
112         .CascoAnim = Casco

        End With
    
114     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(Body, Head, Heading, UserList(UserIndex).Char.CharIndex, Arma, Escudo, UserList(UserIndex).Char.FX, UserList(UserIndex).Char.loops, Casco, False, UserList(UserIndex).flags.Navegando))

        
        Exit Sub

ChangeUserChar_Err:
116     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ChangeUserChar", Erl)

        
End Sub

Sub EraseUserChar(ByVal UserIndex As Integer, ByVal Desvanecer As Boolean)

        On Error GoTo ErrorHandler

        Dim Error As String
   
100     Error = "1"

102     If UserList(UserIndex).Char.CharIndex = 0 Then Exit Sub
   
104     CharList(UserList(UserIndex).Char.CharIndex) = 0
    
106     If UserList(UserIndex).Char.CharIndex = LastChar Then

108         Do Until CharList(LastChar) > 0
110             LastChar = LastChar - 1

112             If LastChar <= 1 Then Exit Do
            Loop

        End If

114     Error = "2"
    
        'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
116     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(UserList(UserIndex).Char.CharIndex, Desvanecer))
118     Error = "3"
120     Call QuitarUser(UserIndex, UserList(UserIndex).Pos.Map)
122     Error = "4"
    
124     MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).UserIndex = 0
126     Error = "5"
128     UserList(UserIndex).Char.CharIndex = 0
    
130     NumChars = NumChars - 1
132     Error = "6"
        Exit Sub
    
ErrorHandler:
134     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.EraseUserChar", Erl)

End Sub

Sub RefreshCharStatus(ByVal UserIndex As Integer)
        
        On Error GoTo RefreshCharStatus_Err
        

        '*************************************************
        'Author: Tararira
        'Last modified: 6/04/2007
        'Refreshes the status and tag of UserIndex.
        '*************************************************
        Dim klan As String, Name As String

100     If UserList(UserIndex).showName Then

102         If UserList(UserIndex).flags.Mimetizado = e_EstadoMimetismo.Desactivado Then

104             If UserList(UserIndex).GuildIndex > 0 Then
106                 klan = modGuilds.GuildName(UserList(UserIndex).GuildIndex)
108                 klan = " <" & klan & ">"
                End If
            
110             Name = UserList(UserIndex).Name & klan

            Else
112             Name = UserList(UserIndex).NameMimetizado
            End If
            
114         If UserList(UserIndex).clase = eClass.Pirat Then
116             If UserList(UserIndex).flags.Oculto = 1 Then
118                 Name = vbNullString
                End If
            End If
            
        End If
    
120     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, UserList(UserIndex).Faccion.Status, Name))

        
        Exit Sub

RefreshCharStatus_Err:
122     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.RefreshCharStatus", Erl)

        
End Sub

Sub MakeUserChar(ByVal toMap As Boolean, _
                 ByVal sndIndex As Integer, _
                 ByVal UserIndex As Integer, _
                 ByVal Map As Integer, _
                 ByVal X As Integer, _
                 ByVal Y As Integer, _
                 Optional ByVal appear As Byte = 0)

        On Error GoTo HayError

        Dim CharIndex As Integer

        Dim TempName  As String
    
100     If InMapBounds(Map, X, Y) Then
        
102         With UserList(UserIndex)
        
                'If needed make a new character in list
104             If .Char.CharIndex = 0 Then
106                 CharIndex = NextOpenCharIndex
108                 .Char.CharIndex = CharIndex
110                 CharList(CharIndex) = UserIndex
                End If

                'Place character on map if needed
112             If toMap Then MapData(Map, X, Y).UserIndex = UserIndex

                'Send make character command to clients
                Dim klan       As String
                Dim clan_nivel As Byte

114             If Not toMap Then
                
116                 If .showName Then
118                     If .flags.Mimetizado = e_EstadoMimetismo.Desactivado Then
120                         If .GuildIndex > 0 Then
                    
122                             klan = modGuilds.GuildName(.GuildIndex)
124                             clan_nivel = modGuilds.NivelDeClan(.GuildIndex)
126                             TempName = .Name & " <" & klan & ">"
                    
                            Else
                        
128                             klan = vbNullString
130                             clan_nivel = 0
                            
132                             If .flags.EnConsulta Then
                                
134                                 TempName = .Name & " [CONSULTA]"
                                
                                Else
                            
136                                 TempName = .Name
                            
                                End If
                            
                            End If
                        Else
138                         TempName = .NameMimetizado
                        End If
                    End If

140                 Call WriteCharacterCreate(sndIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.CharIndex, X, Y, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, TempName, .Faccion.Status, .flags.Privilegios, .Char.ParticulaFx, .Char.Head_Aura, .Char.Arma_Aura, .Char.Body_Aura, .Char.DM_Aura, .Char.RM_Aura, .Char.Otra_Aura, .Char.Escudo_Aura, .Char.speeding, 0, appear, .Grupo.Lider, .GuildIndex, clan_nivel, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMAN, .Stats.MaxMAN, 0, False, .flags.Navegando)
                                         
                Else
            
                    'Hide the name and clan - set privs as normal user
142                 Call AgregarUser(UserIndex, .Pos.Map, appear)
                
                End If
            
            End With
        
        End If

        Exit Sub

HayError:
        
        Dim Desc As String
144         Desc = Err.Description & vbNewLine & _
                    " Usuario: " & UserList(UserIndex).Name & vbNewLine & _
                    "Pos: " & Map & "-" & X & "-" & Y
            
146     Call TraceError(Err.Number, Err.Description, "Usuarios.MakeUserChar", Erl())
        
148     Call CloseSocket(UserIndex)

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
        '17/12/2020 WyroX: Distribución normal de las vidas
        '*************************************************

        On Error GoTo ErrHandler

        Dim Pts              As Integer

        Dim AumentoHIT       As Integer

        Dim AumentoMANA      As Integer

        Dim AumentoSta       As Integer

        Dim AumentoHP        As Integer

        Dim WasNewbie        As Boolean

        Dim Promedio         As Double
        
        Dim PromedioObjetivo As Double
        
        Dim PromedioUser     As Double

        Dim aux              As Integer
    
        Dim PasoDeNivel      As Boolean
        Dim experienceToLevelUp As Long

        ' Randomizo las vidas
100     Randomize Time
    
102     With UserList(UserIndex)

104         WasNewbie = EsNewbie(UserIndex)
106         experienceToLevelUp = ExpLevelUp(.Stats.ELV)
        
108         Do While .Stats.Exp >= experienceToLevelUp And .Stats.ELV < STAT_MAXELV
            
                'Store it!
                'Call Statistics.UserLevelUp(UserIndex)

110             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 106, 0))
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_NIVEL, .Pos.X, .Pos.Y))
114             Call WriteLocaleMsg(UserIndex, "186", FontTypeNames.FONTTYPE_INFO)
            
116             .Stats.Exp = .Stats.Exp - experienceToLevelUp
                
118             Pts = Pts + 5
            
                ' Calculo subida de vida by WyroX
                ' Obtengo el promedio según clase y constitución
120             PromedioObjetivo = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
                ' Obtengo el promedio actual del user
122             PromedioUser = CalcularPromedioVida(UserIndex)
                ' Lo modifico para compensar si está muy bajo o muy alto
124             Promedio = PromedioObjetivo + (PromedioObjetivo - PromedioUser) * DesbalancePromedioVidas
                ' Obtengo un entero al azar con más tendencia al promedio
126             AumentoHP = RandomIntBiased(PromedioObjetivo - RangoVidas, PromedioObjetivo + RangoVidas, Promedio, InfluenciaPromedioVidas)

                ' WyroX: Aumento del resto de stats
128             AumentoSta = ModClase(.clase).AumentoSta
130             AumentoMANA = ModClase(.clase).MultMana * .Stats.UserAtributos(eAtributos.Inteligencia)
132             AumentoHIT = IIf(.Stats.ELV < 36, ModClase(.clase).HitPre36, ModClase(.clase).HitPost36)

134             .Stats.ELV = .Stats.ELV + 1
136             experienceToLevelUp = ExpLevelUp(.Stats.ELV)
                
                'Actualizamos HitPoints
138             .Stats.MaxHp = .Stats.MaxHp + AumentoHP

140             If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
                'Actualizamos Stamina
142             .Stats.MaxSta = .Stats.MaxSta + AumentoSta

144             If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA
                'Actualizamos Mana
146             .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA

148             If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN

                'Actualizamos Golpe Máximo
150             .Stats.MaxHit = .Stats.MaxHit + AumentoHIT
            
                'Actualizamos Golpe Mínimo
152             .Stats.MinHIT = .Stats.MinHIT + AumentoHIT
        
                'Notificamos al user
154             If AumentoHP > 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
156                 Call WriteLocaleMsg(UserIndex, "197", FontTypeNames.FONTTYPE_INFO, AumentoHP)

                End If

158             If AumentoSta > 0 Then
                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoSTA & " puntos de vitalidad.", FontTypeNames.FONTTYPE_INFO)
160                 Call WriteLocaleMsg(UserIndex, "198", FontTypeNames.FONTTYPE_INFO, AumentoSta)

                End If

162             If AumentoMANA > 0 Then
164                 Call WriteLocaleMsg(UserIndex, "199", FontTypeNames.FONTTYPE_INFO, AumentoMANA)

                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & AumentoMANA & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
                End If

166             If AumentoHIT > 0 Then
168                 Call WriteLocaleMsg(UserIndex, "200", FontTypeNames.FONTTYPE_INFO, AumentoHIT)

                    'Call WriteConsoleMsg(UserIndex, "Tu golpe aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
                End If

170             PasoDeNivel = True
             
                ' Call LogDesarrollo(.name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)

172             .Stats.MinHp = .Stats.MaxHp
            
                ' Call UpdateUserInv(True, UserIndex, 0)
            
174             If OroPorNivel > 0 Then
176                 If EsNewbie(UserIndex) Then
                        Dim OroRecompenza As Long
    
178                     OroRecompenza = OroPorNivel * .Stats.ELV * OroMult
180                     .Stats.GLD = .Stats.GLD + OroRecompenza
                        'Call WriteConsoleMsg(UserIndex, "Has ganado " & OroRecompenza & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
182                     Call WriteLocaleMsg(UserIndex, "29", FontTypeNames.FONTTYPE_INFO, PonerPuntos(OroRecompenza))
                    End If
                End If
            
184             If Not EsNewbie(UserIndex) And WasNewbie Then
        
186                 Call QuitarNewbieObj(UserIndex)
            
                End If
        
            Loop
        
188         If PasoDeNivel Then
190             If .Stats.ELV >= STAT_MAXELV Then .Stats.Exp = 0
        
192             Call UpdateUserInv(True, UserIndex, 0)
                'Call CheckearRecompesas(UserIndex, 3)
194             Call WriteUpdateUserStats(UserIndex)
            
196             If Pts > 0 Then
                
198                 .Stats.SkillPts = .Stats.SkillPts + Pts
200                 Call WriteLevelUp(UserIndex, .Stats.SkillPts)
202                 Call WriteLocaleMsg(UserIndex, "187", FontTypeNames.FONTTYPE_INFO, Pts)

                    'Call WriteConsoleMsg(UserIndex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
                End If
                
204             If .Stats.ELV >= MapInfo(.Pos.Map).MaxLevel And Not EsGM(UserIndex) Then
206                 If MapInfo(.Pos.Map).Salida.Map <> 0 Then
208                     Call WriteConsoleMsg(UserIndex, "Tu nivel no te permite seguir en el mapa.", FontTypeNames.FONTTYPE_INFO)
210                     Call WarpUserChar(UserIndex, MapInfo(.Pos.Map).Salida.Map, MapInfo(.Pos.Map).Salida.X, MapInfo(.Pos.Map).Salida.Y, True)
                    End If
                End If

            End If
    
        End With
    
        Exit Sub

ErrHandler:
212     Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.Description)

End Sub

Function MoveUserChar(ByVal UserIndex As Integer, ByVal nHeading As eHeading) As Boolean
        ' 20/01/2021 - WyroX: Lo convierto a función y saco los WritePosUpdate, ahora están en el paquete

        On Error GoTo MoveUserChar_Err

        Dim nPos         As WorldPos
        Dim nPosOriginal As WorldPos
        Dim nPosMuerto   As WorldPos
        Dim IndexMover As Integer
        Dim OppositeHeading As eHeading

100     With UserList(UserIndex)

102         nPos = .Pos
104         Call HeadtoPos(nHeading, nPos)

106         If Not LegalWalk(.Pos.Map, nPos.X, nPos.Y, nHeading, .flags.Navegando = 1, .flags.Navegando = 0, .flags.Montado) Then
                Exit Function
            End If

108         If .Accion.AccionPendiente = True Then
110             .Counters.TimerBarra = 0
112             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, .Accion.Particula, .Counters.TimerBarra, True))
114             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(.Char.CharIndex, .Counters.TimerBarra, Accion_Barra.CancelarAccion))
116             .Accion.AccionPendiente = False
118             .Accion.Particula = 0
120             .Accion.TipoAccion = Accion_Barra.CancelarAccion
122             .Accion.HechizoPendiente = 0
124             .Accion.RunaObj = 0
126             .Accion.ObjSlot = 0
128             .Accion.AccionPendiente = False
            End If

            'If .flags.Muerto = 0 Then
            '    If MapData(nPos.Map, nPos.X, nPos.Y).TileExit.Map <> 0 And .Counters.EnCombate > 0 Then
            '        Call WriteConsoleMsg(UserIndex, "Estás en combate, debes aguardar " & .Counters.EnCombate & " segundo(s) para escapar...", FontTypeNames.FONTTYPE_INFOBOLD)
            '        Exit Function
            '    End If
            'End If

            'Si no estoy solo en el mapa...
130         If MapInfo(.Pos.Map).NumUsers > 1 Then

                ' Intercambia posición si hay un casper o gm invisible
132             IndexMover = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
            
134             If IndexMover <> 0 Then
                    ' Sólo puedo patear caspers/gms invisibles si no es él un gm invisible
136                 If UserList(UserIndex).flags.AdminInvisible <> 0 Then Exit Function

138                 Call WritePosUpdate(IndexMover)

140                 OppositeHeading = InvertHeading(nHeading)
142                 Call HeadtoPos(OppositeHeading, UserList(IndexMover).Pos)
                
                    ' Si es un admin invisible, no se avisa a los demas clientes
144                 If UserList(IndexMover).flags.AdminInvisible = 0 Then
146                     Call SendData(SendTarget.ToPCAreaButIndex, IndexMover, PrepareMessageCharacterMove(UserList(IndexMover).Char.CharIndex, UserList(IndexMover).Pos.X, UserList(IndexMover).Pos.Y))
                    Else
148                     Call SendData(SendTarget.ToAdminAreaButIndex, IndexMover, PrepareMessageCharacterMove(UserList(IndexMover).Char.CharIndex, UserList(IndexMover).Pos.X, UserList(IndexMover).Pos.Y))
                    End If
                    
150                 Call WriteForceCharMove(IndexMover, OppositeHeading)
                
                    'Update map and char
152                 UserList(IndexMover).Char.Heading = OppositeHeading
154                 MapData(UserList(IndexMover).Pos.Map, UserList(IndexMover).Pos.X, UserList(IndexMover).Pos.Y).UserIndex = IndexMover
                
                    'Actualizamos las areas de ser necesario
156                 Call ModAreas.CheckUpdateNeededUser(IndexMover, OppositeHeading, 0)
                End If

158             If .flags.AdminInvisible = 0 Then
160                 Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))
                Else
162                 Call SendData(SendTarget.ToAdminAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))
                End If
            End If
        
            'Update map and user pos
164         If MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex Then
166             MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
            End If

168         .Pos = nPos
170         .Char.Heading = nHeading
172         MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
        
            'Actualizamos las áreas de ser necesario
174         Call ModAreas.CheckUpdateNeededUser(UserIndex, nHeading, 0)

176         If .Counters.Trabajando Then
178             Call WriteMacroTrabajoToggle(UserIndex, False)
            End If

180         If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
    
        End With
    
182     MoveUserChar = True
    
        Exit Function
    
MoveUserChar_Err:
184     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.MoveUserChar", Erl)

        
End Function

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading
        
        On Error GoTo InvertHeading_Err
    
        

        '*************************************************
        'Author: ZaMa
        'Last modified: 30/03/2009
        'Returns the heading opposite to the one passed by val.
        '*************************************************
100     Select Case nHeading

            Case eHeading.EAST
102             InvertHeading = eHeading.WEST

104         Case eHeading.WEST
106             InvertHeading = eHeading.EAST

108         Case eHeading.SOUTH
110             InvertHeading = eHeading.NORTH

112         Case eHeading.NORTH
114             InvertHeading = eHeading.SOUTH

        End Select

        
        Exit Function

InvertHeading_Err:
116     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.InvertHeading", Erl)

        
End Function

Sub ChangeUserInv(ByVal UserIndex As Integer, ByVal Slot As Byte, ByRef Object As UserOBJ)
        
        On Error GoTo ChangeUserInv_Err
        
100     UserList(UserIndex).Invent.Object(Slot) = Object
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
   
100     For LoopC = 1 To MaxUsers + 1

102         If LoopC > MaxUsers Then Exit For
104         If (Not UserList(LoopC).ConnIDValida And UserList(LoopC).flags.UserLogged = False) Then Exit For
106     Next LoopC
   
108     NextOpenUser = LoopC

        
        Exit Function

NextOpenUser_Err:
110     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.NextOpenUser", Erl)

        
End Function

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserStatsTxt_Err
        

        Dim GuildI As Integer

100     Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
102     Call WriteConsoleMsg(sendIndex, "Nivel: " & UserList(UserIndex).Stats.ELV & "  EXP: " & UserList(UserIndex).Stats.Exp & "/" & ExpLevelUp(UserList(UserIndex).Stats.ELV), FontTypeNames.FONTTYPE_INFO)
104     Call WriteConsoleMsg(sendIndex, "Salud: " & UserList(UserIndex).Stats.MinHp & "/" & UserList(UserIndex).Stats.MaxHp & "  Mana: " & UserList(UserIndex).Stats.MinMAN & "/" & UserList(UserIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(UserIndex).Stats.MinSta & "/" & UserList(UserIndex).Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
    
106     If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
108         Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHit & " (" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MaxHit & ")", FontTypeNames.FONTTYPE_INFO)
        Else
110         Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(UserIndex).Stats.MinHIT & "/" & UserList(UserIndex).Stats.MaxHit, FontTypeNames.FONTTYPE_INFO)

        End If
    
112     If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
114         If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
116             Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef + ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef + ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            Else
118             Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)

            End If

        Else
120         Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)

        End If
    
122     If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
124         Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: " & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
126         Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)

        End If
    
128     GuildI = UserList(UserIndex).GuildIndex

130     If GuildI > 0 Then
132         Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)

134         If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(UserList(sendIndex).Name) Then
136             Call WriteConsoleMsg(sendIndex, "Status: Líder", FontTypeNames.FONTTYPE_INFO)

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
144         Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
146         Call WriteConsoleMsg(sendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
        #End If

148     Call WriteConsoleMsg(sendIndex, "Oro: " & UserList(UserIndex).Stats.GLD & "  Posicion: " & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y & " en mapa " & UserList(UserIndex).Pos.Map, FontTypeNames.FONTTYPE_INFO)
150     Call WriteConsoleMsg(sendIndex, "Dados: " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion) & ", " & UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma), FontTypeNames.FONTTYPE_INFO)
152     Call WriteConsoleMsg(sendIndex, "Veces que Moriste: " & UserList(UserIndex).flags.VecesQueMoriste, FontTypeNames.FONTTYPE_INFO)

        Exit Sub

SendUserStatsTxt_Err:
154     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserStatsTxt", Erl)

        
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserMiniStatsTxt_Err
        

        '*************************************************
        'Author: Unknown
        'Last modified: 23/01/2007
        'Shows the users Stats when the user is online.
        '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
        '*************************************************
100     With UserList(UserIndex)
102         Call WriteConsoleMsg(sendIndex, "Pj: " & .Name, FontTypeNames.FONTTYPE_INFO)
104         Call WriteConsoleMsg(sendIndex, "Ciudadanos Matados: " & .Faccion.ciudadanosMatados & " Criminales Matados: " & .Faccion.CriminalesMatados & " UsuariosMatados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
106         Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
108         Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.clase), FontTypeNames.FONTTYPE_INFO)
110         Call WriteConsoleMsg(sendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)

112         If .GuildIndex > 0 Then
114             Call WriteConsoleMsg(sendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)

            End If

116         Call WriteConsoleMsg(sendIndex, "Oro en billetera: " & .Stats.GLD, FontTypeNames.FONTTYPE_INFO)
118         Call WriteConsoleMsg(sendIndex, "Oro en banco: " & .Stats.Banco, FontTypeNames.FONTTYPE_INFO)

        End With

        
        Exit Sub

SendUserMiniStatsTxt_Err:
126     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserMiniStatsTxt", Erl)

        
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserInvTxt_Err
    
        

        

        Dim j As Long
    
100     Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
102     Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(UserIndex).Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
    
104     For j = 1 To UserList(UserIndex).CurrentInventorySlots

106         If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
108             Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).Invent.Object(j).amount, FontTypeNames.FONTTYPE_INFO)

            End If

110     Next j

        
        Exit Sub

SendUserInvTxt_Err:
112     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserInvTxt", Erl)

        
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo SendUserSkillsTxt_Err
    
        

        

        Dim j As Integer

100     Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)

102     For j = 1 To NUMSKILLS
104         Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(UserIndex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
        Next
106     Call WriteConsoleMsg(sendIndex, " SkillLibres:" & UserList(UserIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)

        
        Exit Sub

SendUserSkillsTxt_Err:
108     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SendUserSkillsTxt", Erl)

        
End Sub

Function DameUserIndexConNombre(ByVal nombre As String) As Integer
        
        On Error GoTo DameUserIndexConNombre_Err
        

        Dim LoopC As Integer
  
100     LoopC = 1
  
102     nombre = UCase$(nombre)

104     Do Until UCase$(UserList(LoopC).Name) = nombre

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

Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        On Error GoTo NPCAtacado_Err
        
        ' WyroX: El usuario pierde la protección
100     UserList(UserIndex).Counters.TiempoDeInmunidad = 0
102     UserList(UserIndex).flags.Inmunidad = 0

        'Guardamos el usuario que ataco el npc.
104     If NpcList(NpcIndex).Movement <> Estatico And NpcList(NpcIndex).flags.AttackedFirstBy = vbNullString Then
106         NpcList(NpcIndex).Target = UserIndex
108         NpcList(NpcIndex).Hostile = 1
110         NpcList(NpcIndex).flags.AttackedBy = UserList(UserIndex).Name
        End If

        'Npc que estabas atacando.
        Dim LastNpcHit As Integer

112     LastNpcHit = UserList(UserIndex).flags.NPCAtacado
        'Guarda el NPC que estas atacando ahora.
114     UserList(UserIndex).flags.NPCAtacado = NpcIndex

116     If NpcList(NpcIndex).flags.Faccion = Armada And Status(UserIndex) = e_Facciones.Ciudadano Then
118         Call VolverCriminal(UserIndex)
        End If
        
120     If NpcList(NpcIndex).MaestroUser > 0 And NpcList(NpcIndex).MaestroUser <> UserIndex Then
122         Call AllMascotasAtacanUser(UserIndex, NpcList(NpcIndex).MaestroUser)
        End If

124     Call AllMascotasAtacanNPC(NpcIndex, UserIndex)
        
        Exit Sub

NPCAtacado_Err:
126     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.NPCAtacado", Erl)

        
End Sub

Sub SubirSkill(ByVal UserIndex As Integer, ByVal Skill As Integer)
        On Error GoTo SubirSkill_Err

        Dim Lvl As Integer, maxPermitido As Integer
100         Lvl = UserList(UserIndex).Stats.ELV

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

112     If UserList(UserIndex).flags.Hambre = 0 And UserList(UserIndex).flags.Sed = 0 Then

            Dim Aumenta As Integer
            Dim Prob    As Integer
            Dim Menor   As Byte
114         Menor = 10
             
116         If Lvl <= 3 Then
118             Prob = 25
120         ElseIf Lvl > 3 And Lvl < 6 Then
122             Prob = 27
124         ElseIf Lvl >= 6 And Lvl < 10 Then
126             Prob = 30
128         ElseIf Lvl >= 10 And Lvl < 20 Then
130             Prob = 33
            Else
132             Prob = 38
            End If
             
134         Aumenta = RandomNumber(1, Prob * DificultadSubirSkill)
             
136         If UserList(UserIndex).flags.PendienteDelExperto = 1 Then
138             Menor = 15
            End If
        
140         If Aumenta < Menor Then
142             UserList(UserIndex).Stats.UserSkills(Skill) = UserList(UserIndex).Stats.UserSkills(Skill) + 1
    
144             Call WriteConsoleMsg(UserIndex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(UserIndex).Stats.UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
            
                Dim BonusExp As Long
146             BonusExp = 50& * ExpMult
        
                Call WriteConsoleMsg(UserIndex, "¡Has ganado " & BonusExp & " puntos de experiencia!", FontTypeNames.FONTTYPE_INFOIAO)
                
152             If UserList(UserIndex).Stats.ELV < STAT_MAXELV Then
154                 UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + BonusExp
156                 If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
                    
                    UserList(UserIndex).flags.ModificoSkills = True
                    
158                 If UserList(UserIndex).ChatCombate = 1 Then
160                     Call WriteLocaleMsg(UserIndex, "140", FontTypeNames.FONTTYPE_EXP, BonusExp)
                    End If
                
162                 Call WriteUpdateExp(UserIndex)
164                 Call CheckUserLevel(UserIndex)

                End If

            End If

        End If

        
        Exit Sub

SubirSkill_Err:
166     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SubirSkill", Erl)

        
End Sub

Public Sub SubirSkillDeArmaActual(ByVal UserIndex As Integer)
        On Error GoTo SubirSkillDeArmaActual_Err

100     With UserList(UserIndex)

102         If .Invent.WeaponEqpObjIndex > 0 Then
                ' Arma con proyectiles, subimos armas a distancia
104             If ObjData(.Invent.WeaponEqpObjIndex).Proyectil Then
106                 Call SubirSkill(UserIndex, eSkill.Proyectiles)

                ' Sino, subimos combate con armas
                Else
108                 Call SubirSkill(UserIndex, eSkill.Armas)
                End If

            ' Si no está usando un arma, subimos combate sin armas
            Else
110             Call SubirSkill(UserIndex, eSkill.Wrestling)
            End If

        End With

        Exit Sub

SubirSkillDeArmaActual_Err:
112         Call TraceError(Err.Number, Err.Description, "UsUaRiOs.SubirSkillDeArmaActual", Erl)


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
    
100     With UserList(UserIndex)
102         .Counters.Mimetismo = 0
104         .flags.Mimetizado = e_EstadoMimetismo.Desactivado
106         Call RefreshCharStatus(UserIndex)
    
            'Sonido
108         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(IIf(.genero = eGenero.Hombre, e_SoundIndex.MUERTE_HOMBRE, e_SoundIndex.MUERTE_MUJER), .Pos.X, .Pos.Y))
        
            'Quitar el dialogo del user muerto
110         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
112         .Stats.MinHp = 0
114         .Stats.MinSta = 0
116         .flags.AtacadoPorUser = 0

118         .flags.incinera = 0
120         .flags.Paraliza = 0
122         .flags.Envenena = 0
124         .flags.Estupidiza = 0
126         .flags.Muerto = 1
128         .flags.Ahogandose = 0
            
130         Call WriteUpdateHP(UserIndex)
132         Call WriteUpdateSta(UserIndex)
        
134         aN = .flags.AtacadoPorNpc
    
136         If aN > 0 Then
138             NpcList(aN).Movement = NpcList(aN).flags.OldMovement
140             NpcList(aN).Hostile = NpcList(aN).flags.OldHostil
142             NpcList(aN).flags.AttackedBy = vbNullString
144             NpcList(aN).Target = 0
            End If
        
146         aN = .flags.NPCAtacado
    
148         If aN > 0 Then
150             If NpcList(aN).flags.AttackedFirstBy = .Name Then
152                 NpcList(aN).flags.AttackedFirstBy = vbNullString
                End If
            End If
    
154         .flags.AtacadoPorNpc = 0
156         .flags.NPCAtacado = 0
    
158         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger <> eTrigger.ZONAPELEA Then

160             If (.flags.Privilegios And PlayerType.user) <> 0 Then

162                 If .flags.PendienteDelSacrificio = 0 Then
                
164                     Call TirarTodosLosItems(UserIndex)
    
                    Else
                
                        Dim MiObj As obj

166                     MiObj.amount = 1
168                     MiObj.ObjIndex = PENDIENTE
170                     Call QuitarObjetos(PENDIENTE, 1, UserIndex)
172                     Call MakeObj(MiObj, .Pos.Map, .Pos.X, .Pos.Y)
174                     Call WriteConsoleMsg(UserIndex, "Has perdido tu pendiente del sacrificio.", FontTypeNames.FONTTYPE_INFO)

                    End If
    
                End If
    
            End If
        
176         .flags.CarroMineria = 0
   
            'desequipar montura
178         If .flags.Montado > 0 Then
180             Call DoMontar(UserIndex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)
            End If
        
            ' << Reseteamos los posibles FX sobre el personaje >>
182         If .Char.loops = INFINITE_LOOPS Then
184             .Char.FX = 0
186             .Char.loops = 0
    
            End If
        
188         .flags.VecesQueMoriste = .flags.VecesQueMoriste + 1
        
            ' << Restauramos los atributos >>
190         If .flags.TomoPocion Then
    
192             For i = 1 To 4
194                 .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
196             Next i
    
198             Call WriteFYA(UserIndex)
    
            End If
        
            '<< Cambiamos la apariencia del char >>
200         If .flags.Navegando = 0 Then
202             .Char.Body = iCuerpoMuerto
204             .Char.Head = 0
206             .Char.ShieldAnim = NingunEscudo
208             .Char.WeaponAnim = NingunArma
210             .Char.CascoAnim = NingunCasco
            Else
212             Call EquiparBarco(UserIndex)
            End If
            
214         Call ActualizarVelocidadDeUsuario(UserIndex)
216         Call LimpiarEstadosAlterados(UserIndex)
        
218         For i = 1 To MAXMASCOTAS
220             If .MascotasIndex(i) > 0 Then
222                 Call MuereNpc(.MascotasIndex(i), 0)
                End If
224         Next i
        
226         .NroMascotas = 0
        
            '<< Actualizamos clientes >>
228         Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco)

230         If MapInfo(.Pos.Map).Seguro = 0 Then
232             Call WriteConsoleMsg(UserIndex, "Escribe /HOGAR si deseas regresar rápido a tu hogar.", FontTypeNames.FONTTYPE_New_Naranja)
            End If
            
234         If .flags.EnReto Then
236             Call MuereEnReto(UserIndex)
            End If

        End With

        Exit Sub

ErrorHandler:
238        Call RegistrarError(Err.Number, Err.Description, "UsUaRiOs.UserDie", Erl)

End Sub

Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
            On Error GoTo ContarMuerte_Err


100         If EsNewbie(Muerto) Then Exit Sub
102         If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
104         If Abs(CInt(UserList(Muerto).Stats.ELV) - CInt(UserList(Atacante).Stats.ELV)) > 14 Then Exit Sub
106         If Status(Muerto) = 0 Or Status(Muerto) = 2 Then
108             If UserList(Atacante).flags.LastCrimMatado <> UserList(Muerto).Name Then
110                 UserList(Atacante).flags.LastCrimMatado = UserList(Muerto).Name

112                 If UserList(Atacante).Faccion.CriminalesMatados < MAXUSERMATADOS Then
114                     UserList(Atacante).Faccion.CriminalesMatados = UserList(Atacante).Faccion.CriminalesMatados + 1
                    End If
                End If

116         ElseIf Status(Muerto) = 1 Or Status(Muerto) = 3 Then

118             If UserList(Atacante).flags.LastCiudMatado <> UserList(Muerto).Name Then
120                 UserList(Atacante).flags.LastCiudMatado = UserList(Muerto).Name

122                 If UserList(Atacante).Faccion.ciudadanosMatados < MAXUSERMATADOS Then
124                     UserList(Atacante).Faccion.ciudadanosMatados = UserList(Atacante).Faccion.ciudadanosMatados + 1
                    End If

                End If

            End If

            Exit Sub

ContarMuerte_Err:
126         Call TraceError(Err.Number, Err.Description, "UsUaRiOs.ContarMuerte", Erl)


End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef obj As obj, ByRef Agua As Boolean, ByRef Tierra As Boolean, Optional ByVal InitialPos As Boolean = True)

        
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
        
100     hayobj = False
102     nPos.Map = Pos.Map

104     Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y, Agua, Tierra) Or hayobj
        
106         If LoopC > 15 Then
108             Notfound = True
                Exit Do

            End If
        
110         For tY = Pos.Y - LoopC To Pos.Y + LoopC
112             For tX = Pos.X - LoopC To Pos.X + LoopC
            
114                 If LegalPos(nPos.Map, tX, tY, Agua, Tierra) Then
                        'We continue if: a - the item is different from 0 and the dropped item or b - the Amount dropped + Amount in map exceeds MAX_INVENTORY_OBJS
116                     hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex <> obj.ObjIndex)

118                     If Not hayobj Then hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.amount + obj.amount > MAX_INVENTORY_OBJS)

120                     If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 And (InitialPos Or (tX <> Pos.X And tY <> Pos.Y)) Then
122                         nPos.X = tX
124                         nPos.Y = tY
126                         tX = Pos.X + LoopC
128                         tY = Pos.Y + LoopC

                        End If

                    End If
            
130             Next tX
132         Next tY
        
134         LoopC = LoopC + 1
        
        Loop
    
136     If Notfound = True Then
138         nPos.X = 0
140         nPos.Y = 0

        End If

        
        Exit Sub

Tilelibre_Err:
142     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.Tilelibre", Erl)

        
End Sub

Sub WarpToLegalPos(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal FX As Boolean = False, Optional ByVal AguaValida As Boolean = False)

        On Error GoTo WarpToLegalPos_Err

        Dim LoopC    As Integer

        Dim tX       As Integer

        Dim tY       As Integer

102     Do While True

104         If LoopC > 20 Then Exit Sub

108         For tY = Y - LoopC To Y + LoopC
110             For tX = X - LoopC To X + LoopC
            
112                 If LegalPos(Map, tX, tY, AguaValida, True, UserList(UserIndex).flags.Montado = 1, False, False) Then
                        If MapData(Map, tX, tY).trigger < 50 Then
114                         Call WarpUserChar(UserIndex, Map, tX, tY, FX)
                            Exit Sub
                        End If
                    End If
        
122             Next tX
124         Next tY
    
126         LoopC = LoopC + 1
    
        Loop

        Call WarpUserChar(UserIndex, Map, X, Y, FX)

        Exit Sub

WarpToLegalPos_Err:
132     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.WarpToLegalPos", Erl)

        
End Sub

Sub WarpUserChar(ByVal UserIndex As Integer, _
                 ByVal Map As Integer, _
                 ByVal X As Integer, _
                 ByVal Y As Integer, _
                 Optional ByVal FX As Boolean = False)
        
        On Error GoTo WarpUserChar_Err

        Dim OldMap As Integer
        Dim OldX   As Integer
        Dim OldY   As Integer
    
100     With UserList(UserIndex)

102         If .ComUsu.DestUsu > 0 Then

104             If UserList(.ComUsu.DestUsu).flags.UserLogged Then

106                 If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
108                     Call WriteConsoleMsg(.ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
110                     Call FinComerciarUsu(.ComUsu.DestUsu)

                    End If

                End If

            End If
    
            'Quitar el dialogo
112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
    
114         Call WriteRemoveAllDialogs(UserIndex)
    
116         OldMap = .Pos.Map
118         OldX = .Pos.X
120         OldY = .Pos.Y
    
122         Call EraseUserChar(UserIndex, True)
    
124         If OldMap <> Map Then
126             Call WriteChangeMap(UserIndex, Map)

128             If MapInfo(OldMap).Seguro = 1 And MapInfo(Map).Seguro = 0 And .Stats.ELV < 42 Then
130                 Call WriteConsoleMsg(UserIndex, "Estás saliendo de una zona segura, recuerda que aquí corres riesgo de ser atacado.", FontTypeNames.FONTTYPE_WARNING)

                End If
        
132             .flags.NecesitaOxigeno = RequiereOxigeno(Map)

134             If .flags.NecesitaOxigeno Then
136                 Call WriteContadores(UserIndex)
138                 Call WriteOxigeno(UserIndex)

140                 If .Counters.Oxigeno = 0 Then
142                     .flags.Ahogandose = 1

                    End If

                End If
            
144             If IntervaloPuedeSerAtacado > 0 Then
146                 .Counters.TiempoDeInmunidad = IntervaloPuedeSerAtacado
148                 .flags.Inmunidad = 1
                End If

150             If RequiereOxigeno(OldMap) = True And .flags.NecesitaOxigeno = False Then  'And .Stats.ELV < 35 Then
        
                    'Call WriteConsoleMsg(UserIndex, "Ya no necesitas oxigeno.", FontTypeNames.FONTTYPE_WARNING)
152                 Call WriteContadores(UserIndex)
154                 Call WriteOxigeno(UserIndex)

                End If
        
                'Update new Map Users
156             MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
        
                'Update old Map Users
158             MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1

160             If MapInfo(OldMap).NumUsers < 0 Then
162                 MapInfo(OldMap).NumUsers = 0

                End If

164             If .flags.Traveling = 1 Then
166                 .flags.Traveling = 0
168                 .Counters.goHome = 0
170                 Call WriteConsoleMsg(UserIndex, "El viaje ha terminado.", FontTypeNames.FONTTYPE_INFOBOLD)

                End If
   
            End If
    
172         .Pos.X = X
174         .Pos.Y = Y
176         .Pos.Map = Map

178         If .Grupo.EnGrupo = True Then
180             Call CompartirUbicacion(UserIndex)
            End If
    
182         If FX Then
184             Call MakeUserChar(True, Map, UserIndex, Map, X, Y, 1)
            Else
186             Call MakeUserChar(True, Map, UserIndex, Map, X, Y, 0)
            End If
    
188         Call WriteUserCharIndexInServer(UserIndex)
    
            'Seguis invisible al pasar de mapa
190         If (.flags.invisible = 1 Or .flags.Oculto = 1) And (Not .flags.AdminInvisible = 1) Then

                ' Si el mapa lo permite
192             If MapInfo(Map).SinInviOcul Then
            
194                 .flags.invisible = 0
196                 .flags.Oculto = 0
198                 .Counters.TiempoOculto = 0
                
200                 Call WriteConsoleMsg(UserIndex, "Una fuerza divina que vigila esta zona te ha vuelto visible.", FontTypeNames.FONTTYPE_INFO)
                
                Else
202                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))

                End If

            End If
    
            'Reparacion temporal del bug de particulas. 08/07/09 LADDER
204         If .flags.AdminInvisible = 0 Then
        
206             If FX Then 'FX
208                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
210                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))
                End If

            Else
212             Call SendData(ToIndex, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, True))

            End If
        
214         If .NroMascotas > 0 Then Call WarpMascotas(UserIndex)
    
216         If MapInfo(Map).zone = "DUNGEON" Or MapData(Map, X, Y).trigger >= 9 Then

218             If .flags.Montado > 0 Then
220                 Call DoMontar(UserIndex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)
                End If

            End If
    
        End With

        Exit Sub

WarpUserChar_Err:
222     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.WarpUserChar", Erl)


        
End Sub

Sub WarpFamiliar(ByVal UserIndex As Integer)
        
        On Error GoTo WarpFamiliar_Err
        

100     With UserList(UserIndex)

102         If .Familiar.Invocado = 1 Then
104             Call QuitarNPC(.Familiar.ID)
                ' If MapInfo(UserList(UserIndex).Pos.map).Pk = True Then
106             .Familiar.ID = SpawnNpc(.Familiar.NpcIndex, UserList(UserIndex).Pos, False, True)

                'Controlamos que se sumoneo OK
108             If .Familiar.ID = 0 Then
110                 Call WriteConsoleMsg(UserIndex, "No hay espacio aquí para tu mascota.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

112             Call CargarFamiliar(UserIndex)
            Else

                'Call WriteConsoleMsg(UserIndex, "No se permiten familiares en zona segura. " & .Familiar.Nombre & " te esperará afuera.", FontTypeNames.FONTTYPE_INFO)
            End If
    
        End With
            
        
        Exit Sub

WarpFamiliar_Err:
114     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.WarpFamiliar", Erl)

        
End Sub

Sub Cerrar_Usuario(ByVal UserIndex As Integer)

        On Error GoTo Cerrar_Usuario_Err
    
100     With UserList(UserIndex)

102         If .flags.UserLogged And Not .Counters.Saliendo Then
104             .Counters.Saliendo = True
106             .Counters.Salir = IntervaloCerrarConexion
            
108             If .flags.Traveling = 1 Then
110                 Call WriteConsoleMsg(UserIndex, "Se ha cancelado el viaje a casa", FontTypeNames.FONTTYPE_INFO)
112                 .flags.Traveling = 0
114                 .Counters.goHome = 0
                End If
            
116             Call WriteLocaleMsg(UserIndex, "203", FontTypeNames.FONTTYPE_INFO, .Counters.Salir)
            
118             If EsGM(UserIndex) Or MapInfo(.Pos.Map).Seguro = 1 Then
120                 Call WriteDisconnect(UserIndex)
122                 Call CloseSocket(UserIndex)
                End If

            End If

        End With

        Exit Sub

Cerrar_Usuario_Err:
124     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.Cerrar_Usuario", Erl)


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
100     If UserList(UserIndex).Counters.Saliendo And UserList(UserIndex).ConnIDValida Then

            ' Is the user still connected?
102         If UserList(UserIndex).ConnIDValida Then
104             UserList(UserIndex).Counters.Saliendo = False
106             UserList(UserIndex).Counters.Salir = 0
108             Call WriteConsoleMsg(UserIndex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
            Else

                'Simply reset
110             If UserList(UserIndex).flags.Privilegios = PlayerType.user And MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0 Then
112                 UserList(UserIndex).Counters.Salir = IntervaloCerrarConexion
                Else
114                 Call WriteConsoleMsg(UserIndex, "Gracias por jugar Argentum20.", FontTypeNames.FONTTYPE_INFO)
116                 Call WriteDisconnect(UserIndex)
                
118                 Call CloseSocket(UserIndex)

                End If
            
                'UserList(UserIndex).Counters.Salir = IIf((UserList(UserIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0, IntervaloCerrarConexion, 0)
            End If

        End If

        
        Exit Sub

CancelExit_Err:
120     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.CancelExit", Erl)

        
End Sub

Sub VolverCriminal(ByVal UserIndex As Integer)
        
    On Error GoTo VolverCriminal_Err
        

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 21/06/2006
    'Nacho: Actualiza el tag al cliente
    '**************************************************************
        
100 With UserList(UserIndex)
        
102     If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub

104     If .flags.Privilegios And (PlayerType.user Or PlayerType.Consejero) Then
   
106         If .Faccion.ArmadaReal = 1 Then
                ' WyroX: NUNCA debería pasar, pero dejo un log por si las...
                Call RegistrarError(111, "Un personaje de la Armada Real atacó un ciudadano.", "UsUaRiOs.VolverCriminal")
                'Call ExpulsarFaccionReal(UserIndex)
            End If

        End If

108     If .Faccion.FuerzasCaos = 1 Then Exit Sub

110     .Faccion.Status = 0
        
112     If MapInfo(.Pos.Map).NoPKs And Not EsGM(UserIndex) And MapInfo(.Pos.Map).Salida.Map <> 0 Then
114         Call WriteConsoleMsg(UserIndex, "En este mapa no se admiten criminales.", FontTypeNames.FONTTYPE_INFO)
116         Call WarpUserChar(UserIndex, MapInfo(.Pos.Map).Salida.Map, MapInfo(.Pos.Map).Salida.X, MapInfo(.Pos.Map).Salida.Y, True)
        Else
118         Call RefreshCharStatus(UserIndex)
        End If

    End With
        
    Exit Sub

VolverCriminal_Err:
120     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.VolverCriminal", Erl)

        
End Sub

Sub VolverCiudadano(ByVal UserIndex As Integer)
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 21/06/2006
    'Nacho: Actualiza el tag al cliente.
    '**************************************************************
        
    On Error GoTo VolverCiudadano_Err
        
100 With UserList(UserIndex)

102     If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub

104     .Faccion.Status = 1

106     If MapInfo(.Pos.Map).NoCiudadanos And Not EsGM(UserIndex) And MapInfo(.Pos.Map).Salida.Map <> 0 Then
108         Call WriteConsoleMsg(UserIndex, "En este mapa no se admiten ciudadanos.", FontTypeNames.FONTTYPE_INFO)
110         Call WarpUserChar(UserIndex, MapInfo(.Pos.Map).Salida.Map, MapInfo(.Pos.Map).Salida.X, MapInfo(.Pos.Map).Salida.Y, True)
        Else
112         Call RefreshCharStatus(UserIndex)
        End If

        Call WriteSafeModeOn(UserIndex)
        .flags.Seguro = True

    End With
        
    Exit Sub

VolverCiudadano_Err:
114     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.VolverCiudadano", Erl)

        
End Sub

Public Function getMaxInventorySlots(ByVal UserIndex As Integer) As Byte
        '***************************************************
        'Author: Unknown
        'Last Modification: 30/09/2020
        '
        '***************************************************
        
        On Error GoTo getMaxInventorySlots_Err
        

100     If UserList(UserIndex).Stats.InventLevel > 0 Then
102         getMaxInventorySlots = MAX_USERINVENTORY_SLOTS + UserList(UserIndex).Stats.InventLevel * SLOTS_PER_ROW_INVENTORY
        Else
104         getMaxInventorySlots = MAX_USERINVENTORY_SLOTS

        End If

        
        Exit Function

getMaxInventorySlots_Err:
106     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.getMaxInventorySlots", Erl)

        
End Function

Private Sub WarpMascotas(ByVal UserIndex As Integer)
        
        On Error GoTo WarpMascotas_Err
    
        

        '************************************************
        'Author: Uknown
        'Last Modified: 26/10/2010
        '13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
        '13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
        '11/05/2009: ZaMa - Chequeo si la mascota pueden spawnear para asiganrle los stats.
        '26/10/2010: ZaMa - Ahora las mascotas respawnean de forma aleatoria.
        '************************************************
        Dim i                As Integer

        Dim petType          As Integer

        Dim canWarp          As Boolean

        Dim Index            As Integer

        Dim iMinHP           As Integer
        
        Dim PetTiempoDeVida  As Integer
    
        Dim MascotaQuitada   As Boolean
        Dim ElementalQuitado As Boolean
        Dim SpawnInvalido    As Boolean

100     canWarp = MapInfo(UserList(UserIndex).Pos.Map).Seguro = 0

102     For i = 1 To MAXMASCOTAS
104         Index = UserList(UserIndex).MascotasIndex(i)
        
106         If Index > 0 Then
108             iMinHP = NpcList(Index).Stats.MinHp
110             PetTiempoDeVida = NpcList(Index).Contadores.TiempoExistencia
            
112             NpcList(Index).MaestroUser = 0
            
114             Call QuitarNPC(Index)

116             If PetTiempoDeVida > 0 Then
118                 Call QuitarMascota(UserIndex, Index)
120                 ElementalQuitado = True

122             ElseIf Not canWarp Then
124                 UserList(UserIndex).MascotasIndex(i) = 0
126                 MascotaQuitada = True
                End If
            
            Else
128             iMinHP = 0
130             PetTiempoDeVida = 0
            End If
        
132         petType = UserList(UserIndex).MascotasType(i)
        
134         If petType > 0 And canWarp And UserList(UserIndex).flags.MascotasGuardadas = 0 And PetTiempoDeVida = 0 Then
        
                Dim SpawnPos As WorldPos
        
136             SpawnPos.Map = UserList(UserIndex).Pos.Map
138             SpawnPos.X = UserList(UserIndex).Pos.X + RandomNumber(-3, 3)
140             SpawnPos.Y = UserList(UserIndex).Pos.Y + RandomNumber(-3, 3)
        
142             Index = SpawnNpc(petType, SpawnPos, False, False, False, UserIndex)
            
                'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
                ' Exception: Pets don't spawn in water if they can't swim
144             If Index > 0 Then
146                 UserList(UserIndex).MascotasIndex(i) = Index

                    ' Nos aseguramos de que conserve el hp, si estaba danado
148                 If iMinHP Then NpcList(Index).Stats.MinHp = iMinHP

150                 Call FollowAmo(Index)
            
                Else
152                 SpawnInvalido = True
                End If

            End If

154     Next i

156     If Not canWarp And MascotaQuitada Then
158         Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Estas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)

160     ElseIf SpawnInvalido Then
162         Call WriteConsoleMsg(UserIndex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)

164     ElseIf ElementalQuitado Then
166         Call WriteConsoleMsg(UserIndex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)
        End If

        
        Exit Sub

WarpMascotas_Err:
168     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.WarpMascotas", Erl)

        
End Sub

Function TieneArmaduraCazador(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo TieneArmaduraCazador_Err
    
        

100     If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
        
102         If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Subtipo = 3 Then ' Aguante hardcodear números :D
104             TieneArmaduraCazador = True
            End If
        
        End If

        
        Exit Function

TieneArmaduraCazador_Err:
106     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.TieneArmaduraCazador", Erl)

        
End Function

Public Sub SetModoConsulta(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Torres Patricio (Pato)
        'Last Modification: 05/06/10
        '
        '***************************************************

        Dim sndNick As String

100     With UserList(UserIndex)
102         sndNick = .Name
    
104         If .flags.EnConsulta Then
106             sndNick = sndNick & " [CONSULTA]"
            
            Else

108             If .GuildIndex > 0 Then
110                 sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"
                End If

            End If

112         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateTagAndStatus(UserIndex, .Faccion.Status, sndNick))

        End With

End Sub

' Autor: WyroX - 20/01/2021
' Intenta moverlo hacia un "costado" según el heading indicado.
' Si no hay un lugar válido a los lados, lo mueve a la posición válida más cercana.
Sub MoveUserToSide(ByVal UserIndex As Integer, ByVal Heading As eHeading)

        On Error GoTo Handler

100     With UserList(UserIndex)

            ' Elegimos un lado al azar
            Dim R As Integer
102         R = RandomNumber(0, 1) * 2 - 1 ' -1 o 1

            ' Roto el heading original hacia ese lado
104         Heading = RotateHeading(Heading, R)

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
            Dim NuevaPos As WorldPos
116         Call ClosestLegalPos(.Pos, NuevaPos, .flags.Navegando, .flags.Navegando = 0)
118         Call WarpUserChar(UserIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

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
134             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.CharIndex, 0))
            End If
        
            '<<<< Invisible >>>>
136         If (.flags.invisible = 1 Or .flags.Oculto = 1) And .flags.AdminInvisible = 0 Then
138             .flags.Oculto = 0
140             .flags.invisible = 0
142             .Counters.TiempoOculto = 0
144             .Counters.Invisibilidad = 0
146             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            End If
        
            '<<<< Mimetismo >>>>
148         If .flags.Mimetizado > 0 Then
        
150             If .flags.Navegando Then
            
152                 If .flags.Muerto = 0 Then
154                     .Char.Body = ObjData(UserList(UserIndex).Invent.BarcoObjIndex).Ropaje
                    Else
156                     .Char.Body = iFragataFantasmal
                    End If

158                 .Char.ShieldAnim = NingunEscudo
160                 .Char.WeaponAnim = NingunArma
162                 .Char.CascoAnim = NingunCasco
                
                Else
            
164                 .Char.Body = .CharMimetizado.Body
166                 .Char.Head = .CharMimetizado.Head
168                 .Char.CascoAnim = .CharMimetizado.CascoAnim
170                 .Char.ShieldAnim = .CharMimetizado.ShieldAnim
172                 .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                
                End If
            
174             .Counters.Mimetismo = 0
176             .flags.Mimetizado = e_EstadoMimetismo.Desactivado
            End If
        
            '<<<< Estados obsoletos >>>>
178         .flags.Ahogandose = 0
180         .flags.Incinerado = 0
        
        End With
    
        Exit Sub
    
Handler:
182     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.LimpiarEstadosAlterados", Erl)


End Sub

Public Sub DevolverPosAnterior(ByVal UserIndex As Integer)

100     With UserList(UserIndex).flags
102         Call WarpToLegalPos(UserIndex, .LastPos.Map, .LastPos.X, .LastPos.Y, True)
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
114         If (.flags.Navegando + .flags.Nadando > 0) And (.Invent.BarcoObjIndex > 0) Then
116             modificadorItem = ObjData(.Invent.BarcoObjIndex).velocidad
            End If
        
118         If (.flags.Montado = 1) And (.Invent.MonturaObjIndex > 0) Then
120             modificadorItem = ObjData(.Invent.MonturaObjIndex).velocidad
            End If
        
            ' Algun hechizo le afecto la velocidad
122         If .flags.VelocidadHechizada > 0 Then
124             modificadorHechizo = .flags.VelocidadHechizada
            End If
        
126         velocidad = VelocidadNormal * modificadorItem * modificadorHechizo
        
UpdateSpeed:
128         .Char.speeding = velocidad
        
130         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSpeedingACT(.Char.CharIndex, .Char.speeding))
132         Call WriteVelocidadToggle(UserIndex)
     
        End With

        Exit Function
    
ActualizarVelocidadDeUsuario_Err:
134     Call TraceError(Err.Number, Err.Description, "UsUaRiOs.CalcularVelocidad_Err", Erl)

End Function

