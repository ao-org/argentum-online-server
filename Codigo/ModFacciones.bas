Attribute VB_Name = "ModFacciones"

Option Explicit

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
        On Error GoTo EnlistarArmadaReal_Err

        Dim charIndexStr As String

        With UserList(UserIndex)
            charIndexStr = str(NpcList(.flags.TargetNPC).Char.CharIndex)

            If .Faccion.ArmadaReal = 1 Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a las tropas reales!!! Ve a combatir criminales", charIndexStr, vbWhite)
                Exit Sub

            End If

            If .Faccion.FuerzasCaos = 1 Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Maldito insolente!!! vete de aqui seguidor de las sombras", charIndexStr, vbWhite)
                Exit Sub

            End If

            If Status(UserIndex) = 2 Then
                Call WriteChatOverHead(UserIndex, "¡¡¡No se permiten criminales en el ejercito imperial!!!", charIndexStr, vbWhite)
                Exit Sub

            End If

            If Status(UserIndex) = 0 Then
                Call WriteChatOverHead(UserIndex, "¡¡¡No se permiten neutrales en el ejercito imperial, antes debes solicitar el perdón con un sacerdote!!!", charIndexStr, vbWhite)
                Exit Sub

            End If

            If .clase = eClass.Thief Then
                Call WriteChatOverHead(UserIndex, "¡Los ladrones no son dignos para pertenecer al ejercito imperial!!!", charIndexStr, vbWhite)
                Exit Sub

            End If

            'Si fue miembro de la Legión del Caos no se puede enlistar
            If .Faccion.ciudadanosMatados > 0 Then
                Call WriteChatOverHead(UserIndex, "¡Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!", charIndexStr, vbWhite)
                Exit Sub

            End If

            If .Faccion.Reenlistadas > 0 Then
                Call WriteChatOverHead(UserIndex, "¡Has traicionado al ejercito imperial una vez, no aceptamos traidores!", charIndexStr, vbWhite)
                Exit Sub

            End If

            Dim primerRango As tRangoFaccion
                primerRango = RangosFaccion(1)

            If .Faccion.CriminalesMatados < primerRango.AsesinatosRequeridos Then
                Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos " & primerRango.AsesinatosRequeridos & " criminales, solo has matado " & .Faccion.CriminalesMatados, charIndexStr, vbWhite)
                Exit Sub

            End If

            If .Stats.ELV < primerRango.NivelRequerido Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel " & primerRango.NivelRequerido & "!!!", charIndexStr, vbWhite)
                Exit Sub

            End If


            ' Cumple con los requisitos para enlistarse
            .Faccion.ArmadaReal = 1
            .Faccion.RecompensasReal = primerRango.rank ' Asigna primer rango
            .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
            .Faccion.Status = 3

            If .Faccion.RecibioArmaduraReal = 0 Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al Ejercito Imperial!!!, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", charIndexStr, vbWhite)

                .Faccion.FechaIngreso = Date
                .Faccion.NivelIngreso = .Stats.ELV
                .Faccion.MatadosIngreso = .Faccion.ciudadanosMatados

            End If

            Call WriteConsoleMsg(UserIndex, "¡Ahora perteneces al Ejercito Real!", FontTypeNames.FONTTYPE_INFOIAO)
            Call DarRecompensas(UserIndex)
            Call RefreshCharStatus(UserIndex)

        End With

        'Call LogEjercitoReal(.Name & " ingresó el " & Date & " cuando era nivel " & .Stats.ELV)

        Exit Sub

EnlistarArmadaReal_Err:
        Call RegistrarError(Err.Number, Err.Description, "ModFacciones.EnlistarArmadaReal", Erl)
        Resume Next

End Sub

' Subir de Rango y asignar recompensas.
Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
        On Error GoTo RecompensaArmadaReal_Err

        Dim Crimis As Long, npcCharIndex As String
        Dim proxRango As tRangoFaccion

        With UserList(UserIndex)
            Crimis = .Faccion.CriminalesMatados
            npcCharIndex = str(NpcList(.flags.TargetNPC).Char.CharIndex)

            If .Faccion.RecompensasReal >= MaxRangoFaccion Then
                Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores Soldados. Mataste " & Crimis & ", sigue asi. Ya no tengo más recompensa para darte que mi agradescimiento. ¡Felicidades!", npcCharIndex, vbWhite)

                Exit Sub
            End If

            proxRango = ProximoRango(UserIndex)

            If Crimis < proxRango.AsesinatosRequeridos Then
                Call WriteChatOverHead(UserIndex, "Mata " & proxRango.AsesinatosRequeridos - Crimis & " Criminales más para recibir la próxima Recompensa", npcCharIndex, vbWhite)
                Exit Sub

            End If

            If proxRango.NivelRequerido > .Stats.ELV Then
              Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & (proxRango.NivelRequerido - .Stats.ELV) & " Niveles para poder recibir la próxima Recompensa", npcCharIndex, vbWhite)

            Else ' El usuario cumple con los requerimientos de nivel, se le asigna la recomenza.
              .Faccion.RecompensasReal = proxRango.rank

              Call WriteChatOverHead(UserIndex, "¡¡¡Aqui tienes tu recompensa " + proxRango.Titulo + "!!!", npcCharIndex, vbWhite)
              Call DarRecompensas(UserIndex)
              '.Stats.Exp = .Stats.Exp + ExpX100
            End If
        End With

        Exit Sub

RecompensaArmadaReal_Err:
        Call RegistrarError(Err.Number, Err.Description, "ModFacciones.RecompensaArmadaReal", Erl)
        Resume Next

End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)
        On Error GoTo ExpulsarFaccionReal_Err

        UserList(UserIndex).Faccion.ArmadaReal = 0
        UserList(UserIndex).Faccion.Status = 1
        Call RefreshCharStatus(UserIndex)

        Call PerderItemsFaccionarios(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has sido expulsado del Ejercito Real.", FontTypeNames.FONTTYPE_INFOIAO)


        Exit Sub

ExpulsarFaccionReal_Err:
        Call RegistrarError(Err.Number, Err.Description, "ModFacciones.ExpulsarFaccionReal", Erl)
        Resume Next

End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)

        On Error GoTo ExpulsarFaccionCaos_Err

        UserList(UserIndex).Faccion.FuerzasCaos = 0
        UserList(UserIndex).Faccion.Status = 0
        Call RefreshCharStatus(UserIndex)

        Call PerderItemsFaccionarios(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has sido expulsado de la Legión Oscura.", FontTypeNames.FONTTYPE_INFOIAO)

        Exit Sub

ExpulsarFaccionCaos_Err:
        Call RegistrarError(Err.Number, Err.Description, "ModFacciones.ExpulsarFaccionCaos", Erl)
        Resume Next

End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String
        On Error GoTo TituloReal_Err

        Dim rank As Byte
            rank = UserList(UserIndex).Faccion.RecompensasReal

        If rank > 0 Then
            'Los indices impares son los Rangos de Armada
            TituloReal = RangosFaccion(2 * rank - 1).Titulo
        End If

        Exit Function

TituloReal_Err:
        Call RegistrarError(Err.Number, Err.Description, "ModFacciones.TituloReal", Erl)
        Resume Next

End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
        On Error GoTo EnlistarCaos_Err

        Dim charIndexStr As String

        With UserList(UserIndex)
            charIndexStr = str(NpcList(.flags.TargetNPC).Char.CharIndex)

            If .Faccion.FuerzasCaos = 1 Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a la legión oscura!!!", charIndexStr, vbWhite)
                Exit Sub

            End If

            If .Faccion.ArmadaReal = 1 Then
                Call WriteChatOverHead(UserIndex, "Las sombras reinarán en Argentum. ¡¡¡Fuera de aqui insecto Real!!!", charIndexStr, vbWhite)
                Exit Sub

            End If

            'Si fue miembro de la Armada Real no se puede enlistar
            If .Faccion.RecompensasReal > 0 Then
                Call WriteChatOverHead(UserIndex, "No permitiré que ningún insecto real ingrese a mis tropas.", charIndexStr, vbWhite)
                Exit Sub

            End If

            If Status(UserIndex) = 1 Then
                Call WriteChatOverHead(UserIndex, "¡¡Ja ja ja!! Tu no eres bienvenido aquí asqueroso Ciudadano", charIndexStr, vbWhite)
                Exit Sub

            End If

            If .clase = eClass.Thief Then
                Call WriteChatOverHead(UserIndex, "¡La legión oscura no tiene lugar para escorias como tú! Los ladrones no son dignos de llevar nuestras armaduras.", charIndexStr, vbWhite)
                Exit Sub

            End If

            If UserList(UserIndex).Faccion.Reenlistadas > 0 Then
                Call WriteChatOverHead(UserIndex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. ¡Vete de aquí!", charIndexStr, vbWhite)
                Exit Sub

            End If


            Dim primerRango As tRangoFaccion
                primerRango = RangosFaccion(2) ' 2 es el primer rango del caos

            If .Faccion.ciudadanosMatados < primerRango.AsesinatosRequeridos Then
                Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos " & primerRango.AsesinatosRequeridos & " ciudadanos, solo has matado " & .Faccion.ciudadanosMatados, charIndexStr, vbWhite)
                Exit Sub

            End If

            If .Stats.ELV < primerRango.NivelRequerido Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel " & primerRango.NivelRequerido & "!!!", charIndexStr, vbWhite)
                Exit Sub

            End If


            ' Cumple con los requisitos para enlistarse
            .Faccion.FuerzasCaos = 1
            .Faccion.RecompensasCaos = primerRango.rank ' Asigna primer rango
            .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
            .Faccion.Status = 2

            If .Faccion.RecibioArmaduraCaos = 0 Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al lado oscuro!!! Aqui tienes tus armaduras. Derrama sangre Ciudadana y Real y serás recompensado, lo prometo.", charIndexStr, vbWhite)

                .Faccion.FechaIngreso = Date
                .Faccion.NivelIngreso = .Stats.ELV
            End If

            Call WriteConsoleMsg(UserIndex, "¡Ahora perteneces al caos!", FontTypeNames.FONTTYPE_INFOIAO)
            Call DarRecompensas(UserIndex)
            Call RefreshCharStatus(UserIndex)

        End With

        'Call LogEjercitoCaos(UserList(UserIndex).Name & " ingresó el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

        Exit Sub

EnlistarCaos_Err:
        Call RegistrarError(Err.Number, Err.Description, "ModFacciones.EnlistarCaos", Erl)
        Resume Next

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
        On Error GoTo RecompensaCaos_Err


        Dim ciudadanosMatados As Long, npcCharIndex As String
        Dim proxRango As tRangoFaccion

        With UserList(UserIndex)
            ciudadanosMatados = .Faccion.ciudadanosMatados
            npcCharIndex = str(NpcList(.flags.TargetNPC).Char.CharIndex)

            If .Faccion.RecompensasCaos >= MaxRangoFaccion Then
                Call WriteChatOverHead(UserIndex, "¡Felicitaciones, eres de mis mejores guerreros, estas en lo más alto!", npcCharIndex, vbWhite)
                Exit Sub

            End If

            proxRango = ProximoRango(UserIndex)

            If ciudadanosMatados < proxRango.AsesinatosRequeridos Then
                Call WriteChatOverHead(UserIndex, "Mata " & proxRango.AsesinatosRequeridos - ciudadanosMatados & " Ciudadanos más para recibir la próxima Recompensa", npcCharIndex, vbWhite)
                Exit Sub

            End If

            If proxRango.NivelRequerido > .Stats.ELV Then
              Call WriteChatOverHead(UserIndex, "Mataste Suficientes Ciudadanos, pero te faltan " & (proxRango.NivelRequerido - .Stats.ELV) & " Niveles para poder recibir la próxima Recompensa", npcCharIndex, vbWhite)

            Else ' El usuario cumple con los requerimientos de nivel, se le asigna la recomenza.
              .Faccion.RecompensasCaos = proxRango.rank

              Call WriteChatOverHead(UserIndex, "¡¡¡Bien hecho " + proxRango.Titulo + ", aquí tienes tu recompensa!!!", npcCharIndex, vbWhite)
              Call DarRecompensas(UserIndex)
              '.Stats.Exp = .Stats.Exp + ExpX100
            End If
        End With

        Exit Sub

RecompensaCaos_Err:
        Call RegistrarError(Err.Number, Err.Description, "ModFacciones.RecompensaCaos", Erl)
        Resume Next

End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
        On Error GoTo TituloCaos_Err

        Dim rank As Byte
            rank = UserList(UserIndex).Faccion.RecompensasCaos

        If rank > 0 Then
            'Los indices pares son los Rangos del Caos
            TituloCaos = RangosFaccion(2 * rank).Titulo
        End If

        Exit Function

TituloCaos_Err:
        Call RegistrarError(Err.Number, Err.Description, "ModFacciones.TituloCaos", Erl)
        Resume Next

End Function


' Devuelve el proximo rango para el usuario de la faccion que pertenece.
Private Function ProximoRango(ByVal UserIndex As Integer) As tRangoFaccion
        On Error GoTo ProximoRango_Err

        With UserList(UserIndex)
            If .Faccion.ArmadaReal = 1 And .Faccion.RecompensasReal < MaxRangoFaccion Then
                ProximoRango = RangosFaccion(2 * .Faccion.RecompensasReal + 1)
            ElseIf .Faccion.FuerzasCaos = 1 And .Faccion.RecompensasCaos < MaxRangoFaccion Then
                ProximoRango = RangosFaccion(2 * .Faccion.RecompensasCaos + 2)
            Else ' No pertenece a ninguna faccion.
                ' No devuelve nada
            End If
        End With

        Exit Function

ProximoRango_Err:
        Call RegistrarError(Err.Number, Err.Description, "ModFacciones.TituloCaos", Erl)
        Resume Next

End Function


' Rutina para dar las recompensas de faccion al usuario.
' Si el usuario sube mas de un rango por vez, esta rutina le dara TODOS los objetos
' que debería tener hasta alcanzar su rango.
Private Sub DarRecompensas(ByVal UserIndex As Integer)
        On Error GoTo DarRecompensas_Err

        Dim recompensa As tRecompensaFaccion
        Dim rank As Byte
        Dim ultimaRecompensa As Byte
        Dim objetoRecompensa As obj
        Dim i As Integer

        With UserList(UserIndex)
            If .Faccion.ArmadaReal = 1 Then
                rank = .Faccion.RecompensasReal
                ultimaRecompensa = .Faccion.RecibioArmaduraReal
            ElseIf .Faccion.FuerzasCaos = 1 Then
                rank = .Faccion.RecompensasCaos
                ultimaRecompensa = .Faccion.RecibioArmaduraCaos
            Else ' No pertenece a ninguna faccion.
                Exit Sub
            End If

            If ultimaRecompensa >= rank Then
                Exit Sub
            End If

            ' Esto puede parecer ineficiente, pero DarRecompensas sucede pocas veces en el juego.
            ' Por ahora, iterar por todas las recompensas es mas facil que mantener una estructura mas
            ' complicada (como diccionarios). El total de recompensas se puede aproximar como: `C * R * F * nR`
            ' C = 12 (clases distintas); R = 6 (max rango por faccion); F = 2 (facciones distintas);
            ' nR = 1 (numero de recompensas por rango)
            For i = 1 To UBound(RecompensasFaccion)
                recompensa = RecompensasFaccion(i)

                ' Como puede subir varios rangos todos juntos, nos aseguramos que
                ' entregamos TODAS las recompensas hasta el rango actual desde la ultima recompensa.
                If recompensa.rank <= rank And recompensa.rank > ultimaRecompensa Then
                    ' Por alguna razon, PuedeUsarObjeto devuelve 0 cuando el usuario SI puede usarlo.
                    If PuedeUsarObjeto(UserIndex, recompensa.ObjIndex, False) = 0 Then
                        objetoRecompensa.Amount = 1
                        objetoRecompensa.ObjIndex = recompensa.ObjIndex

                        If Not MeterItemEnInventario(UserIndex, objetoRecompensa) Then
                            Call TirarItemAlPiso(.Pos, objetoRecompensa)

                        End If
                    End If
                End If

            Next i

            ' Guardamos que el usuario recibio las recompensas de su rank.
            If .Faccion.ArmadaReal = 1 Then
              .Faccion.RecibioArmaduraReal = rank
            Else
              .Faccion.RecibioArmaduraCaos = rank
            End If

            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave(48, NO_3D_SOUND, NO_3D_SOUND))
        End With

        Exit Sub

DarRecompensas_Err:
        Call RegistrarError(Err.Number, Err.Description, "ModFacciones.DarRecompensas", Erl)
        Resume Next

End Sub


Private Sub PerderItemsFaccionarios(ByVal UserIndex As Integer)
        On Error GoTo PerderItemsFaccionarios_Err

        Dim i         As Byte
        Dim ItemIndex As Integer

        With UserList(UserIndex)
            For i = 1 To .CurrentInventorySlots
                ItemIndex = .Invent.Object(i).ObjIndex

                If ItemIndex > 0 Then

                    If ObjData(ItemIndex).Real = 1 Or ObjData(ItemIndex).Caos = 1 Then

                        Call QuitarUserInvItem(UserIndex, i, MAX_INVENTORY_OBJS)
                        Call UpdateUserInv(False, UserIndex, i)

                    End If
                End If
            Next i
        End With

        Exit Sub

PerderItemsFaccionarios_Err:
        Call RegistrarError(Err.Number, Err.Description, "ModFacciones.PerderItemsFaccionarios", Erl)
        Resume Next

End Sub
