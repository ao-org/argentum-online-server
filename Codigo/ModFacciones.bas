Attribute VB_Name = "ModFacciones"

Option Explicit

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
            On Error GoTo EnlistarArmadaReal_Err

            Dim charIndexStr As String

100         With UserList(UserIndex)
102             charIndexStr = str(NpcList(.flags.TargetNPC).Char.CharIndex)

104             If .Faccion.ArmadaReal = 1 Then
106                 Call WriteChatOverHead(UserIndex, "Ya perteneces a mi ejército jóven soldado. Ve a combatir el caos en mis tierras para subir de rango en el Ejército Real.", charIndexStr, vbWhite)
                    Exit Sub

                End If

108             If .Faccion.FuerzasCaos = 1 Then
110                 Call WriteChatOverHead(UserIndex, "¡Has llegado al lugar equivocado maldita escoria! Vete de aquí antes de ser encarcelado e interrogado.", charIndexStr, vbWhite)
                    Exit Sub

                End If

112             If Status(UserIndex) = 2 Then
114                 Call WriteChatOverHead(UserIndex, "No se permiten criminales en el Ejército Real.", charIndexStr, vbWhite)
                    Exit Sub
                End If
                

116             If Status(UserIndex) = 0 Then
118                 Call WriteChatOverHead(UserIndex, "No se permiten criminales en el Ejército Real. Si no has derramado sangre inocente ve a inclinarte y pedir perdón ante un sacerdote.", charIndexStr, vbWhite)
                    Exit Sub

                End If

120             If .clase = eClass.Thief Then
122                 Call WriteChatOverHead(UserIndex, "No hay lugar para escoria en el Ejército Real.", charIndexStr, vbWhite)
                    Exit Sub

                End If

                'Si fue miembro de la Legión del Caos no se puede enlistar
124             If .Faccion.ciudadanosMatados > 0 Then
126                 Call WriteChatOverHead(UserIndex, "Has derramado sangre inocente. Solo aceptamos almas nobles en el Ejército Real.", charIndexStr, vbWhite)
                    Exit Sub

                End If
                
128             If Not HayLugarEnInventario(UserIndex) Then
130                 Call WriteChatOverHead(UserIndex, "¡No tienes lugar suficiente en el inventario.", charIndexStr, vbWhite)
                    Exit Sub
                End If
                

132             If .Faccion.Reenlistadas > 0 Then
134                 Call WriteChatOverHead(UserIndex, "Ya has desertado el Ejército Real. No serás aceptado otra vez.", charIndexStr, vbWhite)
                    Exit Sub

                End If

                Dim primerRango As tRangoFaccion
136                 primerRango = RangosFaccion(1)

138             If .Faccion.CriminalesMatados < primerRango.AsesinatosRequeridos Then
140                 Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos " & primerRango.AsesinatosRequeridos & " criminales, solo has matado " & .Faccion.CriminalesMatados, charIndexStr, vbWhite)
                    Exit Sub

                End If

142             If .Stats.ELV < primerRango.NivelRequerido Then
144                 Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel " & primerRango.NivelRequerido & "!!!", charIndexStr, vbWhite)
                    Exit Sub

                End If
                
                'HarThaoS: Lo pongo al final para que lo expulse del clan solamente si cumple todos los requisitos.
                If .GuildIndex > 0 Then
                    If PersonajeEsLeader(.Name) Then
                        Call WriteChatOverHead(UserIndex, "No puedes integrar nuestras fuerzas si perteneces a un clan neutral, deberás abandonarlo si tu deseo es integrar el Ejército Real.", charIndexStr, vbWhite)
                        Exit Sub
                    Else
                        Call m_EcharMiembroDeClan(UserIndex, .Name)
                    End If
                End If

                ' Cumple con los requisitos para enlistarse
146             .Faccion.ArmadaReal = 1
148             .Faccion.RecompensasReal = primerRango.rank ' Asigna primer rango
150             .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
152             .Faccion.Status = 3

154             If .Faccion.RecibioArmaduraReal = 0 Then
156                 Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al Ejercito Imperial!!!, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", charIndexStr, vbWhite)

158                 .Faccion.NivelIngreso = .Stats.ELV
160                 .Faccion.MatadosIngreso = .Faccion.ciudadanosMatados

                End If

162             Call WriteConsoleMsg(UserIndex, "¡Ahora perteneces al Ejercito Real!", FontTypeNames.FONTTYPE_INFOIAO)
164             Call DarRecompensas(UserIndex)
166             Call RefreshCharStatus(UserIndex)

            End With

            'Call LogEjercitoReal(.Name & " ingresó el " & Date & " cuando era nivel " & .Stats.ELV)

            Exit Sub

EnlistarArmadaReal_Err:
168         Call RegistrarError(Err.Number, Err.Description, "ModFacciones.EnlistarArmadaReal", Erl)
170         Resume Next

End Sub

' Subir de Rango y asignar recompensas.
Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
            On Error GoTo RecompensaArmadaReal_Err

            Dim Crimis As Long, npcCharIndex As String
            Dim proxRango As tRangoFaccion

100         With UserList(UserIndex)
102             Crimis = .Faccion.CriminalesMatados
104             npcCharIndex = str(NpcList(.flags.TargetNPC).Char.CharIndex)

106             If .Faccion.RecompensasReal >= MaxRangoFaccion Then
108                 Call WriteChatOverHead(UserIndex, "Has alcanzado el máximo rango dentro de mis soldados. Demuestra tu bondad y liderazgo en el campo de batalla para algún día pertenecer al Consejo de Banderbill.", npcCharIndex, vbWhite)
                    Exit Sub
                End If
                
110             If Not HayLugarEnInventario(UserIndex) Then
112                 Call WriteChatOverHead(UserIndex, "¡No tienes lugar suficiente en el inventario.", npcCharIndex, vbWhite)
                    Exit Sub
                End If

114             proxRango = ProximoRango(UserIndex)

116             If Crimis < proxRango.AsesinatosRequeridos Then
118                 Call WriteChatOverHead(UserIndex, "Mata " & proxRango.AsesinatosRequeridos - Crimis & " Criminales más para recibir la próxima Recompensa", npcCharIndex, vbWhite)
                    Exit Sub

                End If

120             If proxRango.NivelRequerido > .Stats.ELV Then
122               Call WriteChatOverHead(UserIndex, "Has matado suficientes criminales pero, te faltan & (proxRango.NivelRequerido - .Stats.ELV) & para poder recibir la próxima recompensa.", npcCharIndex, vbWhite)

                Else ' El usuario cumple con los requerimientos de nivel, se le asigna la recomenza.
124               .Faccion.RecompensasReal = proxRango.rank

126               Call WriteChatOverHead(UserIndex, "¡¡¡Aqui tienes tu recompensa " + proxRango.Titulo + "!!!", npcCharIndex, vbWhite)
128               Call DarRecompensas(UserIndex)
                  '.Stats.Exp = .Stats.Exp + ExpX100
                End If
            End With

            Exit Sub

RecompensaArmadaReal_Err:
130         Call RegistrarError(Err.Number, Err.Description, "ModFacciones.RecompensaArmadaReal", Erl)
132         Resume Next

End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)
            On Error GoTo ExpulsarFaccionReal_Err

100         UserList(UserIndex).Faccion.ArmadaReal = 0
102         UserList(UserIndex).Faccion.Status = 1
104         Call RefreshCharStatus(UserIndex)

106         Call PerderItemsFaccionarios(UserIndex)
108         Call WriteConsoleMsg(UserIndex, "Has sido expulsado del Ejercito Real.", FontTypeNames.FONTTYPE_INFOIAO)
            

            Exit Sub

ExpulsarFaccionReal_Err:
110         Call RegistrarError(Err.Number, Err.Description, "ModFacciones.ExpulsarFaccionReal", Erl)
112         Resume Next

End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)

            On Error GoTo ExpulsarFaccionCaos_Err

100         UserList(UserIndex).Faccion.FuerzasCaos = 0
102         UserList(UserIndex).Faccion.Status = 0
104         Call RefreshCharStatus(UserIndex)

106         Call PerderItemsFaccionarios(UserIndex)
108         Call WriteConsoleMsg(UserIndex, "Has sido expulsado de la Legión Oscura.", FontTypeNames.FONTTYPE_INFOIAO)

            Exit Sub

ExpulsarFaccionCaos_Err:
110         Call RegistrarError(Err.Number, Err.Description, "ModFacciones.ExpulsarFaccionCaos", Erl)
112         Resume Next

End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String
            On Error GoTo TituloReal_Err

            Dim rank As Byte
100             rank = UserList(UserIndex).Faccion.RecompensasReal

102         If rank > 0 Then
                'Los indices impares son los Rangos de Armada
104             TituloReal = RangosFaccion(2 * rank - 1).Titulo
            End If

            Exit Function

TituloReal_Err:
106         Call RegistrarError(Err.Number, Err.Description, "ModFacciones.TituloReal", Erl)
108         Resume Next

End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
            On Error GoTo EnlistarCaos_Err

            Dim charIndexStr As String

100         With UserList(UserIndex)
102             charIndexStr = str(NpcList(.flags.TargetNPC).Char.CharIndex)

104             If .Faccion.FuerzasCaos = 1 Then
106                 Call WriteChatOverHead(UserIndex, "Ya perteneces a la Legión Oscura.", charIndexStr, vbWhite)
                    Exit Sub

                End If

108             If .Faccion.ArmadaReal = 1 Then
110                 Call WriteChatOverHead(UserIndex, "El caos y el poder de las sombras reinará en este mundo y tu sufrirás maldito súbdito de Tancredo.", charIndexStr, vbWhite)
                    Exit Sub

                End If

                'Si fue miembro de la Armada Real no se puede enlistar
112             If .Faccion.RecompensasReal > 0 Then
114                 Call WriteChatOverHead(UserIndex, "No permitiré que ningún insecto real ingrese a mis tropas.", charIndexStr, vbWhite)
                    Exit Sub

                End If

116             If Status(UserIndex) = 1 Then
118                 Call WriteChatOverHead(UserIndex, "¡¡Ja ja ja!! Tu no eres bienvenido aquí asqueroso Ciudadano", charIndexStr, vbWhite)
                    Exit Sub

                End If
                
120             If Not HayLugarEnInventario(UserIndex) Then
122                 Call WriteChatOverHead(UserIndex, "¡No tienes lugar suficiente en el inventario.", charIndexStr, vbWhite)
                    Exit Sub
                End If

124             If .clase = eClass.Thief Then
126                 Call WriteChatOverHead(UserIndex, "¡La legión oscura no tiene lugar para escorias como tú! Los ladrones no son dignos de llevar nuestras armaduras.", charIndexStr, vbWhite)
                    Exit Sub

                End If

128             If UserList(UserIndex).Faccion.Reenlistadas > 0 Then
130                 Call WriteChatOverHead(UserIndex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. ¡Vete de aquí!", charIndexStr, vbWhite)
                    Exit Sub

                End If


                Dim primerRango As tRangoFaccion
132                 primerRango = RangosFaccion(2) ' 2 es el primer rango del caos

134             If .Faccion.ciudadanosMatados < primerRango.AsesinatosRequeridos Then
136                 Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos " & primerRango.AsesinatosRequeridos & " ciudadanos, solo has matado " & .Faccion.ciudadanosMatados, charIndexStr, vbWhite)
                    Exit Sub

                End If

138             If .Stats.ELV < primerRango.NivelRequerido Then
140                 Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel " & primerRango.NivelRequerido & "!!!", charIndexStr, vbWhite)
                    Exit Sub
                End If
                
                
                'HarThaoS: Lo pongo al final para que lo expulse del clan solamente si cumple todos los requisitos.
                If .GuildIndex > 0 Then
                    If PersonajeEsLeader(.Name) Then
                        Call WriteChatOverHead(UserIndex, "Si quieres unirte a nuestras tropas deberás abandonar tu clan neutral.", charIndexStr, vbWhite)
                        Exit Sub
                    Else
                        Call m_EcharMiembroDeClan(UserIndex, .Name)
                    End If
                End If

                ' Cumple con los requisitos para enlistarse
142             .Faccion.FuerzasCaos = 1
144             .Faccion.RecompensasCaos = primerRango.rank ' Asigna primer rango
146             .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
148             .Faccion.Status = 2

150             If .Faccion.RecibioArmaduraCaos = 0 Then
152                 Call WriteChatOverHead(UserIndex, "Aquí tienes tu armadura legionario, ve a derramar sangre de los súbditos de Tancredo. Esta guerra será larga y cruel.", charIndexStr, vbWhite)
154                 .Faccion.NivelIngreso = .Stats.ELV
                End If

156             Call WriteConsoleMsg(UserIndex, "¡Ahora perteneces a la Legión Oscura.!", FontTypeNames.FONTTYPE_INFOIAO)
158             Call DarRecompensas(UserIndex)
160             Call RefreshCharStatus(UserIndex)

            End With

            'Call LogEjercitoCaos(UserList(UserIndex).Name & " ingresó el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

            Exit Sub

EnlistarCaos_Err:
162         Call RegistrarError(Err.Number, Err.Description, "ModFacciones.EnlistarCaos", Erl)
164         Resume Next

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
            On Error GoTo RecompensaCaos_Err


            Dim ciudadanosMatados As Long, npcCharIndex As String
            Dim proxRango As tRangoFaccion

100         With UserList(UserIndex)
102             ciudadanosMatados = .Faccion.ciudadanosMatados
104             npcCharIndex = str(NpcList(.flags.TargetNPC).Char.CharIndex)

106             If .Faccion.RecompensasCaos >= MaxRangoFaccion Then
108                 Call WriteChatOverHead(UserIndex, "¡Has alcanzado uno de los mejores lugares en mis filas. Mantén firme tu liderazgo y crueldad para algún día formar parte del Concilio de las Sombras.", npcCharIndex, vbWhite)
                    Exit Sub
                End If
                
110             If Not HayLugarEnInventario(UserIndex) Then
112                 Call WriteChatOverHead(UserIndex, "¡No tienes lugar suficiente en el inventario.", npcCharIndex, vbWhite)
                    Exit Sub
                End If

114             proxRango = ProximoRango(UserIndex)

116             If ciudadanosMatados < proxRango.AsesinatosRequeridos Then
118                 Call WriteChatOverHead(UserIndex, "Mata " & proxRango.AsesinatosRequeridos - ciudadanosMatados & " Ciudadanos más para recibir la próxima Recompensa", npcCharIndex, vbWhite)
                    Exit Sub
                End If

120             If proxRango.NivelRequerido > .Stats.ELV Then
122               Call WriteChatOverHead(UserIndex, "Has acabado con la vida de suficientes enemigos pero aún te faltan " & (proxRango.NivelRequerido - .Stats.ELV) & " niveles para alcanzar el siguiente rango.", npcCharIndex, vbWhite)

                Else ' El usuario cumple con los requerimientos de nivel, se le asigna la recomenza.
124               .Faccion.RecompensasCaos = proxRango.rank

126               Call WriteChatOverHead(UserIndex, "¡¡¡Bien hecho " + proxRango.Titulo + ", aquí tienes tu recompensa, sigue pregonando el caos a lo largo de estas tierras.!!! ", npcCharIndex, vbWhite)
128               Call DarRecompensas(UserIndex)
                  '.Stats.Exp = .Stats.Exp + ExpX100
                End If
            End With

            Exit Sub

RecompensaCaos_Err:
130         Call RegistrarError(Err.Number, Err.Description, "ModFacciones.RecompensaCaos", Erl)
132         Resume Next

End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
            On Error GoTo TituloCaos_Err

            Dim rank As Byte
100             rank = UserList(UserIndex).Faccion.RecompensasCaos

102         If rank > 0 Then
                'Los indices pares son los Rangos del Caos
104             TituloCaos = RangosFaccion(2 * rank).Titulo
            End If

            Exit Function

TituloCaos_Err:
106         Call RegistrarError(Err.Number, Err.Description, "ModFacciones.TituloCaos", Erl)
108         Resume Next

End Function


' Devuelve el proximo rango para el usuario de la faccion que pertenece.
Private Function ProximoRango(ByVal UserIndex As Integer) As tRangoFaccion
            On Error GoTo ProximoRango_Err

100         With UserList(UserIndex)
102             If .Faccion.ArmadaReal = 1 And .Faccion.RecompensasReal < MaxRangoFaccion Then
104                 ProximoRango = RangosFaccion(2 * .Faccion.RecompensasReal + 1)
106             ElseIf .Faccion.FuerzasCaos = 1 And .Faccion.RecompensasCaos < MaxRangoFaccion Then
108                 ProximoRango = RangosFaccion(2 * .Faccion.RecompensasCaos + 2)
                Else ' No pertenece a ninguna faccion.
                    ' No devuelve nada
                End If
            End With

            Exit Function

ProximoRango_Err:
110         Call RegistrarError(Err.Number, Err.Description, "ModFacciones.TituloCaos", Erl)
112         Resume Next

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

100         With UserList(UserIndex)
                ' Si es semidios o consejero, no le damos nada
                If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Exit Sub
                End If

102             If .Faccion.ArmadaReal = 1 Then
104                 rank = .Faccion.RecompensasReal
106                 ultimaRecompensa = .Faccion.RecibioArmaduraReal
108             ElseIf .Faccion.FuerzasCaos = 1 Then
110                 rank = .Faccion.RecompensasCaos
112                 ultimaRecompensa = .Faccion.RecibioArmaduraCaos
                Else ' No pertenece a ninguna faccion.
                    Exit Sub
                End If

114             If ultimaRecompensa >= rank Then
                    Exit Sub
                End If

                ' Esto puede parecer ineficiente, pero DarRecompensas sucede pocas veces en el juego.
                ' Por ahora, iterar por todas las recompensas es mas facil que mantener una estructura mas
                ' complicada (como diccionarios). El total de recompensas se puede aproximar como: `C * R * F * nR`
                ' C = 12 (clases distintas); R = 6 (max rango por faccion); F = 2 (facciones distintas);
                ' nR = 1 (numero de recompensas por rango)
116             For i = 1 To UBound(RecompensasFaccion)
118                 recompensa = RecompensasFaccion(i)

                    ' Como puede subir varios rangos todos juntos, nos aseguramos que
                    ' entregamos TODAS las recompensas hasta el rango actual desde la ultima recompensa.
120                 If recompensa.rank <= rank And recompensa.rank > ultimaRecompensa Then
                        ' Por alguna razon, PuedeUsarObjeto devuelve 0 cuando el usuario SI puede usarlo.
122                     If PuedeUsarObjeto(UserIndex, recompensa.ObjIndex, False) = 0 Then
124                         objetoRecompensa.amount = 1
126                         objetoRecompensa.ObjIndex = recompensa.ObjIndex

128                         If Not MeterItemEnInventario(UserIndex, objetoRecompensa) Then
130                             Call TirarItemAlPiso(.Pos, objetoRecompensa)

                            End If
                        End If
                    End If

132             Next i

                ' Guardamos que el usuario recibio las recompensas de su rank.
134             If .Faccion.ArmadaReal = 1 Then
136               .Faccion.RecibioArmaduraReal = rank
                Else
138               .Faccion.RecibioArmaduraCaos = rank
                End If

140             Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave(48, NO_3D_SOUND, NO_3D_SOUND))
            End With

            Exit Sub

DarRecompensas_Err:
142         Call RegistrarError(Err.Number, Err.Description, "ModFacciones.DarRecompensas", Erl)
144         Resume Next

End Sub


Private Sub PerderItemsFaccionarios(ByVal UserIndex As Integer)
            On Error GoTo PerderItemsFaccionarios_Err

            Dim i         As Byte
            Dim ItemIndex As Integer

100         With UserList(UserIndex)
102             For i = 1 To .CurrentInventorySlots
104                 ItemIndex = .Invent.Object(i).ObjIndex

106                 If ItemIndex > 0 Then

108                     If ObjData(ItemIndex).Real = 1 Or ObjData(ItemIndex).Caos = 1 Then

110                         Call QuitarUserInvItem(UserIndex, i, MAX_INVENTORY_OBJS)
112                         Call UpdateUserInv(False, UserIndex, i)

                        End If
                    End If
114             Next i
            End With

            Exit Sub

PerderItemsFaccionarios_Err:
116         Call RegistrarError(Err.Number, Err.Description, "ModFacciones.PerderItemsFaccionarios", Erl)
118         Resume Next

End Sub
