Attribute VB_Name = "ModFacciones"
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

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
        On Error GoTo EnlistarArmadaReal_Err

100     If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
102         Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a las tropas reales!!! Ve a combatir criminales", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

104     If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
106         Call WriteChatOverHead(UserIndex, "¡¡¡Maldito insolente!!! vete de aqui seguidor de las sombras", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

108     If Status(UserIndex) = 2 Then
110         Call WriteChatOverHead(UserIndex, "¡¡¡No se permiten criminales en el ejercito imperial!!!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

112     If Status(UserIndex) = 0 Then
114         Call WriteChatOverHead(UserIndex, "¡¡¡No se permiten neutrales en el ejercito imperial, antes debes solicitar el perdón con un sacerdote!!!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

        If UserList(UserIndex).clase = eClass.Thief Then
            Call WriteChatOverHead(UserIndex, "¡Los ladrones no son dignos para pertenecer al ejercito imperial!!!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

120     If UserList(UserIndex).Faccion.CriminalesMatados < 100 Then
122         Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 100 criminales, solo has matado " & UserList(UserIndex).Faccion.CriminalesMatados, str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

        If UserList(UserIndex).Stats.ELV < 25 Then
            Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

        'Si fue miembro de la Legión del Caos no se puede enlistar
124     If UserList(UserIndex).Faccion.ciudadanosMatados > 0 Then
126         Call WriteChatOverHead(UserIndex, "¡Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

128     If UserList(UserIndex).Faccion.Reenlistadas > 0 Then
            If UserList(UserIndex).Faccion.Reenlistadas = 200 Then
                Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas reales demasiadas veces!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Else
130             Call WriteChatOverHead(UserIndex, "¡Has traicionado al ejercito imperial una vez, no aceptamos traidores!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)

            End If

            Exit Sub

        End If

        ' Cumple con los requisitos para enlistarse
132     UserList(UserIndex).Faccion.ArmadaReal = 1
134     UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1
138     UserList(UserIndex).Faccion.Status = 3
        UserList(UserIndex).Faccion.RecompensasReal = 1 ' Asigna primer rango
        UserList(UserIndex).Faccion.NextRecompensa = 250

144     If UserList(UserIndex).Faccion.RecibioArmaduraReal = 0 Then
            Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al Ejercito Imperial!!!, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)

166         UserList(UserIndex).Faccion.NivelIngreso = UserList(UserIndex).Stats.ELV
168         UserList(UserIndex).Faccion.FechaIngreso = Date
            'Esto por ahora es inútil, siempre va a ser cero, pero bueno, despues va a servir.
170         UserList(UserIndex).Faccion.MatadosIngreso = UserList(UserIndex).Faccion.ciudadanosMatados

        End If

        Call WriteConsoleMsg(UserIndex, "¡Ahora perteneces al Ejercito Real!", FontTypeNames.FONTTYPE_INFOIAO)
        Call DarRecompensas(UserIndex)
        Call RefreshCharStatus(UserIndex)

        'Call LogEjercitoReal(UserList(UserIndex).Name & " ingresó el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

        Exit Sub

EnlistarArmadaReal_Err:
176     Call RegistrarError(Err.Number, Err.Description, "ModFacciones.EnlistarArmadaReal", Erl)
178     Resume Next

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)

        On Error GoTo RecompensaArmadaReal_Err

        '***************************************************
        'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 23/01/2007
        'Handles the way of gaining new ranks in the "Armada Real"
        '***************************************************
        Dim Crimis    As Long
        Dim Lvl       As Byte
        Dim NextRecom As Long

100     Lvl = UserList(UserIndex).Stats.ELV
102     Crimis = UserList(UserIndex).Faccion.CriminalesMatados
104     NextRecom = UserList(UserIndex).Faccion.NextRecompensa

106     If Crimis < NextRecom Then
108         Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - Crimis & " Criminales más para recibir la próxima Recompensa", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

        Dim nivelRequerido As Integer
        Dim criminalesNecesariosParaFuturaRecomensa As Integer
        Dim Rango As Byte
        Rango = UserList(UserIndex).Faccion.RecompensasReal

        Select Case Rango
            Case 0: ' Antes de enlistarse.
                nivelRequerido = 25
                criminalesNecesariosParaFuturaRecomensa = 100

            Case 1:
                nivelRequerido = 30
                criminalesNecesariosParaFuturaRecomensa = 250

            Case 2:
                nivelRequerido = 35
                criminalesNecesariosParaFuturaRecomensa = 400

            Case 3:
                nivelRequerido = 40
                criminalesNecesariosParaFuturaRecomensa = 750

            Case 4:
                nivelRequerido = 43
                criminalesNecesariosParaFuturaRecomensa = 1350

            Case Else:
                Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores Soldados. Mataste " & Crimis & ", sigue asi. Ya no tengo más recompensa para darte que mi agradescimiento. ¡Felicidades!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub

        End Select

        If nivelRequerido - Lvl > 0 Then
          ' Al usuario le falta subir niveles

          Call WriteChatOverHead(UserIndex, "Mataste Suficientes Criminales, pero te faltan " & (nivelRequerido - Lvl) & " Niveles para poder recibir la próxima Recompensa", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Else
          ' El usuario cumple con los requerimientos de nivel, se le asigna la recomenza.

          UserList(UserIndex).Faccion.RecompensasReal = IIf(Rango < 5, Rango + 1, Rango)
          UserList(UserIndex).Faccion.NextRecompensa = criminalesNecesariosParaFuturaRecomensa

          Call WriteChatOverHead(UserIndex, "¡¡¡Aqui tienes tu recompensa " + TituloReal(UserIndex) + "!!!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
          Call DarRecompensas(UserIndex)
          'UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpX100
        End If

        Exit Sub

RecompensaArmadaReal_Err:
156     Call RegistrarError(Err.Number, Err.Description, "ModFacciones.RecompensaArmadaReal", Erl)
158     Resume Next

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

        '***************************************************
        'Autor: Unknown
        'Last Modification: 23/01/2007 Pablo (ToxicWaste)
        'Handles the titles of the members of the "Armada Real"
        '***************************************************
        Select Case UserList(UserIndex).Faccion.RecompensasReal

            Case 1
                TituloReal = "Soldado"

            Case 2
                TituloReal = "Caballero"

            Case 3
                TituloReal = "Capitán"

            Case 4
                TituloReal = "Protector del Reino"

            Case 5
                TituloReal = "Campeón de la Luz"

            Case 6
                TituloReal = "Consejo de Banderbill"

        End Select

        Exit Function

TituloReal_Err:
120     Call RegistrarError(Err.Number, Err.Description, "ModFacciones.TituloReal", Erl)
122     Resume Next
        
End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
        '***************************************************
        'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 23/01/2007
        'Handles the entrance of users to the "Legión Oscura"
        '***************************************************

        On Error GoTo EnlistarCaos_Err


100     If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
102         Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a la legión oscura!!!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

104     If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
106         Call WriteChatOverHead(UserIndex, "Las sombras reinarán en Argentum. ¡¡¡Fuera de aqui insecto Real!!!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

        'Si fue miembro de la Armada Real no se puede enlistar
108     If UserList(UserIndex).Faccion.RecompensasReal > 0 Then
110         Call WriteChatOverHead(UserIndex, "No permitiré que ningún insecto real ingrese a mis tropas.", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

112     If Status(UserIndex) = 1 Then
114         Call WriteChatOverHead(UserIndex, "¡¡Ja ja ja!! Tu no eres bienvenido aquí asqueroso Ciudadano", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

        If UserList(UserIndex).clase = eClass.Thief Then
            Call WriteChatOverHead(UserIndex, "¡La legión oscura no tiene lugar para escorias como tú! Los ladrones no son dignos de llevar nuestras armaduras.", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

116     If UserList(UserIndex).Faccion.ciudadanosMatados < 100 Then
118         Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 100 usuarios, solo has matado " & UserList(UserIndex).Faccion.ciudadanosMatados, str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

120     If UserList(UserIndex).Stats.ELV < 25 Then
122         Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

124     If UserList(UserIndex).Faccion.Reenlistadas > 0 Then
126         If UserList(UserIndex).Faccion.Reenlistadas = 200 Then
128             Call WriteChatOverHead(UserIndex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. ¡Vete de aquí!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Else
130             Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas oscuras demasiadas veces!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)

            End If

            Exit Sub

        End If

132     UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1
134     UserList(UserIndex).Faccion.FuerzasCaos = 1
136     UserList(UserIndex).Faccion.Status = 2
        UserList(UserIndex).Faccion.RecompensasCaos = 1 ' Asigna primer rango
        UserList(UserIndex).Faccion.NextRecompensa = 150

144     If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0 Then
            Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al lado oscuro!!! Aqui tienes tus armaduras. Derrama sangre Ciudadana y Real y serás recompensado, lo prometo.", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)

            UserList(UserIndex).Faccion.NivelIngreso = UserList(UserIndex).Stats.ELV
            UserList(UserIndex).Faccion.FechaIngreso = Date

        End If

        Call WriteConsoleMsg(UserIndex, "¡Ahora perteneces al caos!", FontTypeNames.FONTTYPE_INFOIAO)
        Call RefreshCharStatus(UserIndex)
        Call DarRecompensas(UserIndex)

        'Call LogEjercitoCaos(UserList(UserIndex).Name & " ingresó el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

        Exit Sub

EnlistarCaos_Err:
174     Call RegistrarError(Err.Number, Err.Description, "ModFacciones.EnlistarCaos", Erl)
176     Resume Next

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)

        On Error GoTo RecompensaCaos_Err

        '***************************************************
        'Author: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 23/01/2007
        'Handles the way of gaining new ranks in the "Legión Oscura"
        '***************************************************
        Dim ciudadanosMatados As Integer
        Dim Lvl       As Byte
        Dim NextRecom As Integer

100     Lvl = UserList(UserIndex).Stats.ELV
102     ciudadanosMatados = UserList(UserIndex).Faccion.ciudadanosMatados
104     NextRecom = UserList(UserIndex).Faccion.NextRecompensa

106     If ciudadanosMatados < NextRecom Then
108         Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - ciudadanosMatados & " Cuidadanos más para recibir la próxima Recompensa", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If

        Dim nivelRequerido As Integer
        Dim ciudadanosNecesariosParaFuturaRecomensa As Integer
        Dim Rango As Byte
        Rango = UserList(UserIndex).Faccion.RecompensasCaos

        Select Case Rango
            Case 0: ' Antes de enlistarse
                nivelRequerido = 25
                ciudadanosNecesariosParaFuturaRecomensa = 100

            Case 1:
                nivelRequerido = 30
                ciudadanosNecesariosParaFuturaRecomensa = 250

            Case 2:
                nivelRequerido = 35
                ciudadanosNecesariosParaFuturaRecomensa = 400

            Case 3:
                nivelRequerido = 40
                ciudadanosNecesariosParaFuturaRecomensa = 750

            Case 4:
                nivelRequerido = 43
                ciudadanosNecesariosParaFuturaRecomensa = 1350

            Case Else:
                Call WriteChatOverHead(UserIndex, "¡Felicitaciones, eres de mis mejores guerreros, estas en lo más alto!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)

                Exit Sub

        End Select

        If nivelRequerido - Lvl > 0 Then
          ' Al usuario le falta subir niveles

          Call WriteChatOverHead(UserIndex, "Mataste Suficientes Ciudadanos, pero te faltan " & (nivelRequerido - Lvl) & " Niveles para poder recibir la próxima Recompensa", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
        Else
          ' El usuario cumple con los requerimientos de nivel, se le asigna la recomenza.

          UserList(UserIndex).Faccion.RecompensasCaos = IIf(Rango < 5, Rango + 1, Rango)
          UserList(UserIndex).Faccion.NextRecompensa = ciudadanosNecesariosParaFuturaRecomensa

          Call WriteChatOverHead(UserIndex, "¡¡¡Bien hecho " + TituloCaos(UserIndex) + ", aquí tienes tu recompensa!!!", str(NpcList(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
          Call DarRecompensas(UserIndex)
        End If

        Exit Sub

RecompensaCaos_Err:
142     Call RegistrarError(Err.Number, Err.Description, "ModFacciones.RecompensaCaos", Erl)
144     Resume Next

End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String

        On Error GoTo TituloCaos_Err

        Select Case UserList(UserIndex).Faccion.RecompensasCaos
            Case 1
                TituloCaos = "Acólito"

            Case 2
                TituloCaos = "Emisario del Caos"

            Case 3
                TituloCaos = "Sanguinario"

            Case 4
                TituloCaos = "Caballero de la Oscuridad"

            Case 5
                TituloCaos = "Devorador de Almas"

            Case 6
                TituloCaos = "Concilio de las Sombras"

        End Select

        Exit Function

TituloCaos_Err:
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
                    If PuedeUsarObjeto(UserIndex, recompensa.ObjIndex) = 0 Then
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
