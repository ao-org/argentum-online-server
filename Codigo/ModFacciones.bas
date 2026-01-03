Attribute VB_Name = "ModFacciones"
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
'
Option Explicit

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
    On Error GoTo EnlistarArmadaReal_Err
    Dim charindexstr As String
    With UserList(UserIndex)
        charindexstr = str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex)
        If .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Then
            Call WriteLocaleChatOverHead(UserIndex, 1357, vbNullString, charindexstr, vbWhite)  ' Msg1357=Ya perteneces a mi ejército jóven soldado. Ve a combatir el caos en mis tierras para subir de rango en el Ejército Real.
            Exit Sub
        End If
        If .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then
            Call WriteLocaleChatOverHead(UserIndex, 1358, vbNullString, charindexstr, vbWhite)  ' Msg1358=¡Has llegado al lugar equivocado maldita escoria! Vete de aquí antes de ser encarcelado e interrogado.
            Exit Sub
        End If
        If Status(UserIndex) = e_Facciones.Criminal Then
            Call WriteLocaleChatOverHead(UserIndex, 1359, vbNullString, charindexstr, vbWhite)  ' Msg1359=No se permiten criminales en el Ejército Real.
            Exit Sub
        End If
        If .clase = e_Class.Thief Then
            Call WriteLocaleChatOverHead(UserIndex, 1360, vbNullString, charindexstr, vbWhite)  ' Msg1360=No hay lugar para escoria en el Ejército Real.
            Exit Sub
        End If
        'Si fue miembro de la Legión del Caos no se puede enlistar
        If .Faccion.ciudadanosMatados > 0 Then
            Call WriteLocaleChatOverHead(UserIndex, 1361, vbNullString, charindexstr, vbWhite)  ' Msg1361=Has derramado sangre inocente. Solo aceptamos almas nobles en el Ejército Real.
            Exit Sub
        End If
        If Not HayLugarEnInventario(UserIndex, -1, 999) Then
            Call WriteLocaleChatOverHead(UserIndex, 1362, vbNullString, charindexstr, vbWhite)  ' Msg1362=¡No tienes lugar suficiente en el inventario.
            Exit Sub
        End If
        If .Faccion.Reenlistadas > 0 Then
            Call WriteLocaleChatOverHead(UserIndex, 1363, vbNullString, charindexstr, vbWhite)  ' Msg1363=Ya has desertado el Ejército Real. No serás aceptado otra vez.
            Exit Sub
        End If
        Dim primerRango As t_RangoFaccion
        primerRango = RangosFaccion(1)
        If .Faccion.FactionScore < primerRango.RequiredScore Then
            Call WriteLocaleChatOverHead(UserIndex, 1378, primerRango.RequiredScore & "¬" & .Faccion.FactionScore, charindexstr, vbWhite) ' Msg1378=Para unirte a nuestras fuerzas debes tener al menos ¬1 puntos de faccion, solo tienes ¬2
            Exit Sub
        End If
        If .Stats.ELV < primerRango.NivelRequerido Then
            Call WriteLocaleChatOverHead(UserIndex, 1379, primerRango.NivelRequerido, charindexstr, vbWhite) ' Msg1379=¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel ¬1!!!
            Exit Sub
        End If
        'HarThaoS: Lo pongo al final para que lo expulse del clan solamente si cumple todos los requisitos.
        If .GuildIndex > 0 Then
            If PersonajeEsLeader(.Id) Then
                'Si el clan es neutral no lo dejo, le digo que tiene que salir del clan para poder enlistarse
                If GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_NEUTRAL Then
                    Call WriteLocaleChatOverHead(UserIndex, 1364, vbNullString, charindexstr, vbWhite)  ' Msg1364=No puedes integrar nuestras fuerzas si perteneces a un clan neutral, deberás abandonarlo si tu deseo es integrar el Ejército Real.
                    Exit Sub
                End If
            Else
                If GuildAlignmentIndex(.GuildIndex) <> e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                    Call m_EcharMiembroDeClan(UserIndex, .Id)
                End If
            End If
        End If
        ' Cumple con los requisitos para enlistarse
        .Faccion.RecompensasReal = primerRango.rank ' Asigna primer rango
        .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
        .Faccion.Status = e_Facciones.Armada
        If .Faccion.RecibioArmaduraReal = 0 Then
            Call WriteLocaleChatOverHead(UserIndex, 1365, vbNullString, charindexstr, vbWhite)  ' Msg1365=¡¡¡Bienvenido al Ejercito Imperial!!!, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.
            .Faccion.NivelIngreso = .Stats.ELV
            .Faccion.MatadosIngreso = .Faccion.ciudadanosMatados
        End If
        Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1932, vbNullString, e_FontTypeNames.FONTTYPE_INFOIAO)) ' Msg1932=¡Ahora perteneces al Ejército Real!
        Call DarRecompensas(UserIndex)
        Call RefreshCharStatus(UserIndex)
    End With
    'Call LogEjercitoReal(.Name & " ingresó el " & Date & " cuando era nivel " & .Stats.ELV)
    Exit Sub
EnlistarArmadaReal_Err:
    Call TraceError(Err.Number, Err.Description, "ModFacciones.EnlistarArmadaReal", Erl)
End Sub

' Subir de Rango y asignar recompensas.
Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
    On Error GoTo RecompensaArmadaReal_Err
    Dim npcCharIndex As String
    Dim proxRango    As t_RangoFaccion
    With UserList(UserIndex)
        npcCharIndex = str(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex)
        If .Faccion.RecompensasReal >= MaxRangoFaccion Then
            Call WriteLocaleChatOverHead(UserIndex, 1366, vbNullString, npcCharIndex, vbWhite)  ' Msg1366=Has alcanzado el máximo rango dentro de mis soldados. Demuestra tu bondad y liderazgo en el campo de batalla para algún día pertenecer al Consejo de Banderbill.
            Exit Sub
        End If
        If Not HayLugarEnInventario(UserIndex, -1, 999) Then
            Call WriteLocaleChatOverHead(UserIndex, 1367, vbNullString, npcCharIndex, vbWhite)  ' Msg1367=¡No tienes lugar suficiente en el inventario.
            Exit Sub
        End If
        proxRango = ProximoRango(UserIndex)
        If .Faccion.FactionScore < proxRango.RequiredScore Then
            Call WriteLocaleChatOverHead(UserIndex, 1380, proxRango.RequiredScore - .Faccion.FactionScore, npcCharIndex, vbWhite) ' Msg1380=Te faltan ¬1 Puntos de faccion para subir de rango.
            Exit Sub
        End If
        If proxRango.NivelRequerido > .Stats.ELV Then
            Call WriteLocaleChatOverHead(UserIndex, 1381, proxRango.NivelRequerido - .Stats.ELV, npcCharIndex, vbWhite) ' Msg1381=Has matado suficientes criminales pero, te faltan ¬1 niveles para poder recibir la próxima recompensa.
        Else ' El usuario cumple con los requerimientos de nivel, se le asigna la recomenza.
            .Faccion.RecompensasReal = proxRango.rank
            Call WriteLocaleChatOverHead(UserIndex, 1382, proxRango.Titulo, npcCharIndex, vbWhite) ' Msg1382=¡¡¡Aquí tienes tu recompensa ¬1!!!
            Call DarRecompensas(UserIndex)
            '.Stats.Exp = .Stats.Exp + ExpX100
        End If
    End With
    Exit Sub
RecompensaArmadaReal_Err:
    Call TraceError(Err.Number, Err.Description, "ModFacciones.RecompensaArmadaReal", Erl)
End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)
    On Error GoTo ExpulsarFaccionReal_Err
    UserList(UserIndex).Faccion.Status = e_Facciones.Ciudadano
    Call RefreshCharStatus(UserIndex)
    Call PerderItemsFaccionarios(UserIndex)
    Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1933, vbNullString, e_FontTypeNames.FONTTYPE_INFOIAO)) ' Msg1933=Has sido expulsado del Ejército Real.
    UserList(UserIndex).Faccion.FactionScore = 0
    Exit Sub
ExpulsarFaccionReal_Err:
    Call TraceError(Err.Number, Err.Description, "ModFacciones.ExpulsarFaccionReal", Erl)
End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)
    On Error GoTo ExpulsarFaccionCaos_Err
    UserList(UserIndex).Faccion.Status = e_Facciones.Criminal
    Call RefreshCharStatus(UserIndex)
    Call PerderItemsFaccionarios(UserIndex)
    Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1934, vbNullString, e_FontTypeNames.FONTTYPE_INFOIAO)) ' Msg1934=Has sido expulsado de la Legión Oscura.
    UserList(UserIndex).Faccion.FactionScore = 0
    Exit Sub
ExpulsarFaccionCaos_Err:
    Call TraceError(Err.Number, Err.Description, "ModFacciones.ExpulsarFaccionCaos", Erl)
End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As Byte
    On Error GoTo TituloReal_Err
    TituloReal = UserList(UserIndex).Faccion.RecompensasReal
    Exit Function
TituloReal_Err:
    Call TraceError(Err.Number, Err.Description, "ModFacciones.TituloReal", Erl)
End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
    On Error GoTo EnlistarCaos_Err
    Dim charindexstr As String
    With UserList(UserIndex)
        charindexstr = str(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex)
        If Status(UserIndex) = e_Facciones.Caos Or Status(UserIndex) = e_Facciones.concilio Then
            Call WriteLocaleChatOverHead(UserIndex, 1368, vbNullString, charindexstr, vbWhite)  ' Msg1368=Ya perteneces a la Legión Oscura.
            Exit Sub
        End If
        If Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.Ciudadano Or Status(UserIndex) = e_Facciones.consejo Then
            Call WriteLocaleChatOverHead(UserIndex, 1369, vbNullString, charindexstr, vbWhite)  ' Msg1369=¡¡Ja ja ja!! Tu no eres bienvenido aquí asqueroso Ciudadano
            Exit Sub
        End If
        If Not HayLugarEnInventario(UserIndex, -1, 999) Then
            Call WriteLocaleChatOverHead(UserIndex, 1370, vbNullString, charindexstr, vbWhite)  ' Msg1370=¡No tienes lugar suficiente en el inventario.
            Exit Sub
        End If
        If .clase = e_Class.Thief Then
            Call WriteLocaleChatOverHead(UserIndex, 1371, vbNullString, charindexstr, vbWhite)  ' Msg1371=¡La legión oscura no tiene lugar para escorias como tú! Los ladrones no son dignos de llevar nuestras armaduras.
            Exit Sub
        End If
        If UserList(UserIndex).Faccion.Reenlistadas > 0 Then
            Call WriteLocaleChatOverHead(UserIndex, 1372, vbNullString, charindexstr, vbWhite)  ' Msg1372=Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. ¡Vete de aquí!
            Exit Sub
        End If
        Dim primerRango As t_RangoFaccion
        primerRango = RangosFaccion(2) ' 2 es el primer rango del caos
        If .Faccion.FactionScore < primerRango.RequiredScore Then
            Call WriteLocaleChatOverHead(UserIndex, 1383, primerRango.RequiredScore & "," & .Faccion.FactionScore, charindexstr, vbWhite) ' Msg1383=Para unirte a nuestras fuerzas debes tener al menos ¬1 puntos de facción, solo tienes ¬2
            Exit Sub
        End If
        If .Stats.ELV < primerRango.NivelRequerido Then
            Call WriteLocaleChatOverHead(UserIndex, 1384, primerRango.NivelRequerido, charindexstr, vbWhite) ' Msg1384=¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel ¬1!!!
            Exit Sub
        End If
        'HarThaoS: Lo pongo al final para que lo expulse del clan solamente si cumple todos los requisitos.
        If .GuildIndex > 0 Then
            If PersonajeEsLeader(.Id) Then
                'Si el clan es neutral no lo dejo, le digo que tiene que salir del clan para poder enlistarse
                If GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_NEUTRAL Then
                    Call WriteLocaleChatOverHead(UserIndex, 1373, vbNullString, charindexstr, vbWhite)  ' Msg1373=No puedes integrar nuestras fuerzas si perteneces a un clan neutral, deberás abandonarlo si tu deseo es integrar la Legión Oscura.
                    Exit Sub
                End If
            Else
                If GuildAlignmentIndex(.GuildIndex) <> e_ALINEACION_GUILD.ALINEACION_CRIMINAL Then
                    Call m_EcharMiembroDeClan(UserIndex, .name)
                End If
            End If
        End If
        ' Cumple con los requisitos para enlistarse
        .Faccion.RecompensasCaos = primerRango.rank ' Asigna primer rango
        .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
        .Faccion.Status = e_Facciones.Caos
        If .Faccion.RecibioArmaduraCaos = 0 Then
            Call WriteLocaleChatOverHead(UserIndex, 1374, vbNullString, charindexstr, vbWhite)  ' Msg1374=Aquí tienes tu armadura legionario, ve a derramar sangre de los súbditos de Tancredo. Esta guerra será larga y cruel.
            .Faccion.NivelIngreso = .Stats.ELV
        End If
        Call WriteConsoleMsg(UserIndex, PrepareMessageLocaleMsg(1935, vbNullString, e_FontTypeNames.FONTTYPE_INFOIAO)) ' Msg1935=¡Ahora perteneces a la Legión Oscura!
        Call DarRecompensas(UserIndex)
        Call RefreshCharStatus(UserIndex)
    End With
    'Call LogEjercitoCaos(UserList(UserIndex).Name & " ingresó el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)
    Exit Sub
EnlistarCaos_Err:
    Call TraceError(Err.Number, Err.Description, "ModFacciones.EnlistarCaos", Erl)
End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
    On Error GoTo RecompensaCaos_Err
    Dim ciudadanosMatados As Long, npcCharIndex As String
    Dim proxRango         As t_RangoFaccion
    With UserList(UserIndex)
        ciudadanosMatados = .Faccion.ciudadanosMatados
        npcCharIndex = str(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex)
        If .Faccion.RecompensasCaos >= MaxRangoFaccion Then
            Call WriteLocaleChatOverHead(UserIndex, 1375, vbNullString, npcCharIndex, vbWhite)  ' Msg1375=¡Has alcanzado uno de los mejores lugares en mis filas. Mantén firme tu liderazgo y crueldad para algún día formar parte del Concilio de las Sombras.
            Exit Sub
        End If
        If Not HayLugarEnInventario(UserIndex, -1, 999) Then
            Call WriteLocaleChatOverHead(UserIndex, 1376, vbNullString, npcCharIndex, vbWhite)  ' Msg1376=¡No tienes lugar suficiente en el inventario.
            Exit Sub
        End If
        proxRango = ProximoRango(UserIndex)
        If .Faccion.FactionScore < proxRango.RequiredScore Then
            Call WriteLocaleChatOverHead(UserIndex, 1385, proxRango.RequiredScore - .Faccion.FactionScore, npcCharIndex, vbWhite) ' Msg1385=Te faltan ¬1 Puntos de faccion para subir de rango.
            Exit Sub
        End If
        If proxRango.NivelRequerido > .Stats.ELV Then
            Call WriteLocaleChatOverHead(UserIndex, 1386, proxRango.NivelRequerido - .Stats.ELV, npcCharIndex, vbWhite) ' Msg1386=Has acabado con la vida de suficientes enemigos pero aún te faltan ¬1 niveles para alcanzar el siguiente rango.
        Else ' El usuario cumple con los requerimientos de nivel, se le asigna la recomenza.
            .Faccion.RecompensasCaos = proxRango.rank
            Call WriteLocaleChatOverHead(UserIndex, 1387, proxRango.Titulo, npcCharIndex, vbWhite) ' Msg1387=¡¡¡Bien hecho ¬1, aquí tienes tu recompensa, sigue pregonando el caos a lo largo de estas tierras.!!!
            Call DarRecompensas(UserIndex)
            '.Stats.Exp = .Stats.Exp + ExpX100
        End If
    End With
    Exit Sub
RecompensaCaos_Err:
    Call TraceError(Err.Number, Err.Description, "ModFacciones.RecompensaCaos", Erl)
End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As Byte
    On Error GoTo TituloCaos_Err
    TituloCaos = UserList(UserIndex).Faccion.RecompensasCaos
    Exit Function
TituloCaos_Err:
    Call TraceError(Err.Number, Err.Description, "ModFacciones.TituloCaos", Erl)
End Function

' Devuelve el proximo rango para el usuario de la faccion que pertenece.
Private Function ProximoRango(ByVal UserIndex As Integer) As t_RangoFaccion
    On Error GoTo ProximoRango_Err
    With UserList(UserIndex)
        If .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo And .Faccion.RecompensasReal < MaxRangoFaccion Then
            ProximoRango = RangosFaccion(2 * .Faccion.RecompensasReal + 1)
        ElseIf .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio And .Faccion.RecompensasCaos < MaxRangoFaccion Then
            ProximoRango = RangosFaccion(2 * .Faccion.RecompensasCaos + 2)
        Else ' No pertenece a ninguna faccion.
            ' No devuelve nada
        End If
    End With
    Exit Function
ProximoRango_Err:
    Call TraceError(Err.Number, Err.Description, "ModFacciones.TituloCaos", Erl)
End Function

' Rutina para dar las recompensas de faccion al usuario.
' Si el usuario sube mas de un rango por vez, esta rutina le dara TODOS los objetos
' que debería tener hasta alcanzar su rango.
Private Sub DarRecompensas(ByVal UserIndex As Integer)
    On Error GoTo DarRecompensas_Err
    Dim recompensa       As t_RecompensaFaccion
    Dim rank             As Byte
    Dim ultimaRecompensa As Byte
    Dim objetoRecompensa As t_Obj
    Dim i                As Integer
    With UserList(UserIndex)
        ' Si es semidios o consejero, no le damos nada
        If .flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios) Then
            Exit Sub
        End If
        If .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Then
            rank = .Faccion.RecompensasReal
            ultimaRecompensa = .Faccion.RecibioArmaduraReal
        ElseIf .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then
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
                    objetoRecompensa.amount = 1
                    objetoRecompensa.ObjIndex = recompensa.ObjIndex
                    objetoRecompensa.ElementalTags = e_ElementalTags.Normal
                    If Not MeterItemEnInventario(UserIndex, objetoRecompensa) Then
                        Call TirarItemAlPiso(.pos, objetoRecompensa)
                    End If
                End If
            End If
        Next i
        ' Guardamos que el usuario recibio las recompensas de su rank.
        If .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Then
            .Faccion.RecibioArmaduraReal = rank
        ElseIf .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then
            .Faccion.RecibioArmaduraCaos = rank
        End If
        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave(48, NO_3D_SOUND, NO_3D_SOUND))
    End With
    Exit Sub
DarRecompensas_Err:
    Call TraceError(Err.Number, Err.Description, "ModFacciones.DarRecompensas", Erl)
End Sub

Private Sub PerderItemsFaccionarios(ByVal UserIndex As Integer)
    On Error GoTo PerderItemsFaccionarios_Err
    Dim i         As Byte
    Dim ItemIndex As Integer
    With UserList(UserIndex)
        For i = 1 To .CurrentInventorySlots
            ItemIndex = .invent.Object(i).ObjIndex
            If ItemIndex > 0 Then
                If ObjData(ItemIndex).Real = 1 Or ObjData(ItemIndex).Caos = 1 Then
                    Call QuitarUserInvItem(UserIndex, i, GetMaxInvOBJ())
                    Call UpdateUserInv(False, UserIndex, i)
                End If
            End If
        Next i
    End With
    Exit Sub
PerderItemsFaccionarios_Err:
    Call TraceError(Err.Number, Err.Description, "ModFacciones.PerderItemsFaccionarios", Erl)
End Sub
Public Sub NotifyConnectionToFaction(ByVal UserIndex As Integer)
    On Error GoTo NotifyConnectionToFaction_Err
    With UserList(UserIndex)
        ' Excluir GMs
        Select Case .flags.Privilegios
            Case e_PlayerType.Admin, e_PlayerType.Dios, e_PlayerType.SemiDios, e_PlayerType.Consejero
                Exit Sub
        End Select
        
        Dim msgId As Integer
        msgId = GetRandomFactionMsgId(.faccion.Status)

        Select Case .faccion.Status
            Case e_Facciones.Armada
                If msgId > 0 Then
                    Call SendData(SendTarget.ToRealYRMs, 0, _
                        PrepareMessageLocaleMsg(msgId, .Name, e_FontTypeNames.FONTTYPE_CITIZEN_ARMADA))
                End If

            Case e_Facciones.consejo
                If msgId > 0 Then
                    Call SendData(SendTarget.ToRealYRMs, 0, _
                        PrepareMessageLocaleMsg(msgId, .Name, e_FontTypeNames.FONTTYPE_CONSEJO))
                End If

            Case e_Facciones.Caos
                If msgId > 0 Then
                    Call SendData(SendTarget.ToCaosYRMs, 0, _
                        PrepareMessageLocaleMsg(msgId, .Name, e_FontTypeNames.FONTTYPE_CRIMINAL_CAOS))
                End If

            Case e_Facciones.concilio
                If msgId > 0 Then
                    Call SendData(SendTarget.ToCaosYRMs, 0, _
                        PrepareMessageLocaleMsg(msgId, .Name, e_FontTypeNames.FONTTYPE_CONSEJOCAOS))
                End If
        End Select
    End With
    Exit Sub
NotifyConnectionToFaction_Err:
    Call TraceError(Err.Number, Err.Description, "ModFacciones.NotifyConnectionToFaction", Erl)
End Sub
Private Function GetRandomFactionMsgId(ByVal faction As e_Facciones) As Integer
    On Error GoTo GetRandomFactionMsgId_Err
    Select Case faction
        Case e_Facciones.Armada, e_Facciones.consejo
            Select Case RandomNumber(1, 5)
                Case 1: GetRandomFactionMsgId = 2133
                Case 2: GetRandomFactionMsgId = 2134
                Case 3: GetRandomFactionMsgId = 2135
                Case 4: GetRandomFactionMsgId = 2136
                Case 5: GetRandomFactionMsgId = 2137
            End Select

        Case e_Facciones.Caos, e_Facciones.concilio
            Select Case RandomNumber(1, 5)
                Case 1: GetRandomFactionMsgId = 2138
                Case 2: GetRandomFactionMsgId = 2139
                Case 3: GetRandomFactionMsgId = 2140
                Case 4: GetRandomFactionMsgId = 2141
                Case 5: GetRandomFactionMsgId = 2142
            End Select

        Case Else
            GetRandomFactionMsgId = 0
    End Select
    Exit Function
GetRandomFactionMsgId_Err:
    Call TraceError(Err.Number, Err.Description, "ModFacciones.GetRandomFactionMsgId", Erl)
End Function

