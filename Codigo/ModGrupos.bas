Attribute VB_Name = "ModGrupos"
' Argentum 20 Game Server
'
'    Copyright (C) 2023-2026 Noland Studios LTD
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
Public Grupo            As Tgrupo
Private UniqueIdCounter As Long

Private Function GetNextId() As Long
    UniqueIdCounter = (UniqueIdCounter + 1) And &H7FFFFFFF
    GetNextId = UniqueIdCounter
End Function

Public Sub InvitarMiembro(ByVal UserIndex As Integer, ByVal InvitadoIndex As Integer)
    On Error GoTo InvitarMiembro_Err
    Dim skillsNecesarios As Integer
    Dim Remitente        As t_User: Remitente = UserList(UserIndex)
    Dim Invitado         As t_User: Invitado = UserList(InvitadoIndex)
    ' Comentando linea importante abajo, solo temporalmente hasta que el sistema de grupo nuevo este implementado 3/5/2025 - ako
    ' skillsNecesarios = 15 - Remitente.Stats.UserAtributos(e_Atributos.Carisma) \ 2
    skillsNecesarios = 0
    If Remitente.Stats.UserSkills(e_Skill.liderazgo) < skillsNecesarios Then
        Call WriteLocaleMsg(UserIndex, MSG_FALTAN_PUNTOS_LIDERAZGO_LIDERAR_GRUPO, (skillsNecesarios - Remitente.Stats.UserSkills(e_Skill.liderazgo)), e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2041="Te faltan ¬1 puntos en Liderazgo para liderar un grupo."
        Exit Sub
    End If
    'HarThaoS: Si invita a un gm no lo dejo
    If EsGM(InvitadoIndex) Then
        Call WriteLocaleMsg(UserIndex, MSG_PUEDES_INVITAR_GRUPO_GM, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2042="No puedes invitar a un grupo a un GM."
        Exit Sub
    End If
    'Si es gm tampoco lo dejo
    If EsGM(UserIndex) Then
        Call WriteLocaleMsg(UserIndex, MSG_NO_GMS_PUEDEN_FORMAR_PARTE_GRUPO, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2043="Los GMs no pueden formar parte de un grupo."
        Exit Sub
    End If
    If Invitado.flags.SeguroParty Then
        Call WriteLocaleMsg(UserIndex, MSG_USUARIO_DEBE_DESACTIVAR_SEGURO_GRUPOS_PODER_INVITARLO, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2044="El usuario debe desactivar el seguro de grupos para poder invitarlo."
        Exit Sub
    End If
    If Remitente.Grupo.CantidadMiembros >= UBound(Remitente.Grupo.Miembros) Then
        Call WriteLocaleMsg(UserIndex, MSG_PUEDES_INVITAR_MAS_PERSONAS_LIMITE, CStr(UBound(Remitente.Grupo.Miembros)), e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2045="No puedes invitar a mas personas. (Límite: ¬1)"
        Exit Sub
    End If
    If (Status(UserIndex) = 0 And Status(InvitadoIndex) = 1) Or (Status(UserIndex) = 1 And Status(InvitadoIndex) = 0) Or (Status(UserIndex) = 1 And Status(InvitadoIndex) = 2) Or _
            (Status(UserIndex) = 2 And Status(InvitadoIndex) = 1) Or (Status(UserIndex) = 3 And Status(InvitadoIndex) = 0) Or (Status(UserIndex) = 0 And Status(InvitadoIndex) = _
            3) Or (Status(UserIndex) = 3 And Status(InvitadoIndex) = 2) Or (Status(UserIndex) = 2 And Status(InvitadoIndex) = 3) Then
        Call WriteLocaleMsg(UserIndex, MSG_PODES_CREAR_GRUPO_PERSONAJES_DIFERENTES_FACCIONES, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2046="No podes crear un grupo con personajes de diferentes facciones."
        Exit Sub
    End If
    If (CInt(Remitente.Stats.UserSkills(e_Skill.liderazgo)) >= (15 - Remitente.Stats.UserAtributos(e_Atributos.Carisma) / 2)) Then
        ' Si el lider tiene liderazgo asignado segun su raza, se permite una diferencia de 1 nivel mas
        If Abs(CInt(Invitado.Stats.ELV) - CInt(Remitente.Stats.ELV)) > (SvrConfig.GetValue("PartyELVwLeadership")) Then
            'Msg1438=No podes crear un grupo con personajes con diferencia de más de ¬1 niveles.
            Call WriteLocaleMsg(UserIndex, MSG_NO_PODES_CREAR_GRUPO_PERSONAJES_DIFERENCIA_MAS_NIVELES, e_FontTypeNames.FONTTYPE_New_GRUPO, SvrConfig.GetValue("PartyELVwLeadership"))
            Exit Sub
        End If
    Else
        If Abs(CInt(Invitado.Stats.ELV) - CInt(Remitente.Stats.ELV)) > SvrConfig.GetValue("PartyELV") Then
            'Msg1438=No podes crear un grupo con personajes con diferencia de más de ¬1 niveles.
            Call WriteLocaleMsg(UserIndex, MSG_NO_PODES_CREAR_GRUPO_PERSONAJES_DIFERENCIA_MAS_NIVELES, e_FontTypeNames.FONTTYPE_New_GRUPO, SvrConfig.GetValue("PartyELV"))
            Exit Sub
        End If
    End If
    If Invitado.Grupo.EnGrupo Then
        Call WriteLocaleMsg(UserIndex, MSG_CHARACTER_ALREADY_IN_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
        Exit Sub
    End If
    If UserList(InvitadoIndex).flags.RespondiendoPregunta = False Then
        Call WriteLocaleMsg(UserIndex, MSG_GROUP_REQUEST_SENT_WAIT_RESPONSE, e_FontTypeNames.FONTTYPE_New_GRUPO)
        Call WriteLocaleMsg(InvitadoIndex, MSG_INVITO_UNIRSE_SU_GRUPO, e_FontTypeNames.FONTTYPE_New_GRUPO, Remitente.name) ' Msg2049="¬1 te invitó a unirse a su grupo."
        With UserList(InvitadoIndex)
            Call SetUserRef(.Grupo.PropuestaDe, UserIndex)
            .flags.pregunta = 1
            Call SetUserRef(.Grupo.Lider, UserIndex)
        End With
        Call WritePreguntaBox(InvitadoIndex, 1595, Remitente.name) 'Msg1595= ¬1 te invito a unirse a su grupo. ¿Deseas unirte?
        UserList(InvitadoIndex).flags.RespondiendoPregunta = True
    Else
        Call WriteLocaleMsg(UserIndex, MSG_USUARIO_TIENE_SOLICITUD_PENDIENTE, e_FontTypeNames.FONTTYPE_INFO) ' Msg2050="El usuario tiene una solicitud pendiente."
    End If
    Exit Sub
InvitarMiembro_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.InvitarMiembro", Erl)
End Sub
Public Sub EcharMiembro(ByVal UserIndex As Integer, ByVal Indice As Byte)
    On Error GoTo EcharMiembro_Err
    Dim i              As Long
    Dim LoopC          As Long
    Dim indexviejo     As Byte
    Dim UserIndexEchar As Integer
    Dim GroupLider     As Integer
    With UserList(UserIndex).Grupo
        GroupLider = .Lider.ArrayIndex
        If Not .EnGrupo Then
            Call WriteLocaleMsg(UserIndex, MSG_NINGUN_GRUPO, e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
        If .Lider.ArrayIndex <> UserIndex Then
            Call WriteLocaleMsg(UserIndex, MSG_PODES_ECHAR_USUARIOS_GRUPO, e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
        UserIndexEchar = UserList(.Lider.ArrayIndex).Grupo.Miembros(Indice + 1).ArrayIndex
        If UserIndexEchar = UserIndex Then
            Call WriteLocaleMsg(UserIndex, MSG_PODES_EXPULSARTE_TI_MISMO, e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
        
        ' Guardar cantidad original antes de modificar
        Dim OriginalCount As Integer
        OriginalCount = .CantidadMiembros
        
        For i = 1 To UBound(.Miembros)
            If UserIndexEchar = .Miembros(i).ArrayIndex Then
                Call ClearUserRef(.Miembros(i))
                indexviejo = i
                For LoopC = indexviejo To 5
                    .Miembros(LoopC) = .Miembros(LoopC + 1)
                Next LoopC
                i = UBound(.Miembros)
                Call ClearUserRef(.Miembros(i))
                Exit For
            End If
        Next i
        .CantidadMiembros = .CantidadMiembros - 1
        
        ' Enviar x=0 para ocultar marcadores del miembro expulsado
        Dim j As Integer
        For j = 1 To OriginalCount
            Call WriteUbicacion(UserIndexEchar, j, 0)
        Next j
    End With
    
    With UserList(UserIndexEchar)
        Call WriteLocaleMsg(UserIndex, MSG_FUE_EXPULSADO_GRUPO, e_FontTypeNames.FONTTYPE_New_GRUPO, .Name)
        Call WriteLocaleMsg(UserIndexEchar, "37", e_FontTypeNames.FONTTYPE_New_GRUPO)
        .Grupo.EnGrupo = False
        Call SetUserRef(.Grupo.Lider, 0)
        Call SetUserRef(.Grupo.PropuestaDe, 0)
        .Grupo.CantidadMiembros = 0
        Call SetUserRef(.Grupo.Miembros(1), 0)
        Call RefreshCharStatus(UserIndexEchar)
        .Grupo.Id = -1
        If MapInfo(.pos.Map).OnlyGroups And MapInfo(.pos.Map).Salida.Map <> 0 Then
            Call WriteLocaleMsg(UserIndexEchar, MSG_DEBES_ESTAR_GRUPO_PERMANECER_MAPA, e_FontTypeNames.FONTTYPE_INFO)
            Call WarpUserChar(UserIndexEchar, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
        End If
    End With
    
    With UserList(UserIndex).Grupo
        ' Enviar posiciones actualizadas a todos los miembros restantes
        Dim a As Integer
        Dim b As Integer
        For a = 1 To .CantidadMiembros
            Dim TargetMemberIndex As Integer
            TargetMemberIndex = .Miembros(a).ArrayIndex
            
            ' PRIMERO: Limpiar todos los marcadores de este miembro (incluido el del que fue expulsado)
            For b = 1 To OriginalCount
                Call WriteUbicacion(TargetMemberIndex, b, 0)
            Next b
            
            ' LUEGO: Enviar solo las posiciones de los miembros que quedan (excepto el suyo propio)
            For b = 1 To .CantidadMiembros
                Dim SourceMemberIndex As Integer
                SourceMemberIndex = .Miembros(b).ArrayIndex
                
                ' NO enviar la posición de un miembro a sí mismo.
                ' Solo compartir ubicación si ambos miembros están en el mismo mapa;
                ' en caso contrario, mantener el marcador limpio.
                If TargetMemberIndex <> SourceMemberIndex Then
                    If UserList(TargetMemberIndex).Pos.Map = UserList(SourceMemberIndex).Pos.Map Then
                        Call WriteUbicacion(TargetMemberIndex, b, SourceMemberIndex)
                    Else
                        Call WriteUbicacion(TargetMemberIndex, b, 0)
                    End If
                End If
            Next b
        Next a
        
        If .CantidadMiembros = 1 Then
            Call WriteLocaleMsg(UserIndex, MSG_GROUP_DISBANDED_ALL_MEMBERS_LEFT, e_FontTypeNames.FONTTYPE_New_GRUPO)
            .EnGrupo = False
            Call SetUserRef(.Lider, 0)
            Call SetUserRef(.PropuestaDe, 0)
            .CantidadMiembros = 0
            Call SetUserRef(.Miembros(1), 0)
            .Id = -1
            Call modSendData.SendData(ToIndex, UserIndex, PrepareUpdateGroupInfo(UserIndex))
            Dim LiderMap As Integer: LiderMap = UserList(UserIndex).pos.Map
            If MapInfo(LiderMap).OnlyGroups And MapInfo(LiderMap).Salida.Map <> 0 Then
                Call WriteLocaleMsg(UserIndex, MSG_DEBES_ESTAR_GRUPO_PERMANECER_MAPA_2056, e_FontTypeNames.FONTTYPE_INFO)
                Call WarpUserChar(UserIndex, MapInfo(LiderMap).Salida.Map, MapInfo(LiderMap).Salida.x, MapInfo(LiderMap).Salida.y, True)
            End If
        End If
    End With
    
    Call RefreshCharStatus(UserIndex)
    Call modSendData.SendData(ToGroup, GroupLider, PrepareUpdateGroupInfo(GroupLider))
    Call modSendData.SendData(ToIndex, UserIndexEchar, PrepareUpdateGroupInfo(UserIndexEchar))
    Exit Sub
EcharMiembro_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.EcharMiembro", Erl)
End Sub

Public Sub SalirDeGrupo(ByVal UserIndex As Integer)
    On Error GoTo SalirDeGrupo_Err
    Dim i          As Long
    Dim LoopC      As Long
    Dim indexviejo As Byte
    With UserList(UserIndex)
        If Not .Grupo.EnGrupo Then
            Call WriteLocaleMsg(UserIndex, MSG_NINGUN_GRUPO_2057, e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
        
        ' Guardar cantidad original antes de modificar
        Dim OriginalCount As Integer
        OriginalCount = UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros
        
        .Grupo.EnGrupo = False
        .Grupo.Id = -1
        For i = 1 To UBound(.Grupo.Miembros)
            If .name = UserList(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex).name Then
                Call ClearUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i))
                indexviejo = i
                For LoopC = indexviejo To 5
                    UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(LoopC) = UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(LoopC + 1)
                Next LoopC
                i = UBound(.Grupo.Miembros)
                Call ClearUserRef(.Grupo.Miembros(i))
                Exit For
            End If
        Next i
        UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros = UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros - 1
        
        ' Enviar x=0 para ocultar marcadores del usuario que sale
        Dim j As Integer
        For j = 1 To OriginalCount
            Call WriteUbicacion(UserIndex, j, 0)
        Next j
        
        ' Enviar posiciones actualizadas a todos los miembros restantes
        Dim a As Integer
        Dim b As Integer
        For a = 1 To UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros
            Dim TargetMemberIndex As Integer
            TargetMemberIndex = UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(a).ArrayIndex
            
            ' PRIMERO: Limpiar todos los marcadores de este miembro (incluido el del que se fue)
            For b = 1 To OriginalCount
                Call WriteUbicacion(TargetMemberIndex, b, 0)
            Next b
            
            ' LUEGO: Enviar solo las posiciones de los miembros que quedan (excepto el suyo propio)
            For b = 1 To UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros
                Dim SourceMemberIndex As Integer
                SourceMemberIndex = UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(b).ArrayIndex
                
                ' NO enviar la posición de un miembro a sí mismo
                If TargetMemberIndex <> SourceMemberIndex Then
                    ' Solo compartir ubicación entre miembros que estén en el mismo mapa.
                    ' Si están en mapas distintos, el marcador ya quedó limpio en el paso anterior.
                    If UserList(TargetMemberIndex).pos.Map = UserList(SourceMemberIndex).pos.Map Then
                        Call WriteUbicacion(TargetMemberIndex, b, SourceMemberIndex)
                    End If
                End If
            Next b
        Next a
        
        Call WriteLocaleMsg(UserIndex, MSG_REMOVED_FROM_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO)
        Call WriteLocaleMsg(.Grupo.Lider.ArrayIndex, MSG_USER_LEFT_GROUP, e_FontTypeNames.FONTTYPE_New_GRUPO, .name)
        If UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros = 1 Then
            Call WriteLocaleMsg(.Grupo.Lider.ArrayIndex, MSG_GROUP_DISBANDED_ALL_MEMBERS_LEFT, e_FontTypeNames.FONTTYPE_New_GRUPO)
            Call WriteUbicacion(.Grupo.Lider.ArrayIndex, 1, 0)
            UserList(.Grupo.Lider.ArrayIndex).Grupo.Id = -1
            UserList(.Grupo.Lider.ArrayIndex).Grupo.EnGrupo = False
            Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Lider, 0)
            Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.PropuestaDe, 0)
            UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros = 0
            Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(1), 0)
            Call RefreshCharStatus(.Grupo.Lider.ArrayIndex)
            Call modSendData.SendData(ToIndex, .Grupo.Lider.ArrayIndex, PrepareUpdateGroupInfo(.Grupo.Lider.ArrayIndex))
            Dim LiderMap As Integer: LiderMap = UserList(.Grupo.Lider.ArrayIndex).pos.Map
            If MapInfo(LiderMap).OnlyGroups And MapInfo(LiderMap).Salida.Map <> 0 Then
                Call WriteLocaleMsg(.Grupo.Lider.ArrayIndex, MSG_DEBES_ESTAR_GRUPO_PERMANECER_MAPA_2059, e_FontTypeNames.FONTTYPE_INFO)
                Call WarpUserChar(.Grupo.Lider.ArrayIndex, MapInfo(LiderMap).Salida.Map, MapInfo(LiderMap).Salida.x, MapInfo(LiderMap).Salida.y, True)
            End If
        End If
        Call WriteUbicacion(UserIndex, 1, 0)
        Call modSendData.SendData(ToGroup, .Grupo.Lider.ArrayIndex, PrepareUpdateGroupInfo(.Grupo.Lider.ArrayIndex))
        Call SetUserRef(.Grupo.Lider, 0)
        Call modSendData.SendData(ToIndex, UserIndex, PrepareUpdateGroupInfo(UserIndex))
        If MapInfo(.pos.Map).OnlyGroups And MapInfo(.pos.Map).Salida.Map <> 0 Then
            Call WriteLocaleMsg(UserIndex, MSG_DEBES_ESTAR_GRUPO_PERMANECER_MAPA_2060, e_FontTypeNames.FONTTYPE_INFO)
            Call WarpUserChar(UserIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
        End If
    End With
    Call RefreshCharStatus(UserIndex)
    Exit Sub
SalirDeGrupo_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.SalirDeGrupo", Erl)
End Sub

Public Sub SalirDeGrupoForzado(ByVal UserIndex As Integer)
    On Error GoTo SalirDeGrupoForzado_Err
    Dim i          As Long
    Dim LoopC      As Long
    Dim indexviejo As Byte
    Dim GroupLider As Integer
    With UserList(UserIndex)
        .Grupo.EnGrupo = False
        .Grupo.Id = -1
        For i = 1 To 6
            If .name = UserList(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex).name Then
                Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i), 0)
                indexviejo = i
                For LoopC = indexviejo To 5
                    UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(LoopC) = UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(LoopC + 1)
                Next LoopC
                Exit For
            End If
        Next i
        UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros = UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros - 1
        Call modSendData.SendData(ToGroup, .Grupo.Lider.ArrayIndex, PrepareUpdateGroupInfo(.Grupo.Lider.ArrayIndex))
        Call modSendData.SendData(ToIndex, UserIndex, PrepareUpdateGroupInfo(UserIndex))
        Dim a As Long
        For a = 1 To UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros
            Call WriteUbicacion(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(a).ArrayIndex, indexviejo, 0)
        Next a
        Call WriteLocaleMsg(.Grupo.Lider.ArrayIndex, "202", e_FontTypeNames.FONTTYPE_New_GRUPO, .name)
        If UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros = 1 Then
            Call WriteLocaleMsg(.Grupo.Lider.ArrayIndex, "35", e_FontTypeNames.FONTTYPE_New_GRUPO)
            Call WriteUbicacion(.Grupo.Lider.ArrayIndex, 1, 0)
            UserList(.Grupo.Lider.ArrayIndex).Grupo.EnGrupo = False
            Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Lider, 0)
            Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.PropuestaDe, 0)
            UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros = 0
            Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(1), 0)
            UserList(.Grupo.Lider.ArrayIndex).Grupo.Id = -1
            Call RefreshCharStatus(.Grupo.Lider.ArrayIndex)
            Dim LiderMap As Integer: LiderMap = UserList(.Grupo.Lider.ArrayIndex).pos.Map
            If MapInfo(LiderMap).OnlyGroups And MapInfo(LiderMap).Salida.Map <> 0 Then
                Call WriteLocaleMsg(.Grupo.Lider.ArrayIndex, MSG_DEBES_ESTAR_GRUPO_PERMANECER_MAPA_2061, e_FontTypeNames.FONTTYPE_INFO) ' Msg2061="Debes estar en un grupo para permanecer en este mapa."
                Call WarpUserChar(.Grupo.Lider.ArrayIndex, MapInfo(LiderMap).Salida.Map, MapInfo(LiderMap).Salida.x, MapInfo(LiderMap).Salida.y, True)
            End If
        End If
        If MapInfo(.pos.Map).OnlyGroups And MapInfo(.pos.Map).Salida.Map <> 0 Then
            Call WriteLocaleMsg(UserIndex, MSG_DEBES_ESTAR_GRUPO_PERMANECER_MAPA_2062, e_FontTypeNames.FONTTYPE_INFO) ' Msg2062="Debes estar en un grupo para permanecer en este mapa."
            Call WarpUserChar(UserIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
        End If
    End With
    Exit Sub
SalirDeGrupoForzado_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.SalirDeGrupoForzado", Erl)
End Sub
Public Sub FinalizarGrupo(ByVal LiderIndex As Integer)
    On Error GoTo FinalizarGrupo_Err
    Dim i As Integer
    Dim OriginalMemberCount As Integer
    
    ' Guardar la cantidad original de miembros antes de modificar nada
    OriginalMemberCount = UserList(LiderIndex).Grupo.CantidadMiembros
    
    ' Procesar a todos los miembros del grupo
    For i = 1 To OriginalMemberCount
        Dim MemberIndex As Integer: MemberIndex = UserList(LiderIndex).Grupo.Miembros(i).ArrayIndex
        With UserList(MemberIndex)
            Dim j As Integer
            For j = 1 To OriginalMemberCount
                Call WriteUbicacion(MemberIndex, j, 0)
            Next j
            
            ' Enviar actualización del grupo a cada miembro
            Call modSendData.SendData(ToIndex, MemberIndex, PrepareUpdateGroupInfo(MemberIndex))
            
            ' Limpiar datos del grupo para este miembro
            .Grupo.EnGrupo = False
            .Grupo.Id = -1
            Call SetUserRef(.Grupo.Lider, 0)
            Call SetUserRef(.Grupo.PropuestaDe, 0)
            
            If MemberIndex = LiderIndex Then
                Call WriteLocaleMsg(LiderIndex, MSG_DISUELTO_GRUPO, e_FontTypeNames.FONTTYPE_INFOIAO) ' Msg2063="Has disuelto el grupo."
            Else
                Call WriteLocaleMsg(MemberIndex, MSG_LIDER_HA_ABANDONADO_GRUPO_GRUPO_DISUELVE, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2064="El líder ha abandonado el grupo. El grupo se disuelve."
            End If
            
            If MapInfo(.pos.Map).OnlyGroups And MapInfo(.pos.Map).Salida.Map <> 0 Then
                Call WriteLocaleMsg(MemberIndex, MSG_DEBES_ESTAR_GRUPO_PERMANECER_MAPA_2065, e_FontTypeNames.FONTTYPE_INFO) ' Msg2065="Debes estar en un grupo para permanecer en este mapa."
                Call WarpUserChar(MemberIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
            Else
                Call RefreshCharStatus(MemberIndex)
            End If
        End With
    Next i
    
    ' Limpiar los datos del grupo del líder
    UserList(LiderIndex).Grupo.EnGrupo = False
    UserList(LiderIndex).Grupo.Id = -1
    Call SetUserRef(UserList(LiderIndex).Grupo.Lider, 0)
    Call SetUserRef(UserList(LiderIndex).Grupo.PropuestaDe, 0)
    UserList(LiderIndex).Grupo.CantidadMiembros = 0
    
    ' Enviar actualización final al líder (ahora recibirá GroupSize = 0)
    Call modSendData.SendData(ToIndex, LiderIndex, PrepareUpdateGroupInfo(LiderIndex))
    
    Exit Sub
FinalizarGrupo_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.FinalizarGrupo", Erl)
End Sub

Public Sub CompartirUbicacion(ByVal UserIndex As Integer)
    On Error GoTo CompartirUbicacion_Err
    Dim i       As Byte
    Dim a       As Byte
    Dim indexpj As Byte
    Dim Lider   As t_User
    With UserList(UserIndex)
        Lider = UserList(.Grupo.Lider.ArrayIndex)
        For a = 1 To Lider.Grupo.CantidadMiembros
            If Lider.Grupo.Miembros(a).ArrayIndex = UserIndex Then
                indexpj = a
            End If
        Next a
        For i = 1 To Lider.Grupo.CantidadMiembros
            If Lider.Grupo.Miembros(i).ArrayIndex <> UserIndex Then
                If UserList(Lider.Grupo.Miembros(i).ArrayIndex).pos.Map = .pos.Map Then
                    Call WriteUbicacion(Lider.Grupo.Miembros(i).ArrayIndex, indexpj, UserIndex)
                    'Si va al mapa del compañero
                    Call WriteUbicacion(UserIndex, i, Lider.Grupo.Miembros(i).ArrayIndex)
                Else
                    ' Le borro la ubicacion a ellos
                    Call WriteUbicacion(Lider.Grupo.Miembros(i).ArrayIndex, indexpj, 0)
                    ' Le borro la ubicacion a mi
                    Call WriteUbicacion(UserIndex, i, 0)
                End If
            End If
        Next i
    End With
    Exit Sub
CompartirUbicacion_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.CompartirUbicacion", Erl)
End Sub

Public Sub GroupCreateSuccess(ByVal LiderIndex As Integer)
    Call WriteLocaleMsg(LiderIndex, "36", e_FontTypeNames.FONTTYPE_INFOIAO)
    With UserList(LiderIndex)
        .Grupo.EnGrupo = True
        .Grupo.Id = GetNextId()
    End With
End Sub

Public Sub AddUserToGRoup(ByVal UserIndex As Integer, ByVal GroupLiderIndex As Integer)
    On Error GoTo AddUserToGRoup_Err
    Dim Index As Byte
    For Index = 1 To UserList(GroupLiderIndex).Grupo.CantidadMiembros
        If UserList(GroupLiderIndex).Grupo.Miembros(Index).ArrayIndex = UserIndex Then
            Exit Sub
        End If
    Next Index
    UserList(GroupLiderIndex).Grupo.CantidadMiembros = UserList(GroupLiderIndex).Grupo.CantidadMiembros + 1
    Call SetUserRef(UserList(GroupLiderIndex).Grupo.Miembros(UserList(GroupLiderIndex).Grupo.CantidadMiembros), UserIndex)
    Call SetUserRef(UserList(UserIndex).Grupo.Lider, GroupLiderIndex)
    UserList(UserIndex).Grupo.EnGrupo = True
    UserList(UserIndex).Grupo.Id = UserList(GroupLiderIndex).Grupo.Id
    For Index = 2 To UserList(GroupLiderIndex).Grupo.CantidadMiembros - 1
        Call WriteLocaleMsg(UserList(GroupLiderIndex).Grupo.Miembros(Index).ArrayIndex, "40", e_FontTypeNames.FONTTYPE_INFOIAO, UserList(UserIndex).name)
    Next Index
    Call WriteLocaleMsg(UserList(UserIndex).Grupo.PropuestaDe.ArrayIndex, "40", e_FontTypeNames.FONTTYPE_INFOIAO, UserList(UserIndex).name)
    Call WriteLocaleMsg(UserIndex, MSG_SIDO_ANADIDO_GRUPO, e_FontTypeNames.FONTTYPE_INFOIAO) ' Msg2066="¡Has sido añadido al grupo!"
    Call RefreshCharStatus(GroupLiderIndex)
    Call RefreshCharStatus(UserIndex)
    Call CompartirUbicacion(UserIndex)
    Call modSendData.SendData(ToGroup, UserIndex, PrepareUpdateGroupInfo(UserIndex))
    Exit Sub
AddUserToGRoup_Err:
    Call TraceError(Err.Number, Err.Description, "ModGrupos.AddUserToGRoup", Erl)
End Sub
