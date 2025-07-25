Attribute VB_Name = "ModGrupos"
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

Public Grupo As Tgrupo

Private UniqueIdCounter As Long
Private Function GetNextId() As Long
    UniqueIdCounter = (UniqueIdCounter + 1) And &H7FFFFFFF
    GetNextId = UniqueIdCounter
End Function

Public Sub InvitarMiembro(ByVal UserIndex As Integer, ByVal InvitadoIndex As Integer)
        On Error GoTo InvitarMiembro_Err

        Dim skillsNecesarios As Integer
            
100     Dim Remitente        As t_User: Remitente = UserList(UserIndex)
102     Dim Invitado         As t_User: Invitado = UserList(InvitadoIndex)

104     ' Comentando linea importante abajo, solo temporalmente hasta que el sistema de grupo nuevo este implementado 3/5/2025 - ako
        ' skillsNecesarios = 15 - Remitente.Stats.UserAtributos(e_Atributos.Carisma) \ 2
          skillsNecesarios = 0
    
106     If Remitente.Stats.UserSkills(e_Skill.liderazgo) < skillsNecesarios Then
108         Call WriteLocaleMsg(UserIndex, 2041, (skillsNecesarios - Remitente.Stats.UserSkills(e_Skill.liderazgo)), e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2041="Te faltan ¬1 puntos en Liderazgo para liderar un grupo."
            Exit Sub
        End If
        
        'HarThaoS: Si invita a un gm no lo dejo
110     If EsGM(InvitadoIndex) Then
112         Call WriteLocaleMsg(UserIndex, 2042, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2042="No puedes invitar a un grupo a un GM."
            Exit Sub
        End If
        
        'Si es gm tampoco lo dejo
114     If EsGM(UserIndex) Then
116         Call WriteLocaleMsg(UserIndex, 2043, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2043="Los GMs no pueden formar parte de un grupo."
            Exit Sub
        End If

118     If Invitado.flags.SeguroParty Then
120         Call WriteLocaleMsg(UserIndex, 2044, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2044="El usuario debe desactivar el seguro de grupos para poder invitarlo."
            Exit Sub
        End If
        
122     If Remitente.Grupo.CantidadMiembros >= UBound(Remitente.Grupo.Miembros) Then
124         Call WriteLocaleMsg(UserIndex, 2045, CStr(UBound(Remitente.Grupo.Miembros)), e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2045="No puedes invitar a mas personas. (Límite: ¬1)"
            Exit Sub
        End If
            
126     If _
            (Status(userindex) = 0 And Status(InvitadoIndex) = 1) Or _
            (Status(userindex) = 1 And Status(InvitadoIndex) = 0) Or _
            (Status(userindex) = 1 And Status(InvitadoIndex) = 2) Or _
            (Status(userindex) = 2 And Status(InvitadoIndex) = 1) Or _
            (Status(userindex) = 3 And Status(InvitadoIndex) = 0) Or _
            (Status(userindex) = 0 And Status(InvitadoIndex) = 3) Or _
            (Status(userindex) = 3 And Status(InvitadoIndex) = 2) Or _
            (Status(userindex) = 2 And Status(InvitadoIndex) = 3) _
            Then
        
128         Call WriteLocaleMsg(UserIndex, 2046, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2046="No podes crear un grupo con personajes de diferentes facciones."
            Exit Sub
            
        End If

        If(CInt(Remitente.Stats.UserSkills(e_Skill.liderazgo)) >= (15 - Remitente.Stats.UserAtributos(e_Atributos.Carisma) / 2)) Then
            ' Si el lider tiene liderazgo asignado segun su raza, se permite una diferencia de 1 nivel mas
            If Abs(CInt(Invitado.Stats.ELV) - CInt(Remitente.Stats.ELV)) > (SvrConfig.GetValue("PartyELVwLeadership")) Then
                'Msg1438=No podes crear un grupo con personajes con diferencia de más de ¬1 niveles.
                Call WriteLocaleMsg(UserIndex, "1438", e_FontTypeNames.FONTTYPE_New_GRUPO, SvrConfig.GetValue("PartyELVwLeadership"))
                Exit Sub
            End If

            Else

            If Abs(CInt(Invitado.Stats.ELV) - CInt(Remitente.Stats.ELV)) > SvrConfig.GetValue("PartyELV") Then
                'Msg1438=No podes crear un grupo con personajes con diferencia de más de ¬1 niveles.
                Call WriteLocaleMsg(UserIndex, "1438", e_FontTypeNames.FONTTYPE_New_GRUPO, SvrConfig.GetValue("PartyELV"))
                Exit Sub
            End If

        End If


134     If Invitado.Grupo.EnGrupo Then
            
136         Call WriteLocaleMsg(UserIndex, "41", e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
            
        End If
        
        If UserList(InvitadoIndex).flags.RespondiendoPregunta = False Then
138         Call WriteLocaleMsg(UserIndex, "42", e_FontTypeNames.FONTTYPE_New_GRUPO)
140         Call WriteLocaleMsg(InvitadoIndex, 2049, e_FontTypeNames.FONTTYPE_New_GRUPO, Remitente.Name) ' Msg2049="¬1 te invitó a unirse a su grupo."
                    
142         With UserList(InvitadoIndex)
                    
144             Call SetUserRef(.Grupo.PropuestaDe, userIndex)
146             .flags.pregunta = 1
148             Call SetUserRef(.Grupo.Lider, userIndex)
            End With
150         Call WritePreguntaBox(InvitadoIndex, 1595, Remitente.name) 'Msg1595= ¬1 te invito a unirse a su grupo. ¿Deseas unirte?
            UserList(InvitadoIndex).flags.RespondiendoPregunta = True
        Else
            Call WriteLocaleMsg(UserIndex, 2050, e_FontTypeNames.FONTTYPE_INFO) ' Msg2050="El usuario tiene una solicitud pendiente."
        End If
        Exit Sub

InvitarMiembro_Err:
152     Call TraceError(Err.Number, Err.Description, "ModGrupos.InvitarMiembro", Erl)

        
End Sub

Public Sub EcharMiembro(ByVal UserIndex As Integer, ByVal Indice As Byte)
        
        On Error GoTo EcharMiembro_Err

        Dim i              As Long ' Iterar con long es MAS RAPIDO que otro tipo
        Dim LoopC          As Long
        Dim indexviejo     As Byte
        Dim UserIndexEchar As Integer
        Dim GroupLider     As Integer
    
100     With UserList(UserIndex).Grupo
            GroupLider = .Lider.ArrayIndex
102         If Not .EnGrupo Then
104             Call WriteLocaleMsg(UserIndex, 2051, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2051="No estás en ningun grupo"
                Exit Sub
        
            End If
    
106         If .Lider.ArrayIndex <> userIndex Then
108             Call WriteLocaleMsg(UserIndex, 2052, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2052="No podés echar a usuarios del grupo"
                Exit Sub
            End If
        
110         UserIndexEchar = UserList(.Lider.ArrayIndex).Grupo.Miembros(Indice + 1).ArrayIndex

112         If UserIndexEchar = userIndex Then
114             Call WriteLocaleMsg(UserIndex, 2053, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2053="No podés expulsarte a ti mismo."
                Exit Sub
            End If

116         For i = 1 To UBound(.Miembros)

118             If UserIndexEchar = .Miembros(i).ArrayIndex Then
120                 Call ClearUserRef(.Miembros(i))
122                 indexviejo = i

124                 For LoopC = indexviejo To 5
126                     .Miembros(LoopC) = .Miembros(LoopC + 1)
128                 Next LoopC
                    i = UBound(.Miembros)
                    Call ClearUserRef(.Miembros(i))
                    Exit For
                End If
130         Next i
            
132         .CantidadMiembros = .CantidadMiembros - 1
                    
            Dim a As Long

134         For a = 1 To .CantidadMiembros
136             Call WriteUbicacion(.Miembros(a).ArrayIndex, indexviejo, 0)
138         Next a
    
        End With
    
140     With UserList(UserIndexEchar)
    
142         Call WriteLocaleMsg(UserIndex, 2054, .Name, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2054="¬1 fue expulsado del grupo."
144         Call WriteLocaleMsg(UserIndexEchar, "37", e_FontTypeNames.FONTTYPE_New_GRUPO)
        
146         .Grupo.EnGrupo = False
148         Call SetUserRef(.Grupo.Lider, 0)
150         Call SetUserRef(.Grupo.PropuestaDe, 0)
152         .Grupo.CantidadMiembros = 0
154         Call SetUserRef(.Grupo.Miembros(1), 0)
156         Call RefreshCharStatus(UserIndexEchar)
            .Grupo.ID = -1
            
157         If MapInfo(.pos.Map).OnlyGroups And MapInfo(.pos.Map).Salida.Map <> 0 Then
                Call WriteLocaleMsg(UserIndexEchar, 2055, e_FontTypeNames.FONTTYPE_INFO) ' Msg2055="Debes estar en un grupo para permanecer en este mapa."
                Call WarpUserChar(UserIndexEchar, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
            End If
            
        End With
    
                            
158     With UserList(UserIndex).Grupo
160         If .CantidadMiembros = 1 Then
162             Call WriteLocaleMsg(UserIndex, "35", e_FontTypeNames.FONTTYPE_New_GRUPO)
164             .EnGrupo = False
166             Call SetUserRef(.Lider, 0)
168             Call SetUserRef(.PropuestaDe, 0)
170             .CantidadMiembros = 0
172             Call SetUserRef(.Miembros(1), 0)
                .ID = -1
                Call modSendData.SendData(ToIndex, UserIndex, PrepareUpdateGroupInfo(UserIndex))
                
173             Dim LiderMap As Integer: LiderMap = UserList(UserIndex).pos.Map
                If MapInfo(LiderMap).OnlyGroups And MapInfo(LiderMap).Salida.Map <> 0 Then
                    Call WriteLocaleMsg(UserIndex, 2056, e_FontTypeNames.FONTTYPE_INFO) ' Msg2056="Debes estar en un grupo para permanecer en este mapa."
                    Call WarpUserChar(UserIndex, MapInfo(LiderMap).Salida.Map, MapInfo(LiderMap).Salida.x, MapInfo(LiderMap).Salida.y, True)
                End If
            End If
        End With
174     Call RefreshCharStatus(UserIndex)
        Call modSendData.SendData(ToGroup, GroupLider, PrepareUpdateGroupInfo(GroupLider))
        Call modSendData.SendData(ToIndex, UserIndexEchar, PrepareUpdateGroupInfo(UserIndexEchar))
        Exit Sub

EcharMiembro_Err:
176     Call TraceError(Err.Number, Err.Description, "ModGrupos.EcharMiembro", Erl)

        
End Sub

Public Sub SalirDeGrupo(ByVal UserIndex As Integer)
        
        On Error GoTo SalirDeGrupo_Err

        Dim i          As Long
        Dim LoopC      As Long
        Dim indexviejo As Byte
    
100     With UserList(UserIndex)
    
102         If Not .Grupo.EnGrupo Then
104             Call WriteLocaleMsg(UserIndex, 2057, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2057="No estas en ningun grupo."
                Exit Sub
        
            End If
    
106         .Grupo.EnGrupo = False
            .Grupo.ID = -1
108         For i = 1 To UBound(.Grupo.Miembros)

110             If .name = UserList(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex).name Then
112                 Call ClearUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i))
114                 indexviejo = i

116                 For LoopC = indexviejo To 5
118                     UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(LoopC) = UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(LoopC + 1)
120                 Next LoopC
                    i = UBound(.Grupo.Miembros)
                    Call ClearUserRef(.Grupo.Miembros(i))
                    Exit For
                End If
122         Next i
124         UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros = UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros - 1
            Dim a As Long
126         For a = 1 To UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros
128             Call WriteUbicacion(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(a).ArrayIndex, indexviejo, 0)
130         Next a
        
132         Call WriteLocaleMsg(UserIndex, "37", e_FontTypeNames.FONTTYPE_New_GRUPO) 'quit group message
134         Call WriteLocaleMsg(.Grupo.Lider.ArrayIndex, "202", e_FontTypeNames.FONTTYPE_New_GRUPO, .name)
        
136         If UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros = 1 Then
        
138             Call WriteLocaleMsg(.Grupo.Lider.ArrayIndex, "35", e_FontTypeNames.FONTTYPE_New_GRUPO)
            
140             Call WriteUbicacion(.Grupo.Lider.ArrayIndex, 1, 0)
                UserList(.Grupo.Lider.ArrayIndex).Grupo.ID = -1
142             UserList(.Grupo.Lider.ArrayIndex).Grupo.EnGrupo = False
144             Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Lider, 0)
146             Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.PropuestaDe, 0)
148             UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros = 0
150             Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(1), 0)
152             Call RefreshCharStatus(.Grupo.Lider.ArrayIndex)
                Call modSendData.SendData(ToIndex, .Grupo.Lider.ArrayIndex, PrepareUpdateGroupInfo(.Grupo.Lider.ArrayIndex))
                
153             Dim LiderMap As Integer: LiderMap = UserList(.Grupo.Lider.ArrayIndex).pos.Map
                If MapInfo(LiderMap).OnlyGroups And MapInfo(LiderMap).Salida.Map <> 0 Then
                    Call WriteLocaleMsg(.Grupo.Lider.ArrayIndex, 2059, e_FontTypeNames.FONTTYPE_INFO) ' Msg2059="Debes estar en un grupo para permanecer en este mapa."
                    Call WarpUserChar(.Grupo.Lider.ArrayIndex, MapInfo(LiderMap).Salida.Map, MapInfo(LiderMap).Salida.x, MapInfo(LiderMap).Salida.y, True)
                End If
            End If

154         Call WriteUbicacion(UserIndex, 1, 0)
            Call modSendData.SendData(ToGroup, .Grupo.Lider.ArrayIndex, PrepareUpdateGroupInfo(.Grupo.Lider.ArrayIndex))
156         Call SetUserRef(.Grupo.Lider, 0)
            
            Call modSendData.SendData(ToIndex, UserIndex, PrepareUpdateGroupInfo(UserIndex))
            
157         If MapInfo(.pos.Map).OnlyGroups And MapInfo(.pos.Map).Salida.Map <> 0 Then
                Call WriteLocaleMsg(UserIndex, 2060, e_FontTypeNames.FONTTYPE_INFO) ' Msg2060="Debes estar en un grupo para permanecer en este mapa."
                Call WarpUserChar(UserIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
            End If
            
        End With
    
158     Call RefreshCharStatus(UserIndex)
 
        Exit Sub

SalirDeGrupo_Err:
160     Call TraceError(Err.Number, Err.Description, "ModGrupos.SalirDeGrupo", Erl)

        
End Sub

Public Sub SalirDeGrupoForzado(ByVal UserIndex As Integer)
        
        On Error GoTo SalirDeGrupoForzado_Err

        Dim i          As Long
        Dim LoopC      As Long
        Dim indexviejo As Byte
        Dim GroupLider As Integer
    
100     With UserList(UserIndex)
102         .Grupo.EnGrupo = False
            .Grupo.ID = -1
104         For i = 1 To 6
106             If .name = UserList(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex).name Then
108                 Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i), 0)
110                 indexviejo = i

112                 For LoopC = indexviejo To 5
114                     UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(LoopC) = UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(LoopC + 1)
116                 Next LoopC
                    Exit For
                End If
118         Next i
120         UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros = UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros - 1
            Call modSendData.SendData(ToGroup, .Grupo.Lider.ArrayIndex, PrepareUpdateGroupInfo(.Grupo.Lider.ArrayIndex))
            Call modSendData.SendData(ToIndex, UserIndex, PrepareUpdateGroupInfo(UserIndex))
            Dim a As Long
122         For a = 1 To UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros
124             Call WriteUbicacion(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(a).ArrayIndex, indexviejo, 0)
126         Next a

128         Call WriteLocaleMsg(.Grupo.Lider.ArrayIndex, "202", e_FontTypeNames.FONTTYPE_New_GRUPO, .name)
        
130         If UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros = 1 Then
132             Call WriteLocaleMsg(.Grupo.Lider.ArrayIndex, "35", e_FontTypeNames.FONTTYPE_New_GRUPO)
134             Call WriteUbicacion(.Grupo.Lider.ArrayIndex, 1, 0)
136             UserList(.Grupo.Lider.ArrayIndex).Grupo.EnGrupo = False
138             Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Lider, 0)
140             Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.PropuestaDe, 0)
142             UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros = 0
144             Call SetUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(1), 0)
                UserList(.Grupo.Lider.ArrayIndex).Grupo.ID = -1
146             Call RefreshCharStatus(.Grupo.Lider.ArrayIndex)

147             Dim LiderMap As Integer: LiderMap = UserList(.Grupo.Lider.ArrayIndex).pos.Map
                If MapInfo(LiderMap).OnlyGroups And MapInfo(LiderMap).Salida.Map <> 0 Then
                    Call WriteLocaleMsg(.Grupo.Lider.ArrayIndex, 2061, e_FontTypeNames.FONTTYPE_INFO) ' Msg2061="Debes estar en un grupo para permanecer en este mapa."
                    Call WarpUserChar(.Grupo.Lider.ArrayIndex, MapInfo(LiderMap).Salida.Map, MapInfo(LiderMap).Salida.x, MapInfo(LiderMap).Salida.y, True)
                End If
            End If
            
            If MapInfo(.pos.Map).OnlyGroups And MapInfo(.pos.Map).Salida.Map <> 0 Then
                Call WriteLocaleMsg(UserIndex, 2062, e_FontTypeNames.FONTTYPE_INFO) ' Msg2062="Debes estar en un grupo para permanecer en este mapa."
                Call WarpUserChar(UserIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
            End If
            
        End With
        Exit Sub
SalirDeGrupoForzado_Err:
148     Call TraceError(Err.Number, Err.Description, "ModGrupos.SalirDeGrupoForzado", Erl)

        
End Sub

Public Sub FinalizarGrupo(ByVal LiderIndex As Integer)
On Error GoTo FinalizarGrupo_Err

        Dim i As Integer
        For i = 1 To UserList(LiderIndex).Grupo.CantidadMiembros
            Dim MemberIndex As Integer: MemberIndex = UserList(LiderIndex).Grupo.Miembros(i).ArrayIndex

            With UserList(MemberIndex)

                Dim j As Integer
                For j = 1 To UserList(LiderIndex).Grupo.CantidadMiembros
                    Call WriteUbicacion(MemberIndex, j, 0)
                Next j

                Call modSendData.SendData(ToIndex, MemberIndex, PrepareUpdateGroupInfo(MemberIndex))
                
                .Grupo.EnGrupo = False
                .Grupo.ID = -1

                Call SetUserRef(.Grupo.Lider, 0)
                Call SetUserRef(.Grupo.PropuestaDe, 0)

                If MemberIndex = LiderIndex Then
                    Call WriteLocaleMsg(LiderIndex, 2063, e_FontTypeNames.FONTTYPE_INFOIAO) ' Msg2063="Has disuelto el grupo."
                Else
                    Call WriteLocaleMsg(MemberIndex, 2064, e_FontTypeNames.FONTTYPE_New_GRUPO) ' Msg2064="El líder ha abandonado el grupo. El grupo se disuelve."
                End If

                If MapInfo(.pos.Map).OnlyGroups And MapInfo(.pos.Map).Salida.Map <> 0 Then
                    Call WriteLocaleMsg(MemberIndex, 2065, e_FontTypeNames.FONTTYPE_INFO) ' Msg2065="Debes estar en un grupo para permanecer en este mapa."
                    Call WarpUserChar(MemberIndex, MapInfo(.pos.Map).Salida.Map, MapInfo(.pos.Map).Salida.x, MapInfo(.pos.Map).Salida.y, True)
                Else
                    Call RefreshCharStatus(MemberIndex)
                End If

            End With

        Next i
        
        UserList(LiderIndex).Grupo.CantidadMiembros = 0

        Exit Sub

FinalizarGrupo_Err:
122     Call TraceError(Err.Number, Err.Description, "ModGrupos.FinalizarGrupo", Erl)
End Sub

Public Sub CompartirUbicacion(ByVal UserIndex As Integer)
        
        On Error GoTo CompartirUbicacion_Err

        Dim i       As Byte
        Dim a       As Byte
        Dim indexpj As Byte
        Dim Lider   As t_User
    
100     With UserList(UserIndex)
102         Lider = UserList(.Grupo.Lider.ArrayIndex)
104         For a = 1 To Lider.Grupo.CantidadMiembros
106             If Lider.Grupo.Miembros(a).ArrayIndex = userIndex Then
108                 indexpj = a
                End If
110         Next a

112         For i = 1 To Lider.Grupo.CantidadMiembros
114             If Lider.Grupo.Miembros(i).ArrayIndex <> userIndex Then
116                 If UserList(Lider.Grupo.Miembros(i).ArrayIndex).pos.map = .pos.map Then
118                     Call WriteUbicacion(Lider.Grupo.Miembros(i).ArrayIndex, indexpj, userIndex)
                        'Si va al mapa del compañero
120                     Call WriteUbicacion(userIndex, i, Lider.Grupo.Miembros(i).ArrayIndex)
                    Else
                        ' Le borro la ubicacion a ellos
122                     Call WriteUbicacion(Lider.Grupo.Miembros(i).ArrayIndex, indexpj, 0)
                        ' Le borro la ubicacion a mi
124                     Call WriteUbicacion(UserIndex, i, 0)
                    End If
                End If
126         Next i
        End With
        Exit Sub
CompartirUbicacion_Err:
128     Call TraceError(Err.Number, Err.Description, "ModGrupos.CompartirUbicacion", Erl)
End Sub

Public Sub GroupCreateSuccess(ByVal LiderIndex As Integer)
    Call WriteLocaleMsg(LiderIndex, "36", e_FontTypeNames.FONTTYPE_INFOIAO)
    With UserList(LiderIndex)
        .Grupo.EnGrupo = True
        .Grupo.ID = GetNextId()
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
100 UserList(GroupLiderIndex).Grupo.CantidadMiembros = UserList(GroupLiderIndex).Grupo.CantidadMiembros + 1
102 Call SetUserRef(UserList(GroupLiderIndex).Grupo.Miembros(UserList(GroupLiderIndex).Grupo.CantidadMiembros), UserIndex)
103 Call SetUserRef(UserList(UserIndex).Grupo.Lider, GroupLiderIndex)
104 UserList(UserIndex).Grupo.EnGrupo = True
106 UserList(UserIndex).Grupo.ID = UserList(GroupLiderIndex).Grupo.ID
110 For Index = 2 To UserList(GroupLiderIndex).Grupo.CantidadMiembros - 1
114     Call WriteLocaleMsg(UserList(GroupLiderIndex).Grupo.Miembros(Index).ArrayIndex, "40", e_FontTypeNames.FONTTYPE_INFOIAO, UserList(UserIndex).name)
116 Next Index
    
120 Call WriteLocaleMsg(UserList(UserIndex).Grupo.PropuestaDe.ArrayIndex, "40", e_FontTypeNames.FONTTYPE_INFOIAO, UserList(UserIndex).name)
    
130 Call WriteLocaleMsg(UserIndex, 2066, e_FontTypeNames.FONTTYPE_INFOIAO) ' Msg2066="¡Has sido añadido al grupo!"
140 Call RefreshCharStatus(GroupLiderIndex)
150 Call RefreshCharStatus(UserIndex)
160 Call CompartirUbicacion(UserIndex)
    Call modSendData.SendData(ToGroup, UserIndex, PrepareUpdateGroupInfo(UserIndex))
    Exit Sub
AddUserToGRoup_Err:
122     Call TraceError(Err.Number, Err.Description, "ModGrupos.AddUserToGRoup", Erl)
End Sub

