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

        ' Fundar un party require 15 puntos de liderazgo, pero el carisma ayuda
104     skillsNecesarios = 15 - Remitente.Stats.UserAtributos(e_Atributos.Carisma) \ 2
    
106     If Remitente.Stats.UserSkills(e_Skill.liderazgo) < skillsNecesarios Then
108         Call WriteConsoleMsg(UserIndex, "Te faltan " & (skillsNecesarios - Remitente.Stats.UserSkills(e_Skill.liderazgo)) & " puntos en Liderazgo para liderar un grupo.", e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
        
        'HarThaoS: Si invita a un gm no lo dejo
110     If EsGM(InvitadoIndex) Then
112         Call WriteConsoleMsg(UserIndex, "No puedes invitar a un grupo a un GM.", e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
        
        'Si es gm tampoco lo dejo
114     If EsGM(UserIndex) Then
116         Call WriteConsoleMsg(UserIndex, "Los GMs no pueden formar parte de un grupo.", e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If

118     If Invitado.flags.SeguroParty Then
120         Call WriteConsoleMsg(UserIndex, "El usuario debe desactivar el seguro de grupos para poder invitarlo.", e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
        
122     If Remitente.Grupo.CantidadMiembros >= UBound(Remitente.Grupo.Miembros) Then
124         Call WriteConsoleMsg(UserIndex, "No puedes invitar a mas personas. (Límite: " & CStr(UBound(Remitente.Grupo.Miembros)) & ")", e_FontTypeNames.FONTTYPE_New_GRUPO)
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
        
128         Call WriteConsoleMsg(UserIndex, "No podes crear un grupo con personajes de diferentes facciones.", e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
            
        End If

130     If Abs(CInt(Invitado.Stats.ELV) - CInt(Remitente.Stats.ELV)) > 10 Then
132         Call WriteConsoleMsg(UserIndex, "No podes crear un grupo con personajes con diferencia de más de 10 niveles.", e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
            
        End If

134     If Invitado.Grupo.EnGrupo Then
            'Call WriteConsoleMsg(userindex, "El usuario ya se encuentra en un grupo.", e_FontTypeNames.FONTTYPE_INFOIAO)
136         Call WriteLocaleMsg(UserIndex, "41", e_FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
            
        End If
        
            'Call WriteConsoleMsg(userindex, "Se envio la invitacion a " & UserList(Invitado).name & ", ahora solo resta aguardar su respuesta.", e_FontTypeNames.FONTTYPE_INFOIAO)
        If UserList(InvitadoIndex).flags.RespondiendoPregunta = False Then
138         Call WriteLocaleMsg(UserIndex, "42", e_FontTypeNames.FONTTYPE_New_GRUPO)
140         Call WriteConsoleMsg(InvitadoIndex, Remitente.Name & " te invitó a unirse a su grupo.", e_FontTypeNames.FONTTYPE_New_GRUPO)
                    
142         With UserList(InvitadoIndex)
                    
144             Call SetUserRef(.Grupo.PropuestaDe, userIndex)
146             .flags.pregunta = 1
148             Call SetUserRef(.Grupo.Lider, userIndex)
            End With
150         Call WritePreguntaBox(InvitadoIndex, Remitente.Name & " te invito a unirse a su grupo. ¿Deseas unirte?")
            UserList(InvitadoIndex).flags.RespondiendoPregunta = True
        Else
            Call WriteConsoleMsg(UserIndex, "El usuario tiene una solicitud pendiente.", e_FontTypeNames.FONTTYPE_INFO)
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
104             Call WriteConsoleMsg(UserIndex, "No estás en ningun grupo", e_FontTypeNames.FONTTYPE_New_GRUPO)
                Exit Sub
        
            End If
    
106         If .Lider.ArrayIndex <> userIndex Then
108             Call WriteConsoleMsg(UserIndex, "No podés echar a usuarios del grupo", e_FontTypeNames.FONTTYPE_New_GRUPO)
                Exit Sub
            End If
        
110         UserIndexEchar = UserList(.Lider.ArrayIndex).Grupo.Miembros(Indice + 1).ArrayIndex

112         If UserIndexEchar = userIndex Then
114             Call WriteConsoleMsg(UserIndex, "No podés expulsarte a ti mismo.", e_FontTypeNames.FONTTYPE_New_GRUPO)
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
    
142         Call WriteConsoleMsg(UserIndex, .Name & " fue expulsado del grupo.", e_FontTypeNames.FONTTYPE_New_GRUPO)
144         Call WriteLocaleMsg(UserIndexEchar, "37", e_FontTypeNames.FONTTYPE_New_GRUPO)
        
146         .Grupo.EnGrupo = False
148         Call SetUserRef(.Grupo.Lider, 0)
150         Call SetUserRef(.Grupo.PropuestaDe, 0)
152         .Grupo.CantidadMiembros = 0
154         Call SetUserRef(.Grupo.Miembros(1), 0)
156         Call RefreshCharStatus(UserIndexEchar)
            .Grupo.ID = -1
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
104             Call WriteConsoleMsg(UserIndex, "No estas en ningun grupo.", e_FontTypeNames.FONTTYPE_New_GRUPO)
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
        
                'Call WriteConsoleMsg(.Grupo.Lider, "El grupo se quedo sin miembros, grupo finalizado.", e_FontTypeNames.FONTTYPE_INFOIAO)
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
            End If

154         Call WriteUbicacion(UserIndex, 1, 0)
            Call modSendData.SendData(ToGroup, .Grupo.Lider.ArrayIndex, PrepareUpdateGroupInfo(.Grupo.Lider.ArrayIndex))
156         Call SetUserRef(.Grupo.Lider, 0)
            
            Call modSendData.SendData(ToIndex, UserIndex, PrepareUpdateGroupInfo(UserIndex))
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
            End If
        End With
        Exit Sub
SalirDeGrupoForzado_Err:
148     Call TraceError(Err.Number, Err.Description, "ModGrupos.SalirDeGrupoForzado", Erl)

        
End Sub

Public Sub FinalizarGrupo(ByVal UserIndex As Integer)
On Error GoTo FinalizarGrupo_Err
        Dim i As Long
100     With UserList(UserIndex)
102         For i = 2 To .Grupo.CantidadMiembros
104             UserList(.Grupo.Miembros(i).ArrayIndex).Grupo.EnGrupo = False
106             Call SetUserRef(UserList(.Grupo.Miembros(i).ArrayIndex).Grupo.Lider, 0)
108             Call SetUserRef(UserList(.Grupo.Miembros(i).ArrayIndex).Grupo.PropuestaDe, 0)
110             Call WriteUbicacion(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex, i, 0)
112             Call WriteConsoleMsg(.Grupo.Miembros(i).ArrayIndex, "El líder ha abandonado el grupo. El grupo se disuelve.", e_FontTypeNames.FONTTYPE_New_GRUPO)
114             Call RefreshCharStatus(.Grupo.Miembros(i).ArrayIndex)
116             Call WriteUbicacion(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex, 1, 0)
118             .Grupo.EnGrupo = False
                .Grupo.ID = -1
                Call modSendData.SendData(ToIndex, .Grupo.Miembros(i).ArrayIndex, PrepareUpdateGroupInfo(.Grupo.Miembros(i).ArrayIndex))
120         Next i
        End With
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
    
130 Call WriteConsoleMsg(UserIndex, "¡Has sido añadido al grupo!", e_FontTypeNames.FONTTYPE_INFOIAO)
140 Call RefreshCharStatus(GroupLiderIndex)
150 Call RefreshCharStatus(UserIndex)
160 Call CompartirUbicacion(UserIndex)
    Call modSendData.SendData(ToGroup, UserIndex, PrepareUpdateGroupInfo(UserIndex))
    Exit Sub
AddUserToGRoup_Err:
122     Call TraceError(Err.Number, Err.Description, "ModGrupos.AddUserToGRoup", Erl)
End Sub

