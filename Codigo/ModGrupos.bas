Attribute VB_Name = "ModGrupos"

Type Tgrupo

    EnGrupo As Boolean
    CantidadMiembros As Byte
    Miembros(1 To 6) As Integer
    Lider As Integer
    PropuestaDe As Integer

End Type

Public Grupo As Tgrupo

Public Sub InvitarMiembro(ByVal UserIndex As Integer, ByVal InvitadoIndex As Integer)
        On Error GoTo InvitarMiembro_Err

        Dim skillsNecesarios As Integer
            
100     Dim Remitente        As user: Remitente = UserList(UserIndex)
102     Dim Invitado         As user: Invitado = UserList(InvitadoIndex)

        ' Fundar un party require 15 puntos de liderazgo, pero el carisma ayuda
104     skillsNecesarios = 15 - Remitente.Stats.UserAtributos(eAtributos.Carisma) \ 2
    
106     If Remitente.Stats.UserSkills(eSkill.liderazgo) < skillsNecesarios Then
108         Call WriteConsoleMsg(UserIndex, "Te faltan " & (skillsNecesarios - Remitente.Stats.UserSkills(eSkill.liderazgo)) & " puntos en Liderazgo para liderar un grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
        
        'HarThaoS: Si invita a un gm no lo dejo
110     If EsGM(InvitadoIndex) Then
112         Call WriteConsoleMsg(UserIndex, "No puedes invitar a un grupo a un GM.", FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
        
        'Si es gm tampoco lo dejo
114     If EsGM(UserIndex) Then
116         Call WriteConsoleMsg(UserIndex, "Los GMs no pueden formar parte de un grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If

118     If Invitado.flags.SeguroParty Then
120         Call WriteConsoleMsg(UserIndex, "El usuario debe desactivar el seguro de grupos para poder invitarlo.", FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
        
122     If Remitente.Grupo.CantidadMiembros >= UBound(Remitente.Grupo.Miembros) Then
124         Call WriteConsoleMsg(UserIndex, "No puedes invitar a mas personas. (Límite: " & CStr(UBound(Remitente.Grupo.Miembros)) & ")", FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If
            
126     If Status(UserIndex) <> Status(InvitadoIndex) Or _
            (Status(UserIndex) = 1 And Status(InvitadoIndex) = 3) Or _
            (Status(UserIndex) = 3 And Status(InvitadoIndex) = 1) Or _
            (Status(UserIndex) = 0 And Status(InvitadoIndex) = 2) Or _
            (Status(UserIndex) = 2 And Status(InvitadoIndex) = 0) Then
        
128         Call WriteConsoleMsg(UserIndex, "No podes crear un grupo con personajes de diferentes facciones.", FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
            
        End If

130     If Abs(CInt(Invitado.Stats.ELV) - CInt(Remitente.Stats.ELV)) > 10 Then
132         Call WriteConsoleMsg(UserIndex, "No podes crear un grupo con personajes con diferencia de más de 10 niveles.", FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
            
        End If

134     If Invitado.Grupo.EnGrupo Then
            'Call WriteConsoleMsg(userindex, "El usuario ya se encuentra en un grupo.", FontTypeNames.FONTTYPE_INFOIAO)
136         Call WriteLocaleMsg(UserIndex, "41", FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
            
        End If
        
138     Call WriteLocaleMsg(UserIndex, "42", FontTypeNames.FONTTYPE_New_GRUPO)
        'Call WriteConsoleMsg(userindex, "Se envio la invitacion a " & UserList(Invitado).name & ", ahora solo resta aguardar su respuesta.", FontTypeNames.FONTTYPE_INFOIAO)
140     Call WriteConsoleMsg(InvitadoIndex, Remitente.Name & " te invitó a unirse a su grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
                
142     With UserList(InvitadoIndex)
                
144         .Grupo.PropuestaDe = UserIndex
146         .flags.pregunta = 1
148         .Grupo.Lider = UserIndex
                
        End With

150     Call WritePreguntaBox(InvitadoIndex, Remitente.Name & " te invito a unirse a su grupo. ¿Deseas unirte?")

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
    
100     With UserList(UserIndex).Grupo
    
102         If Not .EnGrupo Then
104             Call WriteConsoleMsg(UserIndex, "No estas en ningun grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
                Exit Sub
        
            End If
    
106         If .Lider = UserIndex Then
108             Call WriteConsoleMsg(UserIndex, "Tu no podés hechar usuarios del grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
                Exit Sub
            
            End If
        
110         UserIndexEchar = UserList(.Lider).Grupo.Miembros(Indice + 1)

112         If UserIndexEchar <> UserIndex Then
114             Call WriteConsoleMsg(UserIndex, "No podés expulsarte a ti mismo.", FontTypeNames.FONTTYPE_New_GRUPO)
                Exit Sub
            End If

116         For i = 1 To 6

118             If UserIndexEchar = .Miembros(i) Then
120                 .Miembros(i) = 0
122                 indexviejo = i

124                 For LoopC = indexviejo To 5
126                     .Miembros(LoopC) = .Miembros(LoopC + 1)
128                 Next LoopC

                    Exit For

                End If

130         Next i
            
132         .CantidadMiembros = .CantidadMiembros - 1
                    
            Dim a As Long

134         For a = 1 To .CantidadMiembros
136             Call WriteUbicacion(.Miembros(a), indexviejo, 0)
138         Next a
    
        End With
    
140     With UserList(UserIndexEchar)
    
142         Call WriteConsoleMsg(UserIndex, .Name & " fue expulsado del grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
            'Call WriteConsoleMsg(UserIndexEchar, "Fuiste eliminado del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
144         Call WriteLocaleMsg(UserIndexEchar, "37", FontTypeNames.FONTTYPE_New_GRUPO)
        
146         .Grupo.EnGrupo = False
148         .Grupo.Lider = 0
150         .Grupo.PropuestaDe = 0
152         .Grupo.CantidadMiembros = 0
154         .Grupo.Miembros(1) = 0
        
156         Call RefreshCharStatus(UserIndexEchar)
        
        End With
    
                            
158     With UserList(UserIndex).Grupo
    
160         If .CantidadMiembros = 1 Then
        
                ' Call WriteConsoleMsg(userindex, "El grupo se quedo sin miembros, grupo finalizado.", FontTypeNames.FONTTYPE_INFOIAO)
162             Call WriteLocaleMsg(UserIndex, "35", FontTypeNames.FONTTYPE_New_GRUPO)
            
164             .EnGrupo = False
166             .Lider = 0
168             .PropuestaDe = 0
170             .CantidadMiembros = 0
172             .Miembros(1) = 0
    
            End If
    
        End With

174     Call RefreshCharStatus(UserIndex)
    
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
104             Call WriteConsoleMsg(UserIndex, "No estas en ningun grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
                Exit Sub
        
            End If
    
106         .Grupo.EnGrupo = False
    
108         For i = 1 To 6

110             If .Name = UserList(UserList(.Grupo.Lider).Grupo.Miembros(i)).Name Then
112                 UserList(.Grupo.Lider).Grupo.Miembros(i) = 0
114                 indexviejo = i

116                 For LoopC = indexviejo To 5
118                     UserList(.Grupo.Lider).Grupo.Miembros(LoopC) = UserList(.Grupo.Lider).Grupo.Miembros(LoopC + 1)
120                 Next LoopC

                    Exit For

                End If

122         Next i

124         UserList(.Grupo.Lider).Grupo.CantidadMiembros = UserList(.Grupo.Lider).Grupo.CantidadMiembros - 1
        
            Dim a As Long
126         For a = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros
128             Call WriteUbicacion(UserList(.Grupo.Lider).Grupo.Miembros(a), indexviejo, 0)
130         Next a
        
            'Call WriteConsoleMsg(userindex, "Has salido del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
            'Call WriteConsoleMsg(.Grupo.Lider, .name & " a salido del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
132         Call WriteLocaleMsg(UserIndex, "37", FontTypeNames.FONTTYPE_New_GRUPO)
134         Call WriteLocaleMsg(.Grupo.Lider, "202", FontTypeNames.FONTTYPE_New_GRUPO, .Name)
        
136         If UserList(.Grupo.Lider).Grupo.CantidadMiembros = 1 Then
        
                'Call WriteConsoleMsg(.Grupo.Lider, "El grupo se quedo sin miembros, grupo finalizado.", FontTypeNames.FONTTYPE_INFOIAO)
138             Call WriteLocaleMsg(.Grupo.Lider, "35", FontTypeNames.FONTTYPE_New_GRUPO)
            
140             Call WriteUbicacion(.Grupo.Lider, 1, 0)
                
142             UserList(.Grupo.Lider).Grupo.EnGrupo = False
144             UserList(.Grupo.Lider).Grupo.Lider = 0
146             UserList(.Grupo.Lider).Grupo.PropuestaDe = 0
148             UserList(.Grupo.Lider).Grupo.CantidadMiembros = 0
150             UserList(.Grupo.Lider).Grupo.Miembros(1) = 0
                
152             Call RefreshCharStatus(.Grupo.Lider)

            End If

154         Call WriteUbicacion(UserIndex, 1, 0)
    
156         .Grupo.Lider = 0
    
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
    
100     With UserList(UserIndex)
    
102         .Grupo.EnGrupo = False
    
104         For i = 1 To 6

106             If .Name = UserList(UserList(.Grupo.Lider).Grupo.Miembros(i)).Name Then
108                 UserList(.Grupo.Lider).Grupo.Miembros(i) = 0
110                 indexviejo = i

112                 For LoopC = indexviejo To 5
114                     UserList(.Grupo.Lider).Grupo.Miembros(LoopC) = UserList(.Grupo.Lider).Grupo.Miembros(LoopC + 1)
116                 Next LoopC

                    Exit For

                End If

118         Next i

120         UserList(.Grupo.Lider).Grupo.CantidadMiembros = UserList(.Grupo.Lider).Grupo.CantidadMiembros - 1
        
            Dim a As Long
122         For a = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros
124             Call WriteUbicacion(UserList(.Grupo.Lider).Grupo.Miembros(a), indexviejo, 0)
126         Next a

            'Call WriteConsoleMsg(.Grupo.Lider, .name & " a salido del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
128         Call WriteLocaleMsg(.Grupo.Lider, "202", FontTypeNames.FONTTYPE_New_GRUPO, .Name)
        
130         If UserList(.Grupo.Lider).Grupo.CantidadMiembros = 1 Then
        
                'Call WriteConsoleMsg(.Grupo.Lider, "El grupo se quedo sin miembros, grupo finalizado.", FontTypeNames.FONTTYPE_INFOIAO)
132             Call WriteLocaleMsg(.Grupo.Lider, "35", FontTypeNames.FONTTYPE_New_GRUPO)
            
134             Call WriteUbicacion(.Grupo.Lider, 1, 0)
                
136             UserList(.Grupo.Lider).Grupo.EnGrupo = False
138             UserList(.Grupo.Lider).Grupo.Lider = 0
140             UserList(.Grupo.Lider).Grupo.PropuestaDe = 0
142             UserList(.Grupo.Lider).Grupo.CantidadMiembros = 0
144             UserList(.Grupo.Lider).Grupo.Miembros(1) = 0
                
146             Call RefreshCharStatus(.Grupo.Lider)

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
        
104             UserList(.Grupo.Miembros(i)).Grupo.EnGrupo = False
106             UserList(.Grupo.Miembros(i)).Grupo.Lider = 0
108             UserList(.Grupo.Miembros(i)).Grupo.PropuestaDe = 0
            
110             Call WriteUbicacion(UserList(.Grupo.Lider).Grupo.Miembros(i), i, 0)
    
112             Call WriteConsoleMsg(.Grupo.Miembros(i), "El líder ha abandonado el grupo. El grupo se disuelve.", FontTypeNames.FONTTYPE_New_GRUPO)
            
114             Call RefreshCharStatus(.Grupo.Miembros(i))
            
116             Call WriteUbicacion(UserList(.Grupo.Lider).Grupo.Miembros(i), 1, 0)
    
118             .Grupo.EnGrupo = False
            
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
        Dim Lider   As user
    
100     With UserList(UserIndex)
        
102         Lider = UserList(.Grupo.Lider)
        
104         For a = 1 To Lider.Grupo.CantidadMiembros

106             If Lider.Grupo.Miembros(a) = UserIndex Then
108                 indexpj = a
                End If

110         Next a

112         For i = 1 To Lider.Grupo.CantidadMiembros

114             If Lider.Grupo.Miembros(i) <> UserIndex Then
            
116                 If UserList(Lider.Grupo.Miembros(i)).Pos.Map = .Pos.Map Then
                
118                     Call WriteUbicacion(Lider.Grupo.Miembros(i), indexpj, UserIndex)

                        'Si va al mapa del compañero
120                     Call WriteUbicacion(UserIndex, i, Lider.Grupo.Miembros(i))
                    
                    Else
                    
                        ' Le borro la ubicacion a ellos
122                     Call WriteUbicacion(Lider.Grupo.Miembros(i), indexpj, 0)
                    
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

