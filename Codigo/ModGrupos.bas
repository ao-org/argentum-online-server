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
            
    Dim Remitente        As user: Remitente = UserList(UserIndex)
    Dim Invitado         As user: Invitado = UserList(InvitadoIndex)

    ' Fundar un party require 15 puntos de liderazgo, pero el carisma ayuda
    skillsNecesarios = 15 - Remitente.Stats.UserAtributos(eAtributos.Carisma) \ 2
    
    If Remitente.Stats.UserSkills(eSkill.liderazgo) < skillsNecesarios Then
        Call WriteConsoleMsg(UserIndex, "Te faltan " & (skillsNecesarios - Remitente.Stats.UserSkills(eSkill.liderazgo)) & " puntos en Liderazgo para liderar un grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
        Exit Sub

    End If

    If Invitado.flags.SeguroParty Then
        Call WriteConsoleMsg(UserIndex, "El usuario debe desactivar el seguro de grupos para poder invitarlo.", FontTypeNames.FONTTYPE_New_GRUPO)
        Exit Sub
    End If
        
    If Remitente.Grupo.CantidadMiembros >= UBound(Remitente.Grupo.Miembros) Then
        Call WriteConsoleMsg(UserIndex, "No puedes invitar a mas personas. (Límite: " & CStr(UBound(Remitente.Grupo.Miembros)) & ")", FontTypeNames.FONTTYPE_New_GRUPO)
        Exit Sub
    End If
            
    If Status(UserIndex) <> Status(InvitadoIndex) Or _
        (Status(UserIndex) = 1 And Status(InvitadoIndex) = 3) Or _
        (Status(UserIndex) = 3 And Status(InvitadoIndex) = 1) Or _
        (Status(UserIndex) = 0 And Status(InvitadoIndex) = 2) Or _
        (Status(UserIndex) = 2 And Status(InvitadoIndex) = 0) Then
        
        Call WriteConsoleMsg(UserIndex, "No podes crear un grupo con personajes de diferentes facciones.", FontTypeNames.FONTTYPE_New_GRUPO)
        Exit Sub
            
    End If

    If Abs(CInt(Invitado.Stats.ELV) - CInt(Remitente.Stats.ELV)) > 10 Then
        Call WriteConsoleMsg(UserIndex, "No podes crear un grupo con personajes con diferencia de más de 10 niveles.", FontTypeNames.FONTTYPE_New_GRUPO)
        Exit Sub
            
    End If

    If Invitado.Grupo.EnGrupo Then
        'Call WriteConsoleMsg(userindex, "El usuario ya se encuentra en un grupo.", FontTypeNames.FONTTYPE_INFOIAO)
        Call WriteLocaleMsg(UserIndex, "41", FontTypeNames.FONTTYPE_New_GRUPO)
        Exit Sub
            
    End If
        
    Call WriteLocaleMsg(UserIndex, "42", FontTypeNames.FONTTYPE_New_GRUPO)
    'Call WriteConsoleMsg(userindex, "Se envio la invitacion a " & UserList(Invitado).name & ", ahora solo resta aguardar su respuesta.", FontTypeNames.FONTTYPE_INFOIAO)
    Call WriteConsoleMsg(InvitadoIndex, Remitente.name & " te invitó a unirse a su grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
                
    With UserList(InvitadoIndex)
                
        .Grupo.PropuestaDe = UserIndex
        .flags.pregunta = 1
        .Grupo.Lider = UserIndex
                
    End With

    Call WritePreguntaBox(InvitadoIndex, Remitente.name & " te invito a unirse a su grupo. ¿Deseas unirte?")

    Exit Sub

InvitarMiembro_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGrupos.InvitarMiembro", Erl)
    Resume Next
        
End Sub

Public Sub EcharMiembro(ByVal UserIndex As Integer, ByVal Indice As Byte)
        
    On Error GoTo EcharMiembro_Err

    Dim i              As Long ' Iterar con long es MAS RAPIDO que otro tipo
    Dim LoopC          As Long
    Dim indexviejo     As Byte
    Dim UserIndexEchar As Integer
    
    With UserList(UserIndex).Grupo
    
        If Not .EnGrupo Then
            Call WriteConsoleMsg(UserIndex, "No estas en ningun grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        
        End If
    
        If .Lider = UserIndex Then
            Call WriteConsoleMsg(UserIndex, "Tu no podés hechar usuarios del grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
            
        End If
        
        UserIndexEchar = UserList(.Lider).Grupo.Miembros(Indice + 1)

        If UserIndexEchar <> UserIndex Then
            Call WriteConsoleMsg(UserIndex, "No podés expulsarte a ti mismo.", FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        End If

        For i = 1 To 6

            If UserIndexEchar = .Miembros(i) Then
                .Miembros(i) = 0
                indexviejo = i

                For LoopC = indexviejo To 5
                    .Miembros(LoopC) = .Miembros(LoopC + 1)
                Next LoopC

                Exit For

            End If

        Next i
            
        .CantidadMiembros = .CantidadMiembros - 1
                    
        Dim a As Long

        For a = 1 To .CantidadMiembros
            Call WriteUbicacion(.Miembros(a), indexviejo, 0)
        Next a
    
    End With
    
    With UserList(UserIndexEchar)
    
        Call WriteConsoleMsg(UserIndex, .name & " fue expulsado del grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
        'Call WriteConsoleMsg(UserIndexEchar, "Fuiste eliminado del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
        Call WriteLocaleMsg(UserIndexEchar, "37", FontTypeNames.FONTTYPE_New_GRUPO)
        
        .Grupo.EnGrupo = False
        .Grupo.Lider = 0
        .Grupo.PropuestaDe = 0
        .Grupo.CantidadMiembros = 0
        .Grupo.Miembros(1) = 0
        
        Call RefreshCharStatus(UserIndexEchar)
        
    End With
    
                            
    With UserList(UserIndex).Grupo
    
        If .CantidadMiembros = 1 Then
        
            ' Call WriteConsoleMsg(userindex, "El grupo se quedo sin miembros, grupo finalizado.", FontTypeNames.FONTTYPE_INFOIAO)
            Call WriteLocaleMsg(UserIndex, "35", FontTypeNames.FONTTYPE_New_GRUPO)
            
            .EnGrupo = False
            .Lider = 0
            .PropuestaDe = 0
            .CantidadMiembros = 0
            .Miembros(1) = 0
    
        End If
    
    End With

    Call RefreshCharStatus(UserIndex)
    
    Exit Sub

EcharMiembro_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGrupos.EcharMiembro", Erl)
    Resume Next
        
End Sub

Public Sub SalirDeGrupo(ByVal UserIndex As Integer)
        
    On Error GoTo SalirDeGrupo_Err

    Dim i          As Long
    Dim LoopC      As Long
    Dim indexviejo As Byte
    
    With UserList(UserIndex)
    
        If Not .Grupo.EnGrupo Then
            Call WriteConsoleMsg(UserIndex, "No estas en ningun grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
            Exit Sub
        
        End If
    
        .Grupo.EnGrupo = False
    
        For i = 1 To 6

            If .name = UserList(UserList(.Grupo.Lider).Grupo.Miembros(i)).name Then
                UserList(.Grupo.Lider).Grupo.Miembros(i) = 0
                indexviejo = i

                For LoopC = indexviejo To 5
                    UserList(.Grupo.Lider).Grupo.Miembros(LoopC) = UserList(.Grupo.Lider).Grupo.Miembros(LoopC + 1)
                Next LoopC

                Exit For

            End If

        Next i

        UserList(.Grupo.Lider).Grupo.CantidadMiembros = UserList(.Grupo.Lider).Grupo.CantidadMiembros - 1
        
        Dim a As Long
        For a = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros
            Call WriteUbicacion(UserList(.Grupo.Lider).Grupo.Miembros(a), indexviejo, 0)
        Next a
        
        'Call WriteConsoleMsg(userindex, "Has salido del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
        'Call WriteConsoleMsg(.Grupo.Lider, .name & " a salido del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
        Call WriteLocaleMsg(UserIndex, "37", FontTypeNames.FONTTYPE_New_GRUPO)
        Call WriteLocaleMsg(.Grupo.Lider, "202", FontTypeNames.FONTTYPE_New_GRUPO, .name)
        
        If UserList(.Grupo.Lider).Grupo.CantidadMiembros = 1 Then
        
            'Call WriteConsoleMsg(.Grupo.Lider, "El grupo se quedo sin miembros, grupo finalizado.", FontTypeNames.FONTTYPE_INFOIAO)
            Call WriteLocaleMsg(.Grupo.Lider, "35", FontTypeNames.FONTTYPE_New_GRUPO)
            
            Call WriteUbicacion(.Grupo.Lider, 1, 0)
                
            UserList(.Grupo.Lider).Grupo.EnGrupo = False
            UserList(.Grupo.Lider).Grupo.Lider = 0
            UserList(.Grupo.Lider).Grupo.PropuestaDe = 0
            UserList(.Grupo.Lider).Grupo.CantidadMiembros = 0
            UserList(.Grupo.Lider).Grupo.Miembros(1) = 0
                
            Call RefreshCharStatus(.Grupo.Lider)

        End If

        Call WriteUbicacion(UserIndex, 1, 0)
    
        .Grupo.Lider = 0
    
    End With
    
    Call RefreshCharStatus(UserIndex)
 
    Exit Sub

SalirDeGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGrupos.SalirDeGrupo", Erl)
    Resume Next
        
End Sub

Public Sub SalirDeGrupoForzado(ByVal UserIndex As Integer)
        
    On Error GoTo SalirDeGrupoForzado_Err

    Dim i          As Long
    Dim LoopC      As Long
    Dim indexviejo As Byte
    
    With UserList(UserIndex)
    
        .Grupo.EnGrupo = False
    
        For i = 1 To 6

            If .name = UserList(UserList(.Grupo.Lider).Grupo.Miembros(i)).name Then
                UserList(.Grupo.Lider).Grupo.Miembros(i) = 0
                indexviejo = i

                For LoopC = indexviejo To 5
                    UserList(.Grupo.Lider).Grupo.Miembros(LoopC) = UserList(.Grupo.Lider).Grupo.Miembros(LoopC + 1)
                Next LoopC

                Exit For

            End If

        Next i

        UserList(.Grupo.Lider).Grupo.CantidadMiembros = UserList(.Grupo.Lider).Grupo.CantidadMiembros - 1
        
        Dim a As Long
        For a = 1 To UserList(.Grupo.Lider).Grupo.CantidadMiembros
            Call WriteUbicacion(UserList(.Grupo.Lider).Grupo.Miembros(a), indexviejo, 0)
        Next a

        'Call WriteConsoleMsg(.Grupo.Lider, .name & " a salido del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
        Call WriteLocaleMsg(.Grupo.Lider, "202", FontTypeNames.FONTTYPE_New_GRUPO, .name)
        
        If UserList(.Grupo.Lider).Grupo.CantidadMiembros = 1 Then
        
            'Call WriteConsoleMsg(.Grupo.Lider, "El grupo se quedo sin miembros, grupo finalizado.", FontTypeNames.FONTTYPE_INFOIAO)
            Call WriteLocaleMsg(.Grupo.Lider, "35", FontTypeNames.FONTTYPE_New_GRUPO)
            
            Call WriteUbicacion(.Grupo.Lider, 1, 0)
                
            UserList(.Grupo.Lider).Grupo.EnGrupo = False
            UserList(.Grupo.Lider).Grupo.Lider = 0
            UserList(.Grupo.Lider).Grupo.PropuestaDe = 0
            UserList(.Grupo.Lider).Grupo.CantidadMiembros = 0
            UserList(.Grupo.Lider).Grupo.Miembros(1) = 0
                
            Call RefreshCharStatus(.Grupo.Lider)

        End If
    
    End With
        
    Exit Sub

SalirDeGrupoForzado_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGrupos.SalirDeGrupoForzado", Erl)
    Resume Next
        
End Sub

Public Sub FinalizarGrupo(ByVal UserIndex As Integer)
        
    On Error GoTo FinalizarGrupo_Err
    
    Dim i As Long
    
    With UserList(UserIndex)

        For i = 2 To .Grupo.CantidadMiembros
        
            UserList(.Grupo.Miembros(i)).Grupo.EnGrupo = False
            UserList(.Grupo.Miembros(i)).Grupo.Lider = 0
            UserList(.Grupo.Miembros(i)).Grupo.PropuestaDe = 0
            
            Call WriteUbicacion(UserList(.Grupo.Lider).Grupo.Miembros(i), i, 0)
    
            Call WriteConsoleMsg(.Grupo.Miembros(i), "El lider abandonado el grupo, grupo finalizado.", FontTypeNames.FONTTYPE_New_GRUPO)
            
            Call RefreshCharStatus(.Grupo.Miembros(i))
            
            Call WriteUbicacion(UserList(.Grupo.Lider).Grupo.Miembros(i), 1, 0)
    
            .Grupo.EnGrupo = False
            
        Next i
    
    End With
        
    Exit Sub

FinalizarGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGrupos.FinalizarGrupo", Erl)
    Resume Next
        
End Sub

Public Sub CompartirUbicacion(ByVal UserIndex As Integer)
        
    On Error GoTo CompartirUbicacion_Err

    Dim i       As Byte
    Dim a       As Byte
    Dim indexpj As Byte
    Dim Lider   As user
    
    With UserList(UserIndex)
        
        Lider = UserList(.Grupo.Lider)
        
        For a = 1 To Lider.Grupo.CantidadMiembros

            If Lider.Grupo.Miembros(a) = UserIndex Then
                indexpj = a
            End If

        Next a

        For i = 1 To Lider.Grupo.CantidadMiembros

            If Lider.Grupo.Miembros(i) <> UserIndex Then
            
                If UserList(Lider.Grupo.Miembros(i)).Pos.Map = .Pos.Map Then
                
                    Call WriteUbicacion(Lider.Grupo.Miembros(i), indexpj, UserIndex)

                    'Si va al mapa del compañero
                    Call WriteUbicacion(UserIndex, i, Lider.Grupo.Miembros(i))
                    
                Else
                    
                    ' Le borro la ubicacion a ellos
                    Call WriteUbicacion(Lider.Grupo.Miembros(i), indexpj, 0)
                    
                    ' Le borro la ubicacion a mi
                    Call WriteUbicacion(UserIndex, i, 0)
                    
                End If
                
            End If
            
        Next i
    
    End With
        
    Exit Sub

CompartirUbicacion_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModGrupos.CompartirUbicacion", Erl)
    Resume Next
        
End Sub

