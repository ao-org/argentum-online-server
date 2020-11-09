Attribute VB_Name = "ModGrupos"

Type Tgrupo

    EnGrupo As Boolean
    CantidadMiembros As Byte
    Miembros(1 To 6) As Integer
    Lider As Integer
    PropuestaDe As Integer

End Type

Public Grupo As Tgrupo

Public Sub InvitarMiembro(ByVal UserIndex As Integer, ByVal Invitado As Integer)

    If UserList(Invitado).flags.SeguroParty = False Then
        
        If Status(UserIndex) = Status(Invitado) Or Status(UserIndex) = 1 And Status(Invitado) = 3 Or Status(UserIndex) = 3 And Status(Invitado) = 1 Or Status(UserIndex) = 0 And Status(Invitado) = 2 Or Status(UserIndex) = 2 And Status(Invitado) = 0 Then
            If Abs(CInt(UserList(Invitado).Stats.ELV) - CInt(UserList(UserIndex).Stats.ELV)) < 6 Then
                If UserList(Invitado).Grupo.EnGrupo = False Then
                    Call WriteLocaleMsg(UserIndex, "42", FontTypeNames.FONTTYPE_New_GRUPO)
                    'Call WriteConsoleMsg(userindex, "Se envió la invitación a " & UserList(Invitado).name & ", ahora solo resta aguardar su respuesta.", FontTypeNames.FONTTYPE_INFOIAO)
                    Call WriteConsoleMsg(Invitado, UserList(UserIndex).name & " te invitó a unirse a su grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
                    UserList(Invitado).Grupo.PropuestaDe = UserIndex
                    UserList(Invitado).flags.pregunta = 1
                    UserList(Invitado).Grupo.Lider = UserIndex

                    Dim pregunta As String

                    pregunta = UserList(UserIndex).name & " te invitó a unirse a su grupo. ¿Deseas unirte?"
                    Call WritePreguntaBox(Invitado, pregunta)
                Else
                    'Call WriteConsoleMsg(userindex, "El usuario ya se encuentra en un grupo.", FontTypeNames.FONTTYPE_INFOIAO)
                    Call WriteLocaleMsg(UserIndex, "41", FontTypeNames.FONTTYPE_New_GRUPO)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "No podés crear un grupo con personajes con diferencia de mas de 5 niveles.", FontTypeNames.FONTTYPE_New_GRUPO)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "No podés crear un grupo con personajes de diferentes facciones.", FontTypeNames.FONTTYPE_New_GRUPO)

        End If

    Else
        Call WriteConsoleMsg(UserIndex, "El usuario debe desactivar el seguro de grupos para poder invitarlo.", FontTypeNames.FONTTYPE_New_GRUPO)

    End If

End Sub

Public Sub HecharMiembro(ByVal UserIndex As Integer, ByVal indice As Byte)

    Dim i               As Byte

    Dim LoopC           As Byte

    Dim indexviejo      As Byte

    Dim UserIndexHechar As Integer

    If UserList(UserIndex).Grupo.EnGrupo Then
        If UserList(UserIndex).Grupo.Lider = UserIndex Then
    
            UserIndexHechar = UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(indice + 1)

            If UserIndexHechar <> UserIndex Then
    
                For i = 1 To 6

                    If UserIndexHechar = UserList(UserIndex).Grupo.Miembros(i) Then
                        UserList(UserIndex).Grupo.Miembros(i) = 0
                        indexviejo = i

                        For LoopC = indexviejo To 5
                            UserList(UserIndex).Grupo.Miembros(LoopC) = UserList(UserIndex).Grupo.Miembros(LoopC + 1)
                        Next LoopC

                        i = 6

                    End If

                Next i
            
                UserList(UserIndex).Grupo.CantidadMiembros = UserList(UserIndex).Grupo.CantidadMiembros - 1
                    
                Dim a As Byte
                               
                For a = 1 To UserList(UserIndex).Grupo.CantidadMiembros
                    Call WriteUbicacion(UserList(UserIndex).Grupo.Miembros(a), indexviejo, 0)
                        
                Next a
                    
                Call WriteConsoleMsg(UserIndex, UserList(UserIndexHechar).name & " fue expulsado del grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
                'Call WriteConsoleMsg(UserIndexHechar, "Fuiste eliminado del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
                Call WriteLocaleMsg(UserIndexHechar, "37", FontTypeNames.FONTTYPE_New_GRUPO)
                UserList(UserIndexHechar).Grupo.EnGrupo = False
                UserList(UserIndexHechar).Grupo.Lider = 0
                UserList(UserIndexHechar).Grupo.PropuestaDe = 0
                UserList(UserIndexHechar).Grupo.CantidadMiembros = 0
                UserList(UserIndexHechar).Grupo.Miembros(1) = 0
                            
                Call RefreshCharStatus(UserIndexHechar)
                    
                If UserList(UserIndex).Grupo.CantidadMiembros = 1 Then
                    ' Call WriteConsoleMsg(userindex, "El grupo se quedo sin miembros, grupo finalizado.", FontTypeNames.FONTTYPE_INFOIAO)
                    Call WriteLocaleMsg(UserIndex, "35", FontTypeNames.FONTTYPE_New_GRUPO)
                    UserList(UserIndex).Grupo.EnGrupo = False
                    UserList(UserIndex).Grupo.Lider = 0
                    UserList(UserIndex).Grupo.PropuestaDe = 0
                    UserList(UserIndex).Grupo.CantidadMiembros = 0
                    UserList(UserIndex).Grupo.Miembros(1) = 0

                End If
            
                'sera esto?
                'UserList(UserIndex).Grupo.Lider = 0
                Call RefreshCharStatus(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No podés expulsarte a ti mismo.", FontTypeNames.FONTTYPE_New_GRUPO)

            End If
    
        Else
            Call WriteConsoleMsg(UserIndex, "Tu no podés hechar usuarios del grupo.", FontTypeNames.FONTTYPE_New_GRUPO)

        End If

    Else
        Call WriteConsoleMsg(UserIndex, "No estas en ningun grupo.", FontTypeNames.FONTTYPE_New_GRUPO)

    End If

End Sub

Public Sub SalirDeGrupo(ByVal UserIndex As Integer)

    Dim i          As Byte

    Dim LoopC      As Byte

    Dim indexviejo As Byte

    If UserList(UserIndex).Grupo.EnGrupo = True Then
        UserList(UserIndex).Grupo.EnGrupo = False
    
        For i = 1 To 6

            If UserList(UserIndex).name = UserList(UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)).name Then
                UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i) = 0
                indexviejo = i

                For LoopC = indexviejo To 5
                    UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(LoopC) = UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(LoopC + 1)
                Next LoopC

                i = 6

            End If

        Next i

        UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros = UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros - 1
        
        Dim a As Byte
                   
        For a = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
            Call WriteUbicacion(UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(a), indexviejo, 0)
            
        Next a
        
        'Call WriteConsoleMsg(userindex, "Has salido del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
        Call WriteLocaleMsg(UserIndex, "37", FontTypeNames.FONTTYPE_New_GRUPO)
        'Call WriteConsoleMsg(UserList(userindex).Grupo.Lider, UserList(userindex).name & " a salido del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
        Call WriteLocaleMsg(UserList(UserIndex).Grupo.Lider, "202", FontTypeNames.FONTTYPE_New_GRUPO, UserList(UserIndex).name)
        
        If UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros = 1 Then
            'Call WriteConsoleMsg(UserList(userindex).Grupo.Lider, "El grupo se quedo sin miembros, grupo finalizado.", FontTypeNames.FONTTYPE_INFOIAO)
            Call WriteLocaleMsg(UserList(UserIndex).Grupo.Lider, "35", FontTypeNames.FONTTYPE_New_GRUPO)
            
            Call WriteUbicacion(UserList(UserIndex).Grupo.Lider, 1, 0)
                
            UserList(UserList(UserIndex).Grupo.Lider).Grupo.EnGrupo = False
            UserList(UserList(UserIndex).Grupo.Lider).Grupo.Lider = 0
            UserList(UserList(UserIndex).Grupo.Lider).Grupo.PropuestaDe = 0
            UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros = 0
            UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(1) = 0
                
            Call RefreshCharStatus(UserList(UserIndex).Grupo.Lider)

        End If

        Call WriteUbicacion(UserIndex, 1, 0)
        UserList(UserIndex).Grupo.Lider = 0
        Call RefreshCharStatus(UserIndex)
    Else
        Call WriteConsoleMsg(UserIndex, "No estas en ningun grupo.", FontTypeNames.FONTTYPE_New_GRUPO)

    End If

End Sub

Public Sub SalirDeGrupoForzado(ByVal UserIndex As Integer)

    Dim i          As Byte

    Dim LoopC      As Byte

    Dim indexviejo As Byte

    UserList(UserIndex).Grupo.EnGrupo = False
    
    For i = 1 To 6

        If UserList(UserIndex).name = UserList(UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)).name Then
            UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i) = 0
            indexviejo = i

            For LoopC = indexviejo To 5
                UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(LoopC) = UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(LoopC + 1)
            Next LoopC

            i = 6

        End If

    Next i

    UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros = UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros - 1
        
    Dim a As Byte
                   
    For a = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros
        Call WriteUbicacion(UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(a), indexviejo, 0)
            
    Next a

    'Call WriteConsoleMsg(UserList(userindex).Grupo.Lider, UserList(userindex).name & " a salido del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
    Call WriteLocaleMsg(UserList(UserIndex).Grupo.Lider, "202", FontTypeNames.FONTTYPE_New_GRUPO, UserList(UserIndex).name)
        
    If UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros = 1 Then
        'Call WriteConsoleMsg(UserList(userindex).Grupo.Lider, "El grupo se quedo sin miembros, grupo finalizado.", FontTypeNames.FONTTYPE_INFOIAO)
        Call WriteLocaleMsg(UserList(UserIndex).Grupo.Lider, "35", FontTypeNames.FONTTYPE_New_GRUPO)
            
        Call WriteUbicacion(UserList(UserIndex).Grupo.Lider, 1, 0)
                
        UserList(UserList(UserIndex).Grupo.Lider).Grupo.EnGrupo = False
        UserList(UserList(UserIndex).Grupo.Lider).Grupo.Lider = 0
        UserList(UserList(UserIndex).Grupo.Lider).Grupo.PropuestaDe = 0
        UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros = 0
        UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(1) = 0
                
        Call RefreshCharStatus(UserList(UserIndex).Grupo.Lider)

    End If

End Sub

Public Sub FinalizarGrupo(ByVal UserIndex As Integer)

    Dim i As Byte

    For i = 2 To UserList(UserIndex).Grupo.CantidadMiembros
        UserList(UserList(UserIndex).Grupo.Miembros(i)).Grupo.EnGrupo = False
        UserList(UserList(UserIndex).Grupo.Miembros(i)).Grupo.Lider = 0
        UserList(UserList(UserIndex).Grupo.Miembros(i)).Grupo.PropuestaDe = 0
        Call WriteUbicacion(UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i), i, 0)
    
        Call WriteConsoleMsg(UserList(UserIndex).Grupo.Miembros(i), "El lider abandonado el grupo, grupo finalizado.", FontTypeNames.FONTTYPE_New_GRUPO)
        Call RefreshCharStatus(UserList(UserIndex).Grupo.Miembros(i))
        Call WriteUbicacion(UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i), 1, 0)
    
        UserList(UserIndex).Grupo.EnGrupo = False
    Next i

End Sub

Public Sub CompartirUbicacion(UserIndex)

    Dim i       As Byte

    Dim a       As Byte

    Dim indexpj As Byte

    For a = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros

        If UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(a) = UserIndex Then
            indexpj = a

        End If

    Next a

    For i = 1 To UserList(UserList(UserIndex).Grupo.Lider).Grupo.CantidadMiembros

        If UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i) <> UserIndex Then
        
            If UserList(UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i)).Pos.Map = UserList(UserIndex).Pos.Map Then
                Call WriteUbicacion(UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i), indexpj, UserIndex)
            Else
                Call WriteUbicacion(UserList(UserList(UserIndex).Grupo.Lider).Grupo.Miembros(i), indexpj, 0)

            End If

        End If
    
    Next i

End Sub

