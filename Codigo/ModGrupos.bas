Attribute VB_Name = "ModGrupos"

Type Tgrupo

    EnGrupo As Boolean
    CantidadMiembros As Byte
    Miembros(1 To 6) As Integer
    Lider As Integer
    PropuestaDe As Integer

End Type

Public Grupo As Tgrupo

Public Sub InvitarMiembro(ByVal Userindex As Integer, ByVal Invitado As Integer)
        
        On Error GoTo InvitarMiembro_Err
        
100     If UserList(Invitado).flags.SeguroParty = False Then
            
102         If UserList(Userindex).Grupo.CantidadMiembros >= UBound(UserList(Userindex).Grupo.Miembros) Then
104             Call WriteConsoleMsg(Userindex, "No puedes invitar a mas personas. (Límite: " & CStr(UBound(UserList(Userindex).Grupo.Miembros)) & ")", FontTypeNames.FONTTYPE_New_GRUPO)
                Exit Sub
            End If
            
106         If Status(Userindex) = Status(Invitado) Or _
                Status(Userindex) = 1 And Status(Invitado) = 3 Or _
                Status(Userindex) = 3 And Status(Invitado) = 1 Or _
                Status(Userindex) = 0 And Status(Invitado) = 2 Or _
                Status(Userindex) = 2 And Status(Invitado) = 0 Then

108             If Abs(CInt(UserList(Invitado).Stats.ELV) - CInt(UserList(Userindex).Stats.ELV)) < 6 Then

110                 If UserList(Invitado).Grupo.EnGrupo = False Then

112                     Call WriteLocaleMsg(Userindex, "42", FontTypeNames.FONTTYPE_New_GRUPO)
                        'Call WriteConsoleMsg(userindex, "Se envió la invitación a " & UserList(Invitado).name & ", ahora solo resta aguardar su respuesta.", FontTypeNames.FONTTYPE_INFOIAO)
114                     Call WriteConsoleMsg(Invitado, UserList(Userindex).name & " te invitó a unirse a su grupo.", FontTypeNames.FONTTYPE_New_GRUPO)

116                     UserList(Invitado).Grupo.PropuestaDe = Userindex
118                     UserList(Invitado).flags.pregunta = 1
120                     UserList(Invitado).Grupo.Lider = Userindex

                        Dim pregunta As String
122                         pregunta = UserList(Userindex).name & " te invitó a unirse a su grupo. ¿Deseas unirte?"

124                     Call WritePreguntaBox(Invitado, pregunta)

                    Else
                        'Call WriteConsoleMsg(userindex, "El usuario ya se encuentra en un grupo.", FontTypeNames.FONTTYPE_INFOIAO)
126                     Call WriteLocaleMsg(Userindex, "41", FontTypeNames.FONTTYPE_New_GRUPO)

                    End If

                Else
128                 Call WriteConsoleMsg(Userindex, "No podés crear un grupo con personajes con diferencia de mas de 5 niveles.", FontTypeNames.FONTTYPE_New_GRUPO)

                End If

            Else
130             Call WriteConsoleMsg(Userindex, "No podés crear un grupo con personajes de diferentes facciones.", FontTypeNames.FONTTYPE_New_GRUPO)

            End If

        Else
132         Call WriteConsoleMsg(Userindex, "El usuario debe desactivar el seguro de grupos para poder invitarlo.", FontTypeNames.FONTTYPE_New_GRUPO)

        End If

        
        Exit Sub

InvitarMiembro_Err:
134     Call RegistrarError(Err.Number, Err.description, "ModGrupos.InvitarMiembro", Erl)
136     Resume Next
        
End Sub

Public Sub HecharMiembro(ByVal Userindex As Integer, ByVal Indice As Byte)
        
        On Error GoTo HecharMiembro_Err
        

        Dim i               As Byte

        Dim LoopC           As Byte

        Dim indexviejo      As Byte

        Dim UserIndexHechar As Integer

100     If UserList(Userindex).Grupo.EnGrupo Then
102         If UserList(Userindex).Grupo.Lider = Userindex Then
    
104             UserIndexHechar = UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(Indice + 1)

106             If UserIndexHechar <> Userindex Then
    
108                 For i = 1 To 6

110                     If UserIndexHechar = UserList(Userindex).Grupo.Miembros(i) Then
112                         UserList(Userindex).Grupo.Miembros(i) = 0
114                         indexviejo = i

116                         For LoopC = indexviejo To 5
118                             UserList(Userindex).Grupo.Miembros(LoopC) = UserList(Userindex).Grupo.Miembros(LoopC + 1)
120                         Next LoopC

122                         i = 6

                        End If

124                 Next i
            
126                 UserList(Userindex).Grupo.CantidadMiembros = UserList(Userindex).Grupo.CantidadMiembros - 1
                    
                    Dim a As Byte
                               
128                 For a = 1 To UserList(Userindex).Grupo.CantidadMiembros
130                     Call WriteUbicacion(UserList(Userindex).Grupo.Miembros(a), indexviejo, 0)
                        
132                 Next a
                    
134                 Call WriteConsoleMsg(Userindex, UserList(UserIndexHechar).name & " fue expulsado del grupo.", FontTypeNames.FONTTYPE_New_GRUPO)
                    'Call WriteConsoleMsg(UserIndexHechar, "Fuiste eliminado del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
136                 Call WriteLocaleMsg(UserIndexHechar, "37", FontTypeNames.FONTTYPE_New_GRUPO)
138                 UserList(UserIndexHechar).Grupo.EnGrupo = False
140                 UserList(UserIndexHechar).Grupo.Lider = 0
142                 UserList(UserIndexHechar).Grupo.PropuestaDe = 0
144                 UserList(UserIndexHechar).Grupo.CantidadMiembros = 0
146                 UserList(UserIndexHechar).Grupo.Miembros(1) = 0
                            
148                 Call RefreshCharStatus(UserIndexHechar)
                    
150                 If UserList(Userindex).Grupo.CantidadMiembros = 1 Then
                        ' Call WriteConsoleMsg(userindex, "El grupo se quedo sin miembros, grupo finalizado.", FontTypeNames.FONTTYPE_INFOIAO)
152                     Call WriteLocaleMsg(Userindex, "35", FontTypeNames.FONTTYPE_New_GRUPO)
154                     UserList(Userindex).Grupo.EnGrupo = False
156                     UserList(Userindex).Grupo.Lider = 0
158                     UserList(Userindex).Grupo.PropuestaDe = 0
160                     UserList(Userindex).Grupo.CantidadMiembros = 0
162                     UserList(Userindex).Grupo.Miembros(1) = 0

                    End If
            
                    'sera esto?
                    'UserList(UserIndex).Grupo.Lider = 0
164                 Call RefreshCharStatus(Userindex)
                Else
166                 Call WriteConsoleMsg(Userindex, "No podés expulsarte a ti mismo.", FontTypeNames.FONTTYPE_New_GRUPO)

                End If
    
            Else
168             Call WriteConsoleMsg(Userindex, "Tu no podés hechar usuarios del grupo.", FontTypeNames.FONTTYPE_New_GRUPO)

            End If

        Else
170         Call WriteConsoleMsg(Userindex, "No estas en ningun grupo.", FontTypeNames.FONTTYPE_New_GRUPO)

        End If

        
        Exit Sub

HecharMiembro_Err:
172     Call RegistrarError(Err.Number, Err.description, "ModGrupos.HecharMiembro", Erl)
174     Resume Next
        
End Sub

Public Sub SalirDeGrupo(ByVal Userindex As Integer)
        
        On Error GoTo SalirDeGrupo_Err
        

        Dim i          As Byte

        Dim LoopC      As Byte

        Dim indexviejo As Byte

100     If UserList(Userindex).Grupo.EnGrupo = True Then
102         UserList(Userindex).Grupo.EnGrupo = False
    
104         For i = 1 To 6

106             If UserList(Userindex).name = UserList(UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(i)).name Then
108                 UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(i) = 0
110                 indexviejo = i

112                 For LoopC = indexviejo To 5
114                     UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(LoopC) = UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(LoopC + 1)
116                 Next LoopC

118                 i = 6

                End If

120         Next i

122         UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros = UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros - 1
        
            Dim a As Byte
                   
124         For a = 1 To UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros
126             Call WriteUbicacion(UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(a), indexviejo, 0)
            
128         Next a
        
            'Call WriteConsoleMsg(userindex, "Has salido del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
130         Call WriteLocaleMsg(Userindex, "37", FontTypeNames.FONTTYPE_New_GRUPO)
            'Call WriteConsoleMsg(UserList(userindex).Grupo.Lider, UserList(userindex).name & " a salido del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
132         Call WriteLocaleMsg(UserList(Userindex).Grupo.Lider, "202", FontTypeNames.FONTTYPE_New_GRUPO, UserList(Userindex).name)
        
134         If UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros = 1 Then
                'Call WriteConsoleMsg(UserList(userindex).Grupo.Lider, "El grupo se quedo sin miembros, grupo finalizado.", FontTypeNames.FONTTYPE_INFOIAO)
136             Call WriteLocaleMsg(UserList(Userindex).Grupo.Lider, "35", FontTypeNames.FONTTYPE_New_GRUPO)
            
138             Call WriteUbicacion(UserList(Userindex).Grupo.Lider, 1, 0)
                
140             UserList(UserList(Userindex).Grupo.Lider).Grupo.EnGrupo = False
142             UserList(UserList(Userindex).Grupo.Lider).Grupo.Lider = 0
144             UserList(UserList(Userindex).Grupo.Lider).Grupo.PropuestaDe = 0
146             UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros = 0
148             UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(1) = 0
                
150             Call RefreshCharStatus(UserList(Userindex).Grupo.Lider)

            End If

152         Call WriteUbicacion(Userindex, 1, 0)
154         UserList(Userindex).Grupo.Lider = 0
156         Call RefreshCharStatus(Userindex)
        Else
158         Call WriteConsoleMsg(Userindex, "No estas en ningun grupo.", FontTypeNames.FONTTYPE_New_GRUPO)

        End If

        
        Exit Sub

SalirDeGrupo_Err:
160     Call RegistrarError(Err.Number, Err.description, "ModGrupos.SalirDeGrupo", Erl)
162     Resume Next
        
End Sub

Public Sub SalirDeGrupoForzado(ByVal Userindex As Integer)
        
        On Error GoTo SalirDeGrupoForzado_Err
        

        Dim i          As Byte

        Dim LoopC      As Byte

        Dim indexviejo As Byte

100     UserList(Userindex).Grupo.EnGrupo = False
    
102     For i = 1 To 6

104         If UserList(Userindex).name = UserList(UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(i)).name Then
106             UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(i) = 0
108             indexviejo = i

110             For LoopC = indexviejo To 5
112                 UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(LoopC) = UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(LoopC + 1)
114             Next LoopC

116             i = 6

            End If

118     Next i

120     UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros = UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros - 1
        
        Dim a As Byte
                   
122     For a = 1 To UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros
124         Call WriteUbicacion(UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(a), indexviejo, 0)
            
126     Next a

        'Call WriteConsoleMsg(UserList(userindex).Grupo.Lider, UserList(userindex).name & " a salido del grupo.", FontTypeNames.FONTTYPE_INFOIAO)
128     Call WriteLocaleMsg(UserList(Userindex).Grupo.Lider, "202", FontTypeNames.FONTTYPE_New_GRUPO, UserList(Userindex).name)
        
130     If UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros = 1 Then
            'Call WriteConsoleMsg(UserList(userindex).Grupo.Lider, "El grupo se quedo sin miembros, grupo finalizado.", FontTypeNames.FONTTYPE_INFOIAO)
132         Call WriteLocaleMsg(UserList(Userindex).Grupo.Lider, "35", FontTypeNames.FONTTYPE_New_GRUPO)
            
134         Call WriteUbicacion(UserList(Userindex).Grupo.Lider, 1, 0)
                
136         UserList(UserList(Userindex).Grupo.Lider).Grupo.EnGrupo = False
138         UserList(UserList(Userindex).Grupo.Lider).Grupo.Lider = 0
140         UserList(UserList(Userindex).Grupo.Lider).Grupo.PropuestaDe = 0
142         UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros = 0
144         UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(1) = 0
                
146         Call RefreshCharStatus(UserList(Userindex).Grupo.Lider)

        End If

        
        Exit Sub

SalirDeGrupoForzado_Err:
148     Call RegistrarError(Err.Number, Err.description, "ModGrupos.SalirDeGrupoForzado", Erl)
150     Resume Next
        
End Sub

Public Sub FinalizarGrupo(ByVal Userindex As Integer)
        
        On Error GoTo FinalizarGrupo_Err
        

        Dim i As Byte

100     For i = 2 To UserList(Userindex).Grupo.CantidadMiembros
102         UserList(UserList(Userindex).Grupo.Miembros(i)).Grupo.EnGrupo = False
104         UserList(UserList(Userindex).Grupo.Miembros(i)).Grupo.Lider = 0
106         UserList(UserList(Userindex).Grupo.Miembros(i)).Grupo.PropuestaDe = 0
108         Call WriteUbicacion(UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(i), i, 0)
    
110         Call WriteConsoleMsg(UserList(Userindex).Grupo.Miembros(i), "El lider abandonado el grupo, grupo finalizado.", FontTypeNames.FONTTYPE_New_GRUPO)
112         Call RefreshCharStatus(UserList(Userindex).Grupo.Miembros(i))
114         Call WriteUbicacion(UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(i), 1, 0)
    
116         UserList(Userindex).Grupo.EnGrupo = False
118     Next i

        
        Exit Sub

FinalizarGrupo_Err:
120     Call RegistrarError(Err.Number, Err.description, "ModGrupos.FinalizarGrupo", Erl)
122     Resume Next
        
End Sub

Public Sub CompartirUbicacion(Userindex)
        
        On Error GoTo CompartirUbicacion_Err
        

        Dim i       As Byte

        Dim a       As Byte

        Dim indexpj As Byte

100     For a = 1 To UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros

102         If UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(a) = Userindex Then
104             indexpj = a

            End If

106     Next a

108     For i = 1 To UserList(UserList(Userindex).Grupo.Lider).Grupo.CantidadMiembros

110         If UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(i) <> Userindex Then
        
112             If UserList(UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(i)).Pos.Map = UserList(Userindex).Pos.Map Then
114                 Call WriteUbicacion(UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(i), indexpj, Userindex)
                Else
116                 Call WriteUbicacion(UserList(UserList(Userindex).Grupo.Lider).Grupo.Miembros(i), indexpj, 0)

                End If

            End If
    
118     Next i

        
        Exit Sub

CompartirUbicacion_Err:
120     Call RegistrarError(Err.Number, Err.description, "ModGrupos.CompartirUbicacion", Erl)
122     Resume Next
        
End Sub

