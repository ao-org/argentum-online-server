Attribute VB_Name = "AI"
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

Public Enum TipoAI

    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    GuardiasAtacanCiudadanos = 6
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcPathfinding = 10

    'Pretorianos
    SacerdotePretorianoAi = 11
    GuerreroPretorianoAi = 12
    MagoPretorianoAi = 13
    CazadorPretorianoAi = 14
    ReyPretoriano = 15

    ' Animado
    Caminata = 20
    
    ' Eventos
    Invasion = 21

End Enum

' WyroX: Hardcodeada de la vida...
Public Const ELEMENTALFUEGO  As Integer = 962
Public Const ELEMENTALTIERRA As Integer = 961
Public Const ELEMENTALAGUA   As Integer = 960
Public Const ELEMENTALVIENTO As Integer = 963
Public Const FUEGOFATUO      As Integer = 964

'Damos a los NPCs el mismo rango de visión que un PJ
Public Const RANGO_VISION_X  As Byte = 11
Public Const RANGO_VISION_Y  As Byte = 9

Public Enum e_Alineacion
    ninguna = 0
    Real = 1
    Caos = 2
    Neutro = 3
End Enum

Public Enum e_Personalidad

    ''Inerte: no tiene objetivos de ningun tipo (npcs vendedores, curas, etc)
    ''Agresivo no magico: Su objetivo es acercarse a las victimas para atacarlas
    ''Agresivo magico: Su objetivo es mantenerse lo mas lejos posible de sus victimas y atacarlas con magia
    ''Mascota: Solo ataca a quien ataque a su amo.
    ''Pacifico: No ataca.
    ninguna = 0
    Inerte = 1
    AgresivoNoMagico = 2
    AgresivoMagico = 3
    Macota = 4
    Pacifico = 5

End Enum

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
        
    On Error GoTo SeguirAgresor_Err
        

    Dim tHeading As Byte
    Dim UI       As Integer
    Dim i        As Long
    Dim SignoNS  As Integer
    Dim SignoEO  As Integer
    
    With NpcList(NpcIndex)

        If .flags.Inmovilizado = 1 Then

            Select Case .Char.Heading

                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0

            End Select
            
            For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                            
                        If UserList(UI).name = .flags.AttackedBy Then
                            
                            If PuedeAtacarUser(UI) Then

                                If .flags.LanzaSpells > 0 Then

                                    Call AnimacionIdle(NpcIndex, True)
                                    Call NpcLanzaUnSpell(NpcIndex, UI)
                                        
                                End If

                                Exit Sub

                            End If

                        End If
                        
                    End If

                End If
                
            Next i

        Else
  
            For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                    
                        If UserList(UI).name = .flags.AttackedBy Then
                               
                            If PuedeAtacarUser(UI) Then

                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)
                                End If
                                    
                                tHeading = FindDirectionEAO(.Pos, UserList(UI).Pos, NpcList(NpcIndex).flags.AguaValida = 1, NpcList(NpcIndex).flags.TierraInvalida = 0)
                                                         
                                If Distancia(.Pos, UserList(UI).Pos) = 1 Then
                                    Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)
                                    Call AnimacionIdle(NpcIndex, True)
                                    Call NpcAtacaUser(NpcIndex, UI, tHeading)

                                Else
                                    
                                    Call MoveNPCChar(NpcIndex, tHeading)

                                End If

                                Exit Sub

                            End If

                        End If
                        
                    End If

                End If
                
            Next i

        End If

    End With
    
    Call RestoreOldMovement(NpcIndex)

        
    Exit Sub

SeguirAgresor_Err:
    Call RegistrarError(Err.Number, Err.Description, "AI.SeguirAgresor", Erl)
    Resume Next
        
End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)
        
    On Error GoTo RestoreOldMovement_Err
        

    With NpcList(NpcIndex)

        If .MaestroUser = 0 Then
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = vbNullString
            .Target = 0
        End If

    End With

    Exit Sub

RestoreOldMovement_Err:
    Call RegistrarError(Err.Number, Err.Description, "AI.RestoreOldMovement", Erl)
    Resume Next
        
End Sub

Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)
        
        On Error GoTo PersigueCiudadano_Err
        

        Dim UI       As Integer
        Dim tHeading As Byte
        Dim i        As Long
    
100     With NpcList(NpcIndex)

102         For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
104             UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                'Is it in it's range of vision??
106             If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
108                 If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
        
110                     If Status(UI) = 1 Or Status(UI) = 3 Then
                        
112                         If PuedeAtacarUser(UI) Then

114                             If .flags.LanzaSpells > 0 Then
116                                 Call NpcLanzaUnSpell(NpcIndex, UI)
                                End If

118                             tHeading = FindDirectionEAO(.Pos, UserList(UI).Pos, NpcList(NpcIndex).flags.AguaValida = 1, NpcList(NpcIndex).flags.TierraInvalida = 0)

120                             Call MoveNPCChar(NpcIndex, tHeading)

                                Exit Sub

                            End If

                        End If
                    
                    End If

                End If
            
122         Next i

        End With
    
124     Call RestoreOldMovement(NpcIndex)

        
        Exit Sub

PersigueCiudadano_Err:
126     Call RegistrarError(Err.Number, Err.Description, "AI.PersigueCiudadano", Erl)
128     Resume Next
        
End Sub

Private Sub CuraResucita(ByVal NpcIndex As Integer)
        
        On Error GoTo CuraResucita_Err
        

        Dim UI       As Integer
        Dim tHeading As Byte
        Dim i        As Long
    
100     With NpcList(NpcIndex)

102         For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
104             UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                'Is it in it's range of vision??
106             If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
108                 If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then

110                     If Not UserList(UI).Accion.Particula = ParticulasIndex.Resucitar Then

112                         If Status(UI) < 2 Then

114                             If UserList(UI).flags.Muerto = 1 Then
116                                 Call SendData(SendTarget.ToPCArea, UI, PrepareMessageParticleFX(UserList(UI).Char.CharIndex, ParticulasIndex.Resucitar, 250, False))
118                                 Call SendData(SendTarget.ToPCArea, UI, PrepareMessageBarFx(UserList(UI).Char.CharIndex, 250, Accion_Barra.Resucitar))

120                                 UserList(UI).Accion.AccionPendiente = True
122                                 UserList(UI).Accion.Particula = ParticulasIndex.Resucitar
124                                 UserList(UI).Accion.TipoAccion = Accion_Barra.Resucitar

                                Else

126                                 If UserList(UI).Stats.MinHp <> UserList(UI).Stats.MaxHp Then
128                                     UserList(UI).Stats.MinHp = UserList(UI).Stats.MaxHp

130                                     Call WriteUpdateUserStats(UI)
132                                     Call SendData(SendTarget.ToPCArea, UI, PrepareMessageParticleFX(UserList(UI).Char.CharIndex, ParticulasIndex.Curar, 100, False))

                                    End If

                                End If

                            End If

                        End If

                        Exit Sub
                        ' End If
                        ' End If
                    
                    End If

                End If
            
134         Next i

        End With

        
        Exit Sub

CuraResucita_Err:
136     Call RegistrarError(Err.Number, Err.Description, "AI.CuraResucita", Erl)
138     Resume Next
        
End Sub

Private Sub PersigueCriminal(ByVal NpcIndex As Integer)
        
        On Error GoTo PersigueCriminal_Err
        

        Dim UI       As Integer
        Dim tHeading As Byte
        Dim i        As Long
        Dim SignoNS  As Integer
        Dim SignoEO  As Integer
    
100     With NpcList(NpcIndex)

102         If .flags.Inmovilizado = 1 Then

104             Select Case .Char.Heading

                    Case eHeading.NORTH
106                     SignoNS = -1
108                     SignoEO = 0
                
110                 Case eHeading.EAST
112                     SignoNS = 0
114                     SignoEO = 1
                
116                 Case eHeading.SOUTH
118                     SignoNS = 1
120                     SignoEO = 0
                
122                 Case eHeading.WEST
124                     SignoEO = -1
126                     SignoNS = 0

                End Select
            
128             For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
130                 UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                    'Is it in it's range of vision??
132                 If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
134                     If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                        
136                         If Status(UI) = 0 Or Status(UI) = 2 Then

138                             If PuedeAtacarUser(UI) Then

140                                 If .flags.LanzaSpells > 0 Then
142                                     Call NpcLanzaUnSpell(NpcIndex, UI)
                                    End If

                                    Exit Sub

                                End If

                            End If
                        
                        End If

                    End If
                    
144             Next i

            Else

146             For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
148                 UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                    'Is it in it's range of vision??
150                 If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
152                     If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
154                         If Status(UI) = 0 Or Status(UI) = 2 Then

156                             If PuedeAtacarUser(UI) Then

158                                 If .flags.LanzaSpells > 0 Then
160                                     Call NpcLanzaUnSpell(NpcIndex, UI)
                                    End If

162                                 If Distancia(.Pos, UserList(UI).Pos) > 1 Then
164                                     tHeading = FindDirectionEAO(.Pos, UserList(UI).Pos, NpcList(NpcIndex).flags.AguaValida = 1, NpcList(NpcIndex).flags.TierraInvalida = 0)
166                                     Call MoveNPCChar(NpcIndex, tHeading)
                                    
                                    Else
                                    
168                                     If .Pos.Y > UserList(UI).Pos.Y Then
170                                         tHeading = 1

172                                     ElseIf .Pos.X < UserList(UI).Pos.X Then
174                                         tHeading = 2

176                                     ElseIf .Pos.Y < UserList(UI).Pos.Y Then
178                                         tHeading = 3

                                        Else
180                                         tHeading = 4

                                        End If

182                                     If NpcAtacaUser(NpcIndex, UI, tHeading) Then
184                                         Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)
                                        End If
                                        
                                    End If
                                    
                                    Exit Sub

                                End If

                            End If
                        
                        End If

                    End If
                
186             Next i

            End If

        End With
    
188     Call RestoreOldMovement(NpcIndex)

190     Call AnimacionIdle(NpcIndex, True)

        
        Exit Sub

PersigueCriminal_Err:
192     Call RegistrarError(Err.Number, Err.Description, "AI.PersigueCriminal", Erl)
194     Resume Next
        
End Sub

Private Sub SeguirAmo(ByVal NpcIndex As Integer)

        On Error GoTo SeguirAmo_Err

        Dim tHeading As Byte
        Dim UI As Integer
        
100     If NpcIndex = 0 Then Exit Sub
    
102     With NpcList(NpcIndex)
            
104         If .MaestroUser = 0 Then Exit Sub
            
106         If .Target = 0 And .TargetNPC = 0 Then
108             UI = .MaestroUser
            
                'Is it in it's range of vision??
110             If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then

112                 If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then

114                     If UserList(UI).flags.Muerto = 0 _
                                And UserList(UI).flags.invisible = 0 _
                                And UserList(UI).flags.Oculto = 0 _
                                And Distancia(.Pos, UserList(UI).Pos) > 3 Then
                                
116                         tHeading = FindDirectionEAO(.Pos, UserList(UI).Pos)

118                         Call MoveNPCChar(NpcIndex, tHeading)

                            Exit Sub
                            
                        Else
                        
120                         If RandomNumber(1, 12) = 3 Then
122                             Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

                            Else
124                             Call AnimacionIdle(NpcIndex, True)

                            End If
                            
                        End If
                        
                    End If
                    
                End If
                
            End If
            
        End With
    
126     Call RestoreOldMovement(NpcIndex)

        Exit Sub

SeguirAmo_Err:
128     Call RegistrarError(Err.Number, Err.Description, "AI.SeguirAmo", Erl)
130     Resume Next

End Sub


Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)

        On Error GoTo SeguirAmo_Err

        Dim tHeading As Byte
        Dim X As Long
        Dim Y As Long
        Dim NI As Integer
        Dim bNoEsta As Boolean
    
        Dim SignoNS As Integer
        Dim SignoEO As Integer
    
100     With NpcList(NpcIndex)
102         If .flags.Inmovilizado = 1 Then

104             Select Case .Char.Heading
                    Case eHeading.NORTH
106                     SignoNS = -1
108                     SignoEO = 0
                
110                 Case eHeading.EAST
112                     SignoNS = 0
114                     SignoEO = 1
                
116                 Case eHeading.SOUTH
118                     SignoNS = 1
120                     SignoEO = 0
                
122                 Case eHeading.WEST
124                     SignoEO = -1
126                     SignoNS = 0
                End Select
            
128             For Y = .Pos.Y To .Pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
130                 For X = .Pos.X To .Pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)
132                     If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then

134                         NI = MapData(.Pos.Map, X, Y).NpcIndex

136                         If NI > 0 Then
138                             If .TargetNPC = NI Then
140                                 bNoEsta = True
142                                 If .Numero = ELEMENTALFUEGO Then

144                                     Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)

146                                     If NpcList(NI).NPCtype = DRAGON Then
148                                         NpcList(NI).CanAttack = 1
150                                         Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                                         End If
                                         
                                     Else
                                     
                                        'aca verificamosss la distancia de ataque
152                                     If Distancia(.Pos, NpcList(NI).Pos) <= 1 Then
154                                         Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                                        End If
                                        
                                     End If
                                     Exit Sub
                                End If
                                
                            End If
                            
                        End If
156                 Next X
158             Next Y

            Else ' No Inmovilizado
            
160             For Y = .Pos.Y - RANGO_VISION_Y To .Pos.Y + RANGO_VISION_Y
162                 For X = .Pos.X - RANGO_VISION_Y To .Pos.X + RANGO_VISION_Y
164                     If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then

166                        NI = MapData(.Pos.Map, X, Y).NpcIndex

168                        If NI > 0 Then

170                             If .TargetNPC = NI Then

172                                  bNoEsta = True

174                                  If .Numero = ELEMENTALFUEGO Then
176                                      Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
178                                      If NpcList(NI).NPCtype = DRAGON Then
180                                         NpcList(NI).CanAttack = 1
182                                         Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                                         End If
                                         Exit Sub
                                     End If

184                                  If .TargetNPC = 0 Then Exit Sub
                                 
186                                  tHeading = FindDirectionEAO(.Pos, NpcList(NI).Pos, .flags.AguaValida = 1, .flags.TierraInvalida = 0)
                                 
188                                 If Distancia(.Pos, NpcList(NI).Pos) <= 1 Then
190                                     Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)
192                                     Call AnimacionIdle(NpcIndex, True)
194                                     Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                                    Else
196                                     Call MoveNPCChar(NpcIndex, tHeading)
                                    End If
                                    
                                    Exit Sub
                                    
                                End If
                                
                           End If
                           
                        End If
198                 Next X
200             Next Y

            End If
        
202         If Not bNoEsta Then

204             If .MaestroUser > 0 Then
206                 Call FollowAmo(NpcIndex)

                Else
208                 .Movement = .flags.OldMovement
210                 .Hostile = .flags.OldHostil
                
212                 Call AnimacionIdle(NpcIndex, True)
                End If
                
            End If
            
        End With
    
        Exit Sub
    
SeguirAmo_Err:
214     Call RegistrarError(Err.Number, Err.Description, "AI.SeguirAmo")
216     Resume Next
End Sub

Sub NPCAI(ByVal NpcIndex As Integer)

        On Error GoTo ErrorHandler

        Dim falladesc As String

100     With NpcList(NpcIndex)

            ' Ningun NPC se puede mover si esta Inmovilizado o Paralizado
            If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then Exit Sub

            '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
102         Select Case .Movement

                Case TipoAI.ESTATICO
                    Rem  Debug.Print "Es un NPC estatico, no hace nada."
104                 falladesc = " fallo en estatico"

106             Case TipoAI.MueveAlAzar
108                 falladesc = " fallo al azar"

112                 If .NPCtype = eNPCType.GuardiaReal Then
114                     Call PersigueCriminal(NpcIndex)

116                 ElseIf .NPCtype = eNPCType.Guardiascaos Then
118                     Call PersigueCiudadano(NpcIndex)

                    Else
120                     If RandomNumber(1, 12) = 3 Then
122                         Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                        Else
124                         Call AnimacionIdle(NpcIndex, True)
                        End If
                    End If

126             Case TipoAI.NpcMaloAtacaUsersBuenos
128                 falladesc = " fallo NpcMaloAtacaUsersBuenos"
130                 Call IrUsuarioCercano(NpcIndex)

                    'Va hacia el usuario que lo ataco(FOLLOW)
132             Case TipoAI.NPCDEFENSA
134                 Call SeguirAgresor2(NpcIndex)

                    'Persigue criminales
136             Case TipoAI.GuardiasAtacanCriminales
138                 Call PersigueCriminal(NpcIndex)

140             Case TipoAI.GuardiasAtacanCiudadanos
142                 Call PersigueCiudadano(NpcIndex)

144             Case TipoAI.NpcAtacaNpc
146                 Call AiNpcAtacaNpc(NpcIndex)

148             Case TipoAI.NpcPathfinding
150                 falladesc = " fallo NpcPathfinding"

154                 If ReCalculatePath(NpcIndex) Then
156                     Call PathFindingAI(NpcIndex)

                        'Existe el camino?
158                     If .PFINFO.NoPath Then 'Si no existe nos movemos al azar
                            'Move randomly
160                         Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))

                        End If

                    Else

162                     If Not PathEnd(NpcIndex) Then
164                         Call FollowPath(NpcIndex)
                        Else
166                         .PFINFO.PathLenght = 0

                        End If

                    End If

168             Case TipoAI.SigueAmo
170                 falladesc = " fallo SigueAmo"

174                 Call SeguirAmo(NpcIndex)

                Case TipoAI.Caminata
                    falladesc = " fallo Caminata"

                    Call HacerCaminata(NpcIndex)
                    
                Case TipoAI.Invasion
                    falladesc = " fallo Invasion"
                    
                    Call MovimientoInvasion(NpcIndex)

            End Select

        End With

        Exit Sub

ErrorHandler:
176     Call LogError("NPCAI " & NpcList(NpcIndex).name & " " & NpcList(NpcIndex).MaestroNPC & " mapa:" & NpcList(NpcIndex).Pos.Map & " x:" & NpcList(NpcIndex).Pos.X & " y:" & NpcList(NpcIndex).Pos.Y & " Mov:" & NpcList(NpcIndex).Movement & " TargU:" & NpcList(NpcIndex).Target & " TargN:" & NpcList(NpcIndex).TargetNPC & falladesc)

        Dim MiNPC As npc

178     MiNPC = NpcList(NpcIndex)
180     Call QuitarNPC(NpcIndex)
182     Call ReSpawnNpc(MiNPC)

End Sub

Function UserNear(ByVal NpcIndex As Integer) As Boolean
        '#################################################################
        'Returns True if there is an user adjacent to the npc position.
        '#################################################################
        
        On Error GoTo UserNear_Err
        
100     UserNear = Not Int(Distance(NpcList(NpcIndex).Pos.X, NpcList(NpcIndex).Pos.Y, UserList(NpcList(NpcIndex).PFINFO.TargetUser).Pos.X, UserList(NpcList(NpcIndex).PFINFO.TargetUser).Pos.Y)) > 1

        
        Exit Function

UserNear_Err:
102     Call RegistrarError(Err.Number, Err.Description, "AI.UserNear", Erl)
104     Resume Next
        
End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean
        
        On Error GoTo ReCalculatePath_Err
        

        '#################################################################
        'Returns true if we have to seek a new path
        '#################################################################
        
100     If NpcList(NpcIndex).PFINFO.PathLenght = 0 Then
102         ReCalculatePath = True

104     ElseIf Not UserNear(NpcIndex) And NpcList(NpcIndex).PFINFO.PathLenght = NpcList(NpcIndex).PFINFO.CurPos - 1 Then
106         ReCalculatePath = True

        End If

        
        Exit Function

ReCalculatePath_Err:
108     Call RegistrarError(Err.Number, Err.Description, "AI.ReCalculatePath", Erl)
110     Resume Next
        
End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
        '#################################################################
        'Coded By Gulfas Morgolock
        'Returns if the npc has arrived to the end of its path
        '#################################################################
        
        On Error GoTo PathEnd_Err
        
100     PathEnd = NpcList(NpcIndex).PFINFO.CurPos = NpcList(NpcIndex).PFINFO.PathLenght

        
        Exit Function

PathEnd_Err:
102     Call RegistrarError(Err.Number, Err.Description, "AI.PathEnd", Erl)
104     Resume Next
        
End Function
 
Function FollowPath(NpcIndex As Integer) As Boolean
        
        On Error GoTo FollowPath_Err
        

        Dim tmpPos   As WorldPos
        Dim tHeading As Byte
 
100     tmpPos.Map = NpcList(NpcIndex).Pos.Map
102     tmpPos.X = NpcList(NpcIndex).PFINFO.Path(NpcList(NpcIndex).PFINFO.CurPos).Y
104     tmpPos.Y = NpcList(NpcIndex).PFINFO.Path(NpcList(NpcIndex).PFINFO.CurPos).X
 
106     tHeading = FindDirectionEAO(NpcList(NpcIndex).Pos, tmpPos, NpcList(NpcIndex).flags.AguaValida = 1, NpcList(NpcIndex).flags.TierraInvalida = 0)
 
108     Call MoveNPCChar(NpcIndex, tHeading)
 
110     NpcList(NpcIndex).PFINFO.CurPos = NpcList(NpcIndex).PFINFO.CurPos + 1
 
        
        Exit Function

FollowPath_Err:
112     Call RegistrarError(Err.Number, Err.Description, "AI.FollowPath", Erl)
114     Resume Next
        
End Function

Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean
        
        On Error GoTo PathFindingAI_Err
        

        '#################################################################
        'Coded By Gulfas Morgolock / 11-07-02
        'www.geocities.com/gmorgolock
        'morgolock@speedy.com.ar
        'This function seeks the shortest path from the Npc
        'to the user's location.
        '#################################################################
        Dim Y As Long
        Dim X As Long
        Dim tmpUserIndex As Integer
    
100     For Y = NpcList(NpcIndex).Pos.Y - 10 To NpcList(NpcIndex).Pos.Y + 10    'Makes a loop that looks at
102         For X = NpcList(NpcIndex).Pos.X - 10 To NpcList(NpcIndex).Pos.X + 10   '5 tiles in every direction
            
                'Make sure tile is legal
104             If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                
                    'look for a user
106                 If MapData(NpcList(NpcIndex).Pos.Map, X, Y).UserIndex > 0 Then

                        'Move towards user
108                     tmpUserIndex = MapData(NpcList(NpcIndex).Pos.Map, X, Y).UserIndex

110                     With UserList(tmpUserIndex)

112                         If PuedeAtacarUser(tmpUserIndex) Then
                                
                                'We have to invert the coordinates, this is because
                                'ORE refers to maps in converse way of my pathfinding
                                'routines.
114                             NpcList(NpcIndex).PFINFO.Target.X = UserList(tmpUserIndex).Pos.Y
116                             NpcList(NpcIndex).PFINFO.Target.Y = UserList(tmpUserIndex).Pos.X 'ops!
118                             NpcList(NpcIndex).PFINFO.TargetUser = tmpUserIndex

120                             Call SeekPath(NpcIndex)

                                Exit Function

                            End If

                        End With

                    End If

                End If

122         Next X
124     Next Y

        
        Exit Function

PathFindingAI_Err:
126     Call RegistrarError(Err.Number, Err.Description, "AI.PathFindingAI", Erl)
128     Resume Next
        
End Function

Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo NpcLanzaUnSpell_Err
        
100     With UserList(UserIndex)
        
102         If Not IntervaloPermiteLanzarHechizo(NpcIndex) Then Exit Sub
104         If NpcList(NpcIndex).Pos.Map <> .Pos.Map Then Exit Sub

106         If .flags.invisible = 1 Or .flags.Oculto = 1 Or .flags.Inmunidad = 1 Or .flags.NoMagiaEfeceto = 1 Or .flags.EnConsulta Then Exit Sub
    
            Dim K As Integer
108             K = RandomNumber(1, NpcList(NpcIndex).flags.LanzaSpells)

110         Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, NpcList(NpcIndex).Spells(K))

112         If NpcList(NpcIndex).Target = 0 Then NpcList(NpcIndex).Target = UserIndex

114         If .flags.AtacadoPorNpc = 0 And .flags.AtacadoPorUser = 0 Then
116             .flags.AtacadoPorNpc = NpcIndex
            End If
        
        End With

        Exit Sub

NpcLanzaUnSpell_Err:
118     Call RegistrarError(Err.Number, Err.Description, "AI.NpcLanzaUnSpell", Erl)

120     Resume Next
        
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)
        
        On Error GoTo NpcLanzaUnSpellSobreNpc_Err
        

        Dim K As Integer
100         K = RandomNumber(1, NpcList(NpcIndex).flags.LanzaSpells)

102     Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, NpcList(NpcIndex).Spells(K))

        
        Exit Sub

NpcLanzaUnSpellSobreNpc_Err:
104     Call RegistrarError(Err.Number, Err.Description, "AI.NpcLanzaUnSpellSobreNpc", Erl)
106     Resume Next
        
End Sub

Private Function PuedeAtacarUser(ByVal targetUserIndex As Integer) As Boolean
    
100     With UserList(targetUserIndex)
            
102         PuedeAtacarUser = (.flags.Muerto = 0 And _
                                .flags.invisible = 0 And _
                                .flags.Inmunidad = 0 And _
                                .flags.Oculto = 0 And _
                                .flags.Mimetizado < e_EstadoMimetismo.FormaBichoSinProteccion And _
                                Not EsGM(targetUserIndex) And _
                                Not .flags.EnConsulta)
                                
        End With

End Function

Private Sub HacerCaminata(ByVal NpcIndex As Integer)
    On Error GoTo Handler
    
    Dim Destino As WorldPos
    Dim Heading As eHeading
    Dim NextTile As WorldPos
    Dim MoveChar As Integer
    Dim PudoMover As Boolean

    With NpcList(NpcIndex)
    
        Destino.Map = .Pos.Map
        Destino.X = .Orig.X + .Caminata(.CaminataActual).Offset.X
        Destino.Y = .Orig.Y + .Caminata(.CaminataActual).Offset.Y

        ' Si todavía no llegó al destino
        If .Pos.X <> Destino.X Or .Pos.Y <> Destino.Y Then
            ' Tratamos de acercarnos (podemos pisar npcs, usuarios o triggers)
            Heading = FindDirectionEAO(.Pos, Destino, .flags.AguaValida, .flags.TierraInvalida = 0, True, True)
            ' Obtengo la posición según el heading
            NextTile = .Pos
            Call HeadtoPos(Heading, NextTile)
            ' Si hay un NPC
            MoveChar = MapData(NextTile.Map, NextTile.X, NextTile.Y).NpcIndex
            If MoveChar Then
                ' Lo movemos hacia un lado
                Call MoveNpcToSide(MoveChar, Heading)
            End If
            ' Si hay un user
            MoveChar = MapData(NextTile.Map, NextTile.X, NextTile.Y).UserIndex
            If MoveChar Then
                ' Si no está muerto o es admin invisible (porque a esos los atraviesa)
                If UserList(MoveChar).flags.AdminInvisible = 0 Or UserList(MoveChar).flags.Muerto = 0 Then
                    ' Lo movemos hacia un lado
                    Call MoveUserToSide(MoveChar, Heading)
                End If
            End If
            ' Movemos al NPC de la caminata
            PudoMover = MoveNPCChar(NpcIndex, Heading)
            ' Si no pudimos moverlo, hacemos como si hubiese llegado a destino... para evitar que se quede atascado
            If Not PudoMover Or Distancia(.Pos, Destino) = 0 Then
                ' Llegamos a destino, ahora esperamos el tiempo necesario para continuar
                .Contadores.IntervaloMovimiento = GetTickCount + .Caminata(.CaminataActual).Espera - .IntervaloMovimiento
                ' Pasamos a la siguiente caminata
                .CaminataActual = .CaminataActual + 1
                ' Si pasamos el último, volvemos al primero
                If .CaminataActual > UBound(.Caminata) Then
                    .CaminataActual = 1
                End If
            End If
        ' Si por alguna razón estamos en el destino, seguimos con la siguiente caminata
        Else
            .CaminataActual = .CaminataActual + 1
            ' Si pasamos el último, volvemos al primero
            If .CaminataActual > UBound(.Caminata) Then
                .CaminataActual = 1
            End If
        End If
    
    End With
    
    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.Description, "AI.HacerCaminata", Erl)
    Resume Next
End Sub

Private Sub MovimientoInvasion(ByVal NpcIndex As Integer)
    On Error GoTo Handler
    
    With NpcList(NpcIndex)
        Dim SpawnBox As tSpawnBox
        SpawnBox = Invasiones(.flags.InvasionIndex).SpawnBoxes(.flags.SpawnBox)
    
        ' Calculamos la distancia a la muralla y generamos una posición de destino
        Dim DistanciaMuralla As Integer, Destino As WorldPos
        Destino = .Pos
        
        If SpawnBox.Heading = eHeading.EAST Or SpawnBox.Heading = eHeading.WEST Then
            DistanciaMuralla = Abs(.Pos.X - SpawnBox.CoordMuralla)
            Destino.X = SpawnBox.CoordMuralla
        Else
            DistanciaMuralla = Abs(.Pos.Y - SpawnBox.CoordMuralla)
            Destino.Y = SpawnBox.CoordMuralla
        End If

        ' Si todavía está lejos de la muralla
        If DistanciaMuralla > 1 Then
            ' Tratamos de acercarnos (sin pisar)
            Dim Heading As eHeading
            Heading = FindDirectionEAO(.Pos, Destino, .flags.AguaValida, .flags.TierraInvalida = 0, True)
            
            ' Nos aseguramos que la posición nueva está dentro del rectángulo válido
            Dim NextTile As WorldPos
            NextTile = .Pos
            Call HeadtoPos(Heading, NextTile)
            
            ' Si la posición nueva queda fuera del rectángulo válido
            If Not InsideRectangle(SpawnBox.LegalBox, NextTile.X, NextTile.Y) Then
                ' Invertimos la dirección de movimiento
                Heading = InvertHeading(Heading)
            End If
            
            ' Movemos el NPC
            Call MoveNPCChar(NpcIndex, Heading)
        
        ' Si está pegado a la muralla
        Else
            ' Chequeamos el intervalo de ataque
            If Not IntervaloPermiteAtacarNPC(NpcIndex) Then
                Exit Sub
            End If
            
            ' Nos aseguramos que mire hacia la muralla
            If .Char.Heading <> SpawnBox.Heading Then
                Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, SpawnBox.Heading)
            End If
            
            ' Sonido de ataque (si tiene)
            If .flags.Snd1 > 0 Then
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
            End If
            
            ' Sonido de impacto
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            ' Dañamos la muralla
            Call HacerDañoMuralla(.flags.InvasionIndex, RandomNumber(.Stats.MinHIT, .Stats.MaxHit))  ' TODO: Defensa de la muralla? No hace falta creo...

        End If
    
    End With

    Exit Sub
    
Handler:
    Call RegistrarError(Err.Number, Err.Description, "AI.MovimientoInvasion", Erl)
    Resume Next
End Sub
