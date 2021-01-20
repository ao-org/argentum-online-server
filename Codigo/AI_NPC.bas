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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo AI_NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'AI de los NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
        
        On Error GoTo IrUsuarioCercano_Err

        Dim tHeading   As Byte
        Dim UI         As Integer
        Dim Pos        As WorldPos
        Dim i          As Long
        Dim comoatacto As Byte

100     With NpcList(NpcIndex)
    
            Dim rangox As Byte
            Dim rangoy As Byte
    
102         If .Distancia <> 0 Then
104             rangox = .Distancia
106             rangoy = .Distancia

            Else
108             rangox = RANGO_VISION_X
110             rangoy = RANGO_VISION_Y

            End If

            'If NpcList(NpcIndex).Target = 0 Then Exit Sub
112         If .flags.Inmovilizado = 1 Then

114             If .flags.LanzaSpells <> 0 Then

116                 For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys

118                     UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
    
120                     If UI > 0 Then
                            'Is it in it's range of vision??
            
122                         If NpcList(NpcIndex).Target = 0 Then Exit Sub
                    
124                         If Abs(UserList(NpcList(NpcIndex).Target).Pos.X - .Pos.X) <= RANGO_VISION_X Then
126                             If Abs(UserList(NpcList(NpcIndex).Target).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                                
128                                 If PuedeAtacarUser(NpcList(NpcIndex).Target) Then
                                                        
130                                     If Distancia(.Pos, UserList(NpcList(NpcIndex).Target).Pos) > 1 Then

132                                         If .flags.LanzaSpells <> 0 Then
134                                             Call NpcLanzaUnSpell(NpcIndex, NpcList(NpcIndex).Target)
                                            End If
                                                                
                                        Else
                                                            
136                                         tHeading = FindDirectionEAO(.Pos, UserList(NpcList(NpcIndex).Target).Pos, NpcList(NpcIndex).flags.AguaValida = 1, NpcList(NpcIndex).flags.TierraInvalida = 0)

138                                         Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)
140                                         Call NpcAtacaUser(NpcIndex, NpcList(NpcIndex).Target, tHeading)

                                        End If
            
                                        Exit Sub

                                    End If
                                                        
                                End If
    
                            End If
    
                        End If

142                 Next i
                
                Else
                
144                 Pos = .Pos
146                 Call HeadtoPos(.Char.Heading, Pos)
                    
148                 UI = MapData(Pos.Map, Pos.X, Pos.Y).UserIndex
                    
150                 If UI > 0 Then

152                     If PuedeAtacarUser(UI) Then
154                         Call NpcAtacaUser(NpcIndex, UI, .Char.Heading)

                            Exit Sub

                        End If
                        
                    End If
                    
                End If

            Else
        
156             If NpcList(NpcIndex).Target <> 0 Then
            
158                 If Abs(UserList(NpcList(NpcIndex).Target).Pos.X - .Pos.X) <= RANGO_VISION_X Then
160                     If Abs(UserList(NpcList(NpcIndex).Target).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
162                         If PuedeAtacarUser(NpcList(NpcIndex).Target) Then

164                             If .flags.LanzaSpells <> 0 And Distancia(.Pos, UserList(NpcList(NpcIndex).Target).Pos) > 1 Then
166                                 Call NpcLanzaUnSpell(NpcIndex, NpcList(NpcIndex).Target)

                                End If

168                             If Distancia(.Pos, UserList(NpcList(NpcIndex).Target).Pos) = 1 Then

170                                 tHeading = FindDirectionEAO(.Pos, UserList(NpcList(NpcIndex).Target).Pos, NpcList(NpcIndex).flags.AguaValida = 1, NpcList(NpcIndex).flags.TierraInvalida = 0)

172                                 Call AnimacionIdle(NpcIndex, True)
174                                 Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)

176                                 If .flags.LanzaSpells <> 0 Then
                                  
178                                     comoatacto = RandomNumber(1, 2)

180                                     If comoatacto = 1 Then
182                                         If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, NpcList(NpcIndex).Target)

                                        Else
184                                         Call NpcAtacaUser(NpcIndex, NpcList(NpcIndex).Target, tHeading)

                                            Exit Sub

                                        End If

                                    Else
186                                     Call NpcAtacaUser(NpcIndex, NpcList(NpcIndex).Target, tHeading)

                                        Exit Sub

                                    End If

                                End If

188                             tHeading = FindDirectionEAO(.Pos, UserList(NpcList(NpcIndex).Target).Pos, NpcList(NpcIndex).flags.AguaValida = 1, NpcList(NpcIndex).flags.TierraInvalida = 0)

190                             Call MoveNPCChar(NpcIndex, tHeading)

                                Exit Sub

                            End If
                        
                        End If

                    End If

                Else
        
192                 For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
194                     UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)

196                     If Abs(UserList(UI).Pos.X - .Pos.X) <= rangox Then
198                         If Abs(UserList(UI).Pos.Y - .Pos.Y) <= rangoy Then
                        
200                             If PuedeAtacarUser(UI) Then

202                                 If .flags.LanzaSpells <> 0 And Distancia(.Pos, UserList(UI).Pos) > 1 Then
204                                     Call NpcLanzaUnSpell(NpcIndex, UI)
                                    End If
                            
206                                 If Distancia(.Pos, UserList(UI).Pos) = 1 Then

208                                     tHeading = FindDirectionEAO(.Pos, UserList(UI).Pos, NpcList(NpcIndex).flags.AguaValida = 1, NpcList(NpcIndex).flags.TierraInvalida = 0)
210                                     Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)
212                                     Call AnimacionIdle(NpcIndex, True)

214                                     If .flags.LanzaSpells <> 0 Then
                                  
216                                         comoatacto = RandomNumber(1, 2)

218                                         If comoatacto = 1 Then
220                                             If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)

                                            Else
222                                             Call NpcAtacaUser(NpcIndex, UI, tHeading)

                                            End If

                                            Exit Sub
                                            
                                        Else
                                        
224                                         Call NpcAtacaUser(NpcIndex, UI, tHeading)

                                            Exit Sub

                                        End If
                                
                                    End If
                            
226                                 tHeading = FindDirectionEAO(.Pos, UserList(UI).Pos, NpcList(NpcIndex).flags.AguaValida = 1, NpcList(NpcIndex).flags.TierraInvalida = 0)

228                                 Call MoveNPCChar(NpcIndex, tHeading)
                           
                                    Exit Sub

                                End If
                        
                            End If

                        End If

230                 Next i

                End If
            
                'Si llega aca es que no había ningún usuario cercano vivo.
                'A bailar. Pablo (ToxicWaste)
            
232             NpcList(NpcIndex).Target = 0

234             If RandomNumber(0, 10) = 0 Then
                
236                 Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

                Else
                
238                 Call AnimacionIdle(NpcIndex, True)

                End If

            End If

        End With

240     Call RestoreOldMovement(NpcIndex)
        
        Exit Sub

IrUsuarioCercano_Err:
242     Call RegistrarError(Err.Number, Err.Description, "AI.IrUsuarioCercano", Erl)

244     Resume Next
        
End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
        
        On Error GoTo SeguirAgresor_Err
        

        Dim tHeading As Byte
        Dim UI       As Integer
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
                            
136                         If UserList(UI).name = .flags.AttackedBy Then
                            
138                             If PuedeAtacarUser(UI) Then

140                                 If .flags.LanzaSpells > 0 Then

142                                     Call AnimacionIdle(NpcIndex, True)
144                                     Call NpcLanzaUnSpell(NpcIndex, UI)
                                        
                                    End If

                                    Exit Sub

                                End If

                            End If
                        
                        End If

                    End If
                
146             Next i

            Else
  
148             For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
150                 UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                    'Is it in it's range of vision??
152                 If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
154                     If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                    
156                         If UserList(UI).name = .flags.AttackedBy Then
                               
158                             If PuedeAtacarUser(UI) Then

160                                 If .flags.LanzaSpells > 0 Then
162                                     Call NpcLanzaUnSpell(NpcIndex, UI)
                                    End If
                                    
164                                 tHeading = FindDirectionEAO(.Pos, UserList(UI).Pos, NpcList(NpcIndex).flags.AguaValida = 1, NpcList(NpcIndex).flags.TierraInvalida = 0)
                                                         
166                                 If Distancia(.Pos, UserList(UI).Pos) = 1 Then
168                                     Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)
170                                     Call AnimacionIdle(NpcIndex, True)
172                                     Call NpcAtacaUser(NpcIndex, UI, tHeading)

                                    Else
                                    
174                                     Call MoveNPCChar(NpcIndex, tHeading)

                                    End If

                                    Exit Sub

                                End If

                            End If
                        
                        End If

                    End If
                
176             Next i

            End If

        End With
    
178     Call RestoreOldMovement(NpcIndex)

        
        Exit Sub

SeguirAgresor_Err:
180     Call RegistrarError(Err.Number, Err.Description, "AI.SeguirAgresor", Erl)
182     Resume Next
        
End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)
        
        On Error GoTo RestoreOldMovement_Err
        

100     With NpcList(NpcIndex)

102         If .MaestroUser = 0 Then
104             .Movement = .flags.OldMovement
106             .Hostile = .flags.OldHostil
108             .flags.AttackedBy = vbNullString
            End If

        End With

        Exit Sub

RestoreOldMovement_Err:
110     Call RegistrarError(Err.Number, Err.Description, "AI.RestoreOldMovement", Erl)
112     Resume Next
        
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

            Else
            
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

            '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
102         Select Case .Movement

                Case TipoAI.ESTATICO
                    Rem  Debug.Print "Es un NPC estatico, no hace nada."
104                 falladesc = " fallo en estatico"
            
106             Case TipoAI.MueveAlAzar
108                 falladesc = " fallo al azar"

110                 If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then Exit Sub
112                 If .NPCtype = eNPCType.GuardiaReal Then
                        'Call GuardiasAI(NpcIndex, False)
114                     Call PersigueCriminal(NpcIndex)

116                 ElseIf .NPCtype = eNPCType.Guardiascaos Then
                        'Call GuardiasAI(NpcIndex, True)
118                     Call PersigueCiudadano(NpcIndex)

                    Else
120                     If RandomNumber(1, 12) = 3 Then
122                         Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                        Else
124                         Call AnimacionIdle(NpcIndex, True)
                        End If
                    End If
            
                    'Va hacia el usuario cercano
126             Case TipoAI.NpcMaloAtacaUsersBuenos
128                 falladesc = " fallo NpcMaloAtacaUsersBuenos"
                    'Debug.Print "atacar "
                    'Call PersigueCiudadano(NpcIndex)
130                 Call IrUsuarioCercano(NpcIndex)
            
                    'Va hacia el usuario que lo ataco(FOLLOW)
132             Case TipoAI.NPCDEFENSA

134                 Call SeguirAgresor(NpcIndex)
            
                    'Persigue criminales
136             Case TipoAI.GuardiasAtacanCriminales
138                 Call PersigueCriminal(NpcIndex)
                    
140             Case TipoAI.GuardiasAtacanCiudadanos
142                 Call PersigueCiudadano(NpcIndex)
                        
144             Case TipoAI.NpcAtacaNpc
146                 Call AiNpcAtacaNpc(NpcIndex)
            
148             Case TipoAI.NpcPathfinding
150                 falladesc = " fallo NpcPathfinding"

152                 If .flags.Inmovilizado = 1 Then Exit Sub

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
            
172                 If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then Exit Sub
174                 Call SeguirAmo(NpcIndex)
            
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
    
100     For Y = NpcList(NpcIndex).Pos.Y - 10 To NpcList(NpcIndex).Pos.Y + 10    'Makes a loop that looks at
102         For X = NpcList(NpcIndex).Pos.X - 10 To NpcList(NpcIndex).Pos.X + 10   '5 tiles in every direction
            
                'Make sure tile is legal
104             If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                
                    'look for a user
106                 If MapData(NpcList(NpcIndex).Pos.Map, X, Y).UserIndex > 0 Then

                        'Move towards user
                        Dim tmpUserIndex As Integer
108                         tmpUserIndex = MapData(NpcList(NpcIndex).Pos.Map, X, Y).UserIndex

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
102         PuedeAtacarUser = (.flags.Muerto = 0 And .flags.invisible = 0 And .flags.Inmunidad = 0 And .flags.Oculto = 0 And Not EsGM(targetUserIndex) And Not .flags.EnConsulta)
        End With

End Function
