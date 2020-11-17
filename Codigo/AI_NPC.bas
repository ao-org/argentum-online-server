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

Public Const ELEMENTALFUEGO  As Integer = 93

Public Const ELEMENTALTIERRA As Integer = 94

Public Const ELEMENTALAGUA   As Integer = 92

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

Private Sub GuardiasAI(ByVal NpcIndex As Integer, ByVal DelCaos As Boolean)
        
        On Error GoTo GuardiasAI_Err
        

        Dim nPos        As WorldPos

        Dim headingloop As Byte

        Dim UI          As Integer
    
100     With Npclist(NpcIndex)

102         For headingloop = eHeading.NORTH To eHeading.WEST
104             nPos = .Pos

106             If .flags.Inmovilizado = 0 Or headingloop = .Char.heading Then
108                 Call HeadtoPos(headingloop, nPos)

110                 If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
112                     UI = MapData(nPos.Map, nPos.X, nPos.Y).Userindex

114                     If UI > 0 Then
116                         If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then

                                '¿ES CRIMINAL?
118                             If Not DelCaos Then
120                                 If Status(UI) <> 1 Then
122                                     If NpcAtacaUser(NpcIndex, UI) Then
124                                         Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)

                                        End If

                                        Exit Sub
126                                 ElseIf .flags.AttackedBy = UserList(UI).name And Not .flags.Follow Then
                                    
128                                     If NpcAtacaUser(NpcIndex, UI) Then
130                                         Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)

                                        End If

                                        Exit Sub

                                    End If

                                Else
                           
132                                 If Status(UI) = 1 Then
134                                     If NpcAtacaUser(NpcIndex, UI) Then
136                                         Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)

                                        End If

                                        Exit Sub
138                                 ElseIf .flags.AttackedBy = UserList(UI).name And Not .flags.Follow Then
                                      
140                                     If NpcAtacaUser(NpcIndex, UI) Then
142                                         Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)

                                        End If

                                        Exit Sub

                                    End If

                                End If

                            End If

                        End If

                    End If

                End If  'not inmovil

144         Next headingloop

        End With
    
146     Call RestoreOldMovement(NpcIndex)

        
        Exit Sub

GuardiasAI_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.GuardiasAI", Erl)
        Resume Next
        
End Sub

Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)
        
        On Error GoTo HostilMalvadoAI_Err
        

        Dim nPos        As WorldPos

        Dim headingloop As Byte

        Dim UI          As Integer

        Dim NPCI        As Integer

        Dim atacoPJ     As Boolean
    
100     atacoPJ = False
    
102     With Npclist(NpcIndex)

104         For headingloop = eHeading.NORTH To eHeading.WEST
106             nPos = .Pos

108             If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
110                 Call HeadtoPos(headingloop, nPos)

112                 If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
114                     UI = MapData(nPos.Map, nPos.X, nPos.Y).Userindex
116                     NPCI = MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex

118                     If UI > 0 And Not atacoPJ Then
120                         If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
122                             atacoPJ = True

124                             If .flags.LanzaSpells <> 0 Then
126                                 Call NpcLanzaUnSpell(NpcIndex, UI)

                                End If

128                             If NpcAtacaUser(NpcIndex, MapData(nPos.Map, nPos.X, nPos.Y).Userindex) Then
130                                 Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)

                                End If

                                Exit Sub

                            End If

                        End If

                    End If

                End If  'inmo

132         Next headingloop

        End With
    
134     Call RestoreOldMovement(NpcIndex)

        
        Exit Sub

HostilMalvadoAI_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.HostilMalvadoAI", Erl)
        Resume Next
        
End Sub

Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)
        
        On Error GoTo HostilBuenoAI_Err
        

        Dim nPos        As WorldPos

        Dim headingloop As eHeading

        Dim UI          As Integer
    
100     With Npclist(NpcIndex)

102         For headingloop = eHeading.NORTH To eHeading.WEST
104             nPos = .Pos

106             If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
108                 Call HeadtoPos(headingloop, nPos)

110                 If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
112                     UI = MapData(nPos.Map, nPos.X, nPos.Y).Userindex

114                     If UI > 0 Then
116                         If UserList(UI).name = .flags.AttackedBy Then
118                             If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
120                                 If .flags.LanzaSpells > 0 Then
122                                     Call NpcLanzaUnSpell(NpcIndex, UI)

                                    End If
                                
124                                 If NpcAtacaUser(NpcIndex, UI) Then
126                                     Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)

                                    End If

                                    Exit Sub

                                End If

                            End If

                        End If

                    End If

                End If

128         Next headingloop

        End With
    
130     Call RestoreOldMovement(NpcIndex)

        
        Exit Sub

HostilBuenoAI_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.HostilBuenoAI", Erl)
        Resume Next
        
End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
        
        On Error GoTo IrUsuarioCercano_Err
        

        Dim tHeading   As Byte

        Dim UI         As Integer

        Dim Pos        As WorldPos

        Dim i          As Long
    
        Dim comoatacto As Byte

100     With Npclist(NpcIndex)
    
            Dim rangox As Byte

            Dim rangoy As Byte
    
102         If .Distancia <> 0 Then
104             rangox = .Distancia
106             rangoy = .Distancia
            Else
108             rangox = RANGO_VISION_X
110             rangoy = RANGO_VISION_Y

            End If

            'If Npclist(NpcIndex).Target = 0 Then Exit Sub
112         If .flags.Inmovilizado = 1 Then
                If .flags.LanzaSpells <> 0 Then
                    For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys

140                 UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)

142                 If UI > 0 Then
                        'Is it in it's range of vision??
                    
144                     If Npclist(NpcIndex).Target = 0 Then Exit Sub
                
146                     If Abs(UserList(Npclist(NpcIndex).Target).Pos.X - .Pos.X) <= RANGO_VISION_X Then
148                         If Abs(UserList(Npclist(NpcIndex).Target).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                            
150                             If UserList(Npclist(NpcIndex).Target).flags.Muerto = 0 Then
152                                 If UserList(Npclist(NpcIndex).Target).flags.invisible = 0 Then
154                                     If UserList(Npclist(NpcIndex).Target).flags.Oculto = 0 Then
156                                         If UserList(Npclist(NpcIndex).Target).flags.AdminPerseguible Then
158                                             If .flags.LanzaSpells <> 0 And Distancia(.Pos, UserList(Npclist(NpcIndex).Target).Pos) > 1 Then
160                                                 Call NpcLanzaUnSpell(NpcIndex, Npclist(NpcIndex).Target)

                                                End If

162                                             If Distancia(.Pos, UserList(Npclist(NpcIndex).Target).Pos) = 1 Then
164                                                 tHeading = FindDirectionEAO(.Pos, UserList(Npclist(NpcIndex).Target).Pos, (Npclist(NpcIndex).flags.AguaValida))
166                                                 Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)
168                                                 Call NpcAtacaUser(NpcIndex, Npclist(NpcIndex).Target)
                                                    Exit Sub

                                                End If

                                                Exit Sub

                                            End If

                                        End If

                                    End If

                                End If
                            
                            End If

                        End If

                    End If

170             Next i
                
                Else
                    Pos = .Pos
                    Call HeadtoPos(.Char.heading, Pos)
                    
                    UI = MapData(Pos.Map, Pos.X, Pos.Y).Userindex
                    
                    If UI > 0 Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
171                         Call NpcAtacaUser(NpcIndex, UI)
                            Exit Sub
                        End If
                    End If
                End If

            Else
        
172             If Npclist(NpcIndex).Target <> 0 Then
            
174                 If Abs(UserList(Npclist(NpcIndex).Target).Pos.X - .Pos.X) <= RANGO_VISION_X Then
176                     If Abs(UserList(Npclist(NpcIndex).Target).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
178                         If UserList(Npclist(NpcIndex).Target).flags.Muerto = 0 And UserList(Npclist(NpcIndex).Target).flags.invisible = 0 And UserList(Npclist(NpcIndex).Target).flags.Oculto = 0 And UserList(Npclist(NpcIndex).Target).flags.AdminPerseguible Then
180                             If .flags.LanzaSpells <> 0 And Distancia(.Pos, UserList(Npclist(NpcIndex).Target).Pos) > 1 Then
182                                 Call NpcLanzaUnSpell(NpcIndex, Npclist(NpcIndex).Target)

                                End If

184                             If Distancia(.Pos, UserList(Npclist(NpcIndex).Target).Pos) = 1 Then
186                                 tHeading = FindDirectionEAO(.Pos, UserList(Npclist(NpcIndex).Target).Pos, (Npclist(NpcIndex).flags.AguaValida))
188                                 Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)

190                                 If .flags.LanzaSpells <> 0 Then
                                  
192                                     comoatacto = RandomNumber(1, 2)

194                                     If comoatacto = 1 Then
196                                         If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, Npclist(NpcIndex).Target)
                                        Else
198                                         Call NpcAtacaUser(NpcIndex, Npclist(NpcIndex).Target)
                                            Exit Sub
                                            Exit Sub

                                        End If

                                    Else
200                                     Call NpcAtacaUser(NpcIndex, Npclist(NpcIndex).Target)
                                        Exit Sub

                                    End If

                                End If

202                             tHeading = FindDirectionEAO(.Pos, UserList(Npclist(NpcIndex).Target).Pos, (Npclist(NpcIndex).flags.AguaValida))
204                             Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub

                            End If
                        
                        End If

                    End If

                Else
        
206                 For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
208                     UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)

210                     If Abs(UserList(UI).Pos.X - .Pos.X) <= rangox Then
212                         If Abs(UserList(UI).Pos.Y - .Pos.Y) <= rangoy Then
                        
214                             If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
216                                 If .flags.LanzaSpells <> 0 And Distancia(.Pos, UserList(UI).Pos) > 1 Then
218                                     Call NpcLanzaUnSpell(NpcIndex, UI)

                                        '  End If
                                    End If
                            
220                                 If Distancia(.Pos, UserList(UI).Pos) = 1 Then
222                                     tHeading = FindDirectionEAO(.Pos, UserList(UI).Pos, (Npclist(NpcIndex).flags.AguaValida))
224                                     Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)

226                                     If .flags.LanzaSpells <> 0 Then
                                  
228                                         comoatacto = RandomNumber(1, 2)

230                                         If comoatacto = 1 Then
232                                             If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                                            Else
234                                             Call NpcAtacaUser(NpcIndex, UI)

                                            End If

                                            Exit Sub
                                        Else
236                                         Call NpcAtacaUser(NpcIndex, UI)
                                            Exit Sub

                                        End If
                                
                                    End If
                            
238                                 tHeading = FindDirectionEAO(.Pos, UserList(UI).Pos, (Npclist(NpcIndex).flags.AguaValida))
240                                 Call MoveNPCChar(NpcIndex, tHeading)
                           
                                    Exit Sub

                                End If
                        
                            End If

                        End If

242                 Next i

                End If
            
                'Si llega aca es que no había ningún usuario cercano vivo.
                'A bailar. Pablo (ToxicWaste)
            
244             Npclist(NpcIndex).Target = 0

246             If RandomNumber(0, 10) = 0 Then
                
248                 Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

                End If

            End If

        End With

250     Call RestoreOldMovement(NpcIndex)

        
        Exit Sub

IrUsuarioCercano_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.IrUsuarioCercano", Erl)
        Resume Next
        
End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
        
        On Error GoTo SeguirAgresor_Err
        

        Dim tHeading As Byte

        Dim UI       As Integer
    
        Dim i        As Long
    
        Dim SignoNS  As Integer

        Dim SignoEO  As Integer
    
100     With Npclist(NpcIndex)

102         If .flags.Inmovilizado = 1 Then

104             Select Case .Char.heading

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
                            
138                             If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
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
                    
154                         If UserList(UI).name = .flags.AttackedBy Then
                               
156                             If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
158                                 If .flags.LanzaSpells > 0 Then
160                                     Call NpcLanzaUnSpell(NpcIndex, UI)

                                    End If
                                                         
162                                 If Distancia(.Pos, UserList(UI).Pos) = 1 Then
164                                     Call NpcAtacaUser(NpcIndex, UI)

                                    End If

166                                 tHeading = FindDirectionEAO(.Pos, UserList(UI).Pos, (Npclist(NpcIndex).flags.AguaValida))
168                                 Call MoveNPCChar(NpcIndex, tHeading)
                                 
                                    Exit Sub

                                End If

                            End If
                        
                        End If

                    End If
                
170             Next i

            End If

        End With
    
172     Call RestoreOldMovement(NpcIndex)

        
        Exit Sub

SeguirAgresor_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.SeguirAgresor", Erl)
        Resume Next
        
End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)
        
        On Error GoTo RestoreOldMovement_Err
        

100     With Npclist(NpcIndex)
102         .Movement = .flags.OldMovement
104         .Hostile = .flags.OldHostil
106         .flags.AttackedBy = vbNullString

        End With

        
        Exit Sub

RestoreOldMovement_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.RestoreOldMovement", Erl)
        Resume Next
        
End Sub

Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)
        
        On Error GoTo PersigueCiudadano_Err
        

        Dim UI       As Integer

        Dim tHeading As Byte

        Dim i        As Long
    
100     With Npclist(NpcIndex)

102         For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
104             UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
                'Is it in it's range of vision??
106             If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
108                 If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
        
110                     If Status(UI) = 1 Or Status(UI) = 3 Then
                        
112                         If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
114                             If .flags.LanzaSpells > 0 Then
116                                 Call NpcLanzaUnSpell(NpcIndex, UI)

                                End If

118                             tHeading = FindDirectionEAO(.Pos, UserList(UI).Pos, (Npclist(NpcIndex).flags.AguaValida))
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
        Call RegistrarError(Err.Number, Err.description, "AI.PersigueCiudadano", Erl)
        Resume Next
        
End Sub

Private Sub CuraResucita(ByVal NpcIndex As Integer)
        
        On Error GoTo CuraResucita_Err
        

        Dim UI       As Integer

        Dim tHeading As Byte

        Dim i        As Long
    
100     With Npclist(NpcIndex)

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
        Call RegistrarError(Err.Number, Err.description, "AI.CuraResucita", Erl)
        Resume Next
        
End Sub

Private Sub PersigueCriminal(ByVal NpcIndex As Integer)
        
        On Error GoTo PersigueCriminal_Err
        

        Dim UI       As Integer

        Dim tHeading As Byte

        Dim i        As Long

        Dim SignoNS  As Integer

        Dim SignoEO  As Integer
    
100     With Npclist(NpcIndex)

102         If .flags.Inmovilizado = 1 Then

104             Select Case .Char.heading

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
138                             If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
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
156                             If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
158                                 If .flags.LanzaSpells > 0 Then
160                                     Call NpcLanzaUnSpell(NpcIndex, UI)

                                    End If

162                                 If .flags.Inmovilizado = 1 Then Exit Sub
164                                 tHeading = FindDirectionEAO(.Pos, UserList(UI).Pos, (Npclist(NpcIndex).flags.AguaValida))
166                                 Call MoveNPCChar(NpcIndex, tHeading)
                                    Exit Sub

                                End If

                            End If
                        
                        End If

                    End If
                
168             Next i

            End If

        End With
    
170     Call RestoreOldMovement(NpcIndex)

        
        Exit Sub

PersigueCriminal_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.PersigueCriminal", Erl)
        Resume Next
        
End Sub

Private Sub SeguirAmo(ByVal NpcIndex As Integer)

        On Error GoTo SeguirAmo_Err

        Dim tHeading As Byte
        Dim UI As Integer
    
100     With Npclist(NpcIndex)

102         If .Target = 0 And .TargetNPC = 0 Then
104             UI = .MaestroUser
            
                'Is it in it's range of vision??
106             If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then

108                 If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then

110                     If UserList(UI).flags.Muerto = 0 _
                                And UserList(UI).flags.invisible = 0 _
                                And UserList(UI).flags.Oculto = 0 _
                                And Distancia(.Pos, UserList(UI).Pos) > 3 Then
                                
112                         tHeading = FindDirection(.Pos, UserList(UI).Pos)

114                         Call MoveNPCChar(NpcIndex, tHeading)

                            Exit Sub
                            
                        End If
                        
                    End If
                    
                End If
                
            End If
            
        End With
    
116     Call RestoreOldMovement(NpcIndex)

        Exit Sub

SeguirAmo_Err:
118     Call RegistrarError(Err.Number, Err.description, "AI.SeguirAmo", Erl)
120     Resume Next

End Sub


Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)
        
        On Error GoTo AiNpcAtacaNpc_Err
        

        Dim tHeading As Byte

        Dim X        As Long

        Dim Y        As Long

        Dim NI       As Integer

        Dim bNoEsta  As Boolean
    
        Dim SignoNS  As Integer

        Dim SignoEO  As Integer
    
100     With Npclist(NpcIndex)

102         If .flags.Inmovilizado = 1 Then

104             Select Case .Char.heading

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

146                                     If Npclist(NI).NPCtype = DRAGON Then
148                                         Npclist(NI).CanAttack = 1
150                                         Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)

                                        End If

                                    Else

                                        'aca verificamosss la distancia de ataque
152                                     If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
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
166                         NI = MapData(.Pos.Map, X, Y).NpcIndex

168                         If NI > 0 Then
170                             If .TargetNPC = NI Then
172                                 bNoEsta = True

174                                 If .Numero = ELEMENTALFUEGO Then
176                                     Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)

178                                     If Npclist(NI).NPCtype = DRAGON Then
180                                         Npclist(NI).CanAttack = 1
182                                         Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)

                                        End If

                                    Else

                                        'aca verificamosss la distancia de ataque
184                                     If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
186                                         Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)

                                        End If

                                    End If

188                                 If .flags.Inmovilizado = 1 Then Exit Sub
190                                 If .TargetNPC = 0 Then Exit Sub
192                                 tHeading = FindDirectionEAO(.Pos, Npclist(MapData(.Pos.Map, X, Y).NpcIndex).Pos, (Npclist(NpcIndex).flags.AguaValida))
                                 
194                                 Call MoveNPCChar(NpcIndex, tHeading)
                                 
                                    Exit Sub

                                End If

                            End If

                        End If

196                 Next X
198             Next Y

            End If
        
200         If Not bNoEsta Then
202             .Movement = .flags.OldMovement
204             .Hostile = .flags.OldHostil

            End If

        End With

        
        Exit Sub

AiNpcAtacaNpc_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.AiNpcAtacaNpc", Erl)
        Resume Next
        
End Sub

Sub NPCAI(ByVal NpcIndex As Integer)

    On Error GoTo ErrorHandler

    Dim falladesc As String

    With Npclist(NpcIndex)
        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        ' If .MaestroUser = 0 Then
        'Busca a alguien para atacar
        '¿Es un guardia?
        '    If .NPCtype = eNPCType.GuardiaReal Then
        'Call GuardiasAI(NpcIndex, False)
                
        ' ElseIf .NPCtype = eNPCType.Guardiascaos Then
        '     Call GuardiasAI(NpcIndex, True)
        ' ElseIf .Hostile And .Stats.Alineacion <> 0 Then
        '      Call HostilMalvadoAI(NpcIndex)
        '  ElseIf .Hostile And .Stats.Alineacion = 0 Then
        '      Call HostilBuenoAI(NpcIndex)
                
        '   ElseIf .NPCtype = Revividor Then
        '  Call CuraResucita(NpcIndex)
        '   End If
        ' Else
        'Evitamos que ataque a su amo, a menos
        'que el amo lo ataque.
        'Call HostilBuenoAI(NpcIndex)
        '  End If
        
        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
        Select Case .Movement

            Case TipoAI.ESTATICO
                Rem  Debug.Print "Es un NPC estatico, no hace nada."
                falladesc = " fallo en estatico"
            
            Case TipoAI.MueveAlAzar
                falladesc = " fallo al azar"

                If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then Exit Sub
                If .NPCtype = eNPCType.GuardiaReal Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

                    End If

                    Call PersigueCriminal(NpcIndex)
                ElseIf .NPCtype = eNPCType.Guardiascaos Then

                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

                    End If
                    
                    Call PersigueCiudadano(NpcIndex)
                Else

                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

                    End If

                End If
            
                'Va hacia el usuario cercano
            Case TipoAI.NpcMaloAtacaUsersBuenos
                falladesc = " fallo NpcMaloAtacaUsersBuenos"
                'Debug.Print "atacar "
                'Call PersigueCiudadano(NpcIndex)
                Call IrUsuarioCercano(NpcIndex)
            
                'Va hacia el usuario que lo ataco(FOLLOW)
            Case TipoAI.NPCDEFENSA

                Call SeguirAgresor(NpcIndex)
            
                'Persigue criminales
            Case TipoAI.GuardiasAtacanCriminales
                Call PersigueCriminal(NpcIndex)
                    
            Case TipoAI.GuardiasAtacanCiudadanos
                Call PersigueCiudadano(NpcIndex)
                        
            Case TipoAI.NpcAtacaNpc
                Call AiNpcAtacaNpc(NpcIndex)
            
            Case TipoAI.NpcPathfinding
                falladesc = " fallo NpcPathfinding"

                If .flags.Inmovilizado = 1 Then Exit Sub
                If ReCalculatePath(NpcIndex) Then
                    Call PathFindingAI(NpcIndex)

                    'Existe el camino?
                    If .PFINFO.NoPath Then 'Si no existe nos movemos al azar
                        'Move randomly
                        Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))

                    End If

                Else

                    If Not PathEnd(NpcIndex) Then
                        Call FollowPath(NpcIndex)
                    Else
                        .PFINFO.PathLenght = 0

                    End If

                End If
            
            Case TipoAI.SigueAmo
                If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then Exit Sub
                Call SeguirAmo(NpcIndex)
                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                End If
            
        End Select

    End With

    Exit Sub

ErrorHandler:
    Call LogError("NPCAI " & Npclist(NpcIndex).name & " " & Npclist(NpcIndex).MaestroNPC & " mapa:" & Npclist(NpcIndex).Pos.Map & " x:" & Npclist(NpcIndex).Pos.X & " y:" & Npclist(NpcIndex).Pos.Y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).Target & " TargN:" & Npclist(NpcIndex).TargetNPC & falladesc)

    Dim MiNPC As npc

    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)

End Sub

Function UserNear(ByVal NpcIndex As Integer) As Boolean
        '#################################################################
        'Returns True if there is an user adjacent to the npc position.
        '#################################################################
        
        On Error GoTo UserNear_Err
        
100     UserNear = Not Int(Distance(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.X, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.Y)) > 1

        
        Exit Function

UserNear_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.UserNear", Erl)
        Resume Next
        
End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean
        
        On Error GoTo ReCalculatePath_Err
        

        '#################################################################
        'Returns true if we have to seek a new path
        '#################################################################
100     If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
102         ReCalculatePath = True
104     ElseIf Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1 Then
106         ReCalculatePath = True

        End If

        
        Exit Function

ReCalculatePath_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.ReCalculatePath", Erl)
        Resume Next
        
End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
        '#################################################################
        'Coded By Gulfas Morgolock
        'Returns if the npc has arrived to the end of its path
        '#################################################################
        
        On Error GoTo PathEnd_Err
        
100     PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght

        
        Exit Function

PathEnd_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.PathEnd", Erl)
        Resume Next
        
End Function
 
Function FollowPath(NpcIndex As Integer) As Boolean
        
        On Error GoTo FollowPath_Err
        

        Dim tmpPos   As WorldPos

        Dim tHeading As Byte
 
100     tmpPos.Map = Npclist(NpcIndex).Pos.Map
102     tmpPos.X = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).Y
104     tmpPos.Y = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).X
 
106     tHeading = FindDirectionEAO(Npclist(NpcIndex).Pos, tmpPos, (Npclist(NpcIndex).flags.AguaValida))
 
108     MoveNPCChar NpcIndex, tHeading
 
110     Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.CurPos + 1
 
        
        Exit Function

FollowPath_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.FollowPath", Erl)
        Resume Next
        
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
    
100     For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10    'Makes a loop that looks at
102         For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10   '5 tiles in every direction
            
                'Make sure tile is legal
104             If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                
                    'look for a user
106                 If MapData(Npclist(NpcIndex).Pos.Map, X, Y).Userindex > 0 Then

                        'Move towards user
                        Dim tmpUserIndex As Integer

108                     tmpUserIndex = MapData(Npclist(NpcIndex).Pos.Map, X, Y).Userindex

110                     With UserList(tmpUserIndex)

112                         If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible Then
                                'We have to invert the coordinates, this is because
                                'ORE refers to maps in converse way of my pathfinding
                                'routines.
114                             Npclist(NpcIndex).PFINFO.Target.X = UserList(tmpUserIndex).Pos.Y
116                             Npclist(NpcIndex).PFINFO.Target.Y = UserList(tmpUserIndex).Pos.X 'ops!
118                             Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
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
        Call RegistrarError(Err.Number, Err.description, "AI.PathFindingAI", Erl)
        Resume Next
        
End Function

Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal Userindex As Integer)
        
        On Error GoTo NpcLanzaUnSpell_Err
        

100     If Not IntervaloPermiteLanzarHechizo(NpcIndex) Then Exit Sub
102     If Npclist(NpcIndex).Pos.Map <> UserList(Userindex).Pos.Map Then Exit Sub
104     If UserList(Userindex).flags.invisible = 1 Or UserList(Userindex).flags.Oculto = 1 Or UserList(Userindex).flags.NoMagiaEfeceto = 1 Then Exit Sub
    
        Dim K As Integer

106     K = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
108     Call NpcLanzaSpellSobreUser(NpcIndex, Userindex, Npclist(NpcIndex).Spells(K))

110     If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = Userindex
112     If UserList(Userindex).flags.AtacadoPorNpc = 0 And UserList(Userindex).flags.AtacadoPorUser = 0 Then UserList(Userindex).flags.AtacadoPorNpc = NpcIndex

        
        Exit Sub

NpcLanzaUnSpell_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.NpcLanzaUnSpell", Erl)
        Resume Next
        
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)
        
        On Error GoTo NpcLanzaUnSpellSobreNpc_Err
        

        Dim K As Integer

100     K = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
102     Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Npclist(NpcIndex).Spells(K))

        
        Exit Sub

NpcLanzaUnSpellSobreNpc_Err:
        Call RegistrarError(Err.Number, Err.description, "AI.NpcLanzaUnSpellSobreNpc", Erl)
        Resume Next
        
End Sub
