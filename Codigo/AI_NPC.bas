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

    Dim nPos        As WorldPos

    Dim headingloop As Byte

    Dim UI          As Integer
    
    With Npclist(NpcIndex)

        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos

            If .flags.Inmovilizado = 0 Or headingloop = .Char.heading Then
                Call HeadtoPos(headingloop, nPos)

                If InMapBounds(nPos.Map, nPos.x, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.x, nPos.Y).UserIndex

                    If UI > 0 Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then

                            '¿ES CRIMINAL?
                            If Not DelCaos Then
                                If Status(UI) <> 1 Then
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)

                                    End If

                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).name And Not .flags.Follow Then
                                    
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)

                                    End If

                                    Exit Sub

                                End If

                            Else
                           
                                If Status(UI) = 1 Then
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)

                                    End If

                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).name And Not .flags.Follow Then
                                      
                                    If NpcAtacaUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)

                                    End If

                                    Exit Sub

                                End If

                            End If

                        End If

                    End If

                End If

            End If  'not inmovil

        Next headingloop

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)

    Dim nPos        As WorldPos

    Dim headingloop As Byte

    Dim UI          As Integer

    Dim NPCI        As Integer

    Dim atacoPJ     As Boolean
    
    atacoPJ = False
    
    With Npclist(NpcIndex)

        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos

            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)

                If InMapBounds(nPos.Map, nPos.x, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.x, nPos.Y).UserIndex
                    NPCI = MapData(nPos.Map, nPos.x, nPos.Y).NpcIndex

                    If UI > 0 And Not atacoPJ Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
                            atacoPJ = True

                            If .flags.LanzaSpells <> 0 Then
                                Call NpcLanzaUnSpell(NpcIndex, UI)

                            End If

                            If NpcAtacaUser(NpcIndex, MapData(nPos.Map, nPos.x, nPos.Y).UserIndex) Then
                                Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)

                            End If

                            Exit Sub

                        End If

                    End If

                End If

            End If  'inmo

        Next headingloop

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)

    Dim nPos        As WorldPos

    Dim headingloop As eHeading

    Dim UI          As Integer
    
    With Npclist(NpcIndex)

        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos

            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)

                If InMapBounds(nPos.Map, nPos.x, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.x, nPos.Y).UserIndex

                    If UI > 0 Then
                        If UserList(UI).name = .flags.AttackedBy Then
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)

                                End If
                                
                                If NpcAtacaUser(NpcIndex, UI) Then
                                    Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)

                                End If

                                Exit Sub

                            End If

                        End If

                    End If

                End If

            End If

        Next headingloop

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)

    Dim theading   As Byte

    Dim UI         As Integer

    Dim SignoNS    As Integer

    Dim SignoEO    As Integer

    Dim i          As Long
    
    Dim comoatacto As Byte

    With Npclist(NpcIndex)
    
        Dim rangox As Byte

        Dim rangoy As Byte
    
        If .Distancia <> 0 Then
            rangox = .Distancia
            rangoy = .Distancia
        Else
            rangox = RANGO_VISION_X
            rangoy = RANGO_VISION_Y

        End If

        'If Npclist(NpcIndex).Target = 0 Then Exit Sub
        If .flags.Inmovilizado = 1 Then

            Select Case .Char.heading

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

                If UI > 0 Then
                    'Is it in it's range of vision??
                    
                    If Npclist(NpcIndex).Target = 0 Then Exit Sub
                
                    If Abs(UserList(Npclist(NpcIndex).Target).Pos.x - .Pos.x) <= RANGO_VISION_X Then
                        If Abs(UserList(Npclist(NpcIndex).Target).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                            
                            If UserList(Npclist(NpcIndex).Target).flags.Muerto = 0 Then
                                If UserList(Npclist(NpcIndex).Target).flags.invisible = 0 Then
                                    If UserList(Npclist(NpcIndex).Target).flags.Oculto = 0 Then
                                        If UserList(Npclist(NpcIndex).Target).flags.AdminPerseguible Then
                                            If .flags.LanzaSpells <> 0 And Distancia(.Pos, UserList(Npclist(NpcIndex).Target).Pos) > 1 Then
                                                Call NpcLanzaUnSpell(NpcIndex, Npclist(NpcIndex).Target)

                                            End If

                                            If Distancia(.Pos, UserList(Npclist(NpcIndex).Target).Pos) = 1 Then
                                                theading = FindDirectionEAO(.Pos, UserList(Npclist(NpcIndex).Target).Pos, (Npclist(NpcIndex).flags.AguaValida))
                                                Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, theading)
                                                Call NpcAtacaUser(NpcIndex, Npclist(NpcIndex).Target)
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

            Next i

        Else
        
            If Npclist(NpcIndex).Target <> 0 Then
            
                If Abs(UserList(Npclist(NpcIndex).Target).Pos.x - .Pos.x) <= RANGO_VISION_X Then
                    If Abs(UserList(Npclist(NpcIndex).Target).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        If UserList(Npclist(NpcIndex).Target).flags.Muerto = 0 And UserList(Npclist(NpcIndex).Target).flags.invisible = 0 And UserList(Npclist(NpcIndex).Target).flags.Oculto = 0 And UserList(Npclist(NpcIndex).Target).flags.AdminPerseguible Then
                            If .flags.LanzaSpells <> 0 And Distancia(.Pos, UserList(Npclist(NpcIndex).Target).Pos) > 1 Then
                                Call NpcLanzaUnSpell(NpcIndex, Npclist(NpcIndex).Target)

                            End If

                            If Distancia(.Pos, UserList(Npclist(NpcIndex).Target).Pos) = 1 Then
                                theading = FindDirectionEAO(.Pos, UserList(Npclist(NpcIndex).Target).Pos, (Npclist(NpcIndex).flags.AguaValida))
                                Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, theading)

                                If .flags.LanzaSpells <> 0 Then
                                  
                                    comoatacto = RandomNumber(1, 2)

                                    If comoatacto = 1 Then
                                        If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, Npclist(NpcIndex).Target)
                                    Else
                                        Call NpcAtacaUser(NpcIndex, Npclist(NpcIndex).Target)
                                        Exit Sub
                                        Exit Sub

                                    End If

                                Else
                                    Call NpcAtacaUser(NpcIndex, Npclist(NpcIndex).Target)
                                    Exit Sub

                                End If

                            End If

                            theading = FindDirectionEAO(.Pos, UserList(Npclist(NpcIndex).Target).Pos, (Npclist(NpcIndex).flags.AguaValida))
                            Call MoveNPCChar(NpcIndex, theading)
                            Exit Sub

                        End If
                        
                    End If

                End If

            Else
        
                For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
                    UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)

                    If Abs(UserList(UI).Pos.x - .Pos.x) <= rangox Then
                        If Abs(UserList(UI).Pos.Y - .Pos.Y) <= rangoy Then
                        
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                                If .flags.LanzaSpells <> 0 And Distancia(.Pos, UserList(UI).Pos) > 1 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)

                                    '  End If
                                End If
                            
                                If Distancia(.Pos, UserList(UI).Pos) = 1 Then
                                    theading = FindDirectionEAO(.Pos, UserList(UI).Pos, (Npclist(NpcIndex).flags.AguaValida))
                                    Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, theading)

                                    If .flags.LanzaSpells <> 0 Then
                                  
                                        comoatacto = RandomNumber(1, 2)

                                        If comoatacto = 1 Then
                                            If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                                        Else
                                            Call NpcAtacaUser(NpcIndex, UI)

                                        End If

                                        Exit Sub
                                    Else
                                        Call NpcAtacaUser(NpcIndex, UI)
                                        Exit Sub

                                    End If
                                
                                End If
                            
                                theading = FindDirectionEAO(.Pos, UserList(UI).Pos, (Npclist(NpcIndex).flags.AguaValida))
                                Call MoveNPCChar(NpcIndex, theading)
                           
                                Exit Sub

                            End If
                        
                        End If

                    End If

                Next i

            End If
            
            'Si llega aca es que no había ningún usuario cercano vivo.
            'A bailar. Pablo (ToxicWaste)
            
            Npclist(NpcIndex).Target = 0

            If RandomNumber(0, 10) = 0 Then
                
                Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

            End If

        End If

    End With

    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)

    Dim theading As Byte

    Dim UI       As Integer
    
    Dim i        As Long
    
    Dim SignoNS  As Integer

    Dim SignoEO  As Integer
    
    With Npclist(NpcIndex)

        If .flags.Inmovilizado = 1 Then

            Select Case .Char.heading

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
                If Abs(UserList(UI).Pos.x - .Pos.x) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.x - .Pos.x) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                            
                        If UserList(UI).name = .flags.AttackedBy Then
                            
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                If .flags.LanzaSpells > 0 Then
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
                If Abs(UserList(UI).Pos.x - .Pos.x) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                    
                        If UserList(UI).name = .flags.AttackedBy Then
                               
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)

                                End If
                                                         
                                If Distancia(.Pos, UserList(UI).Pos) = 1 Then
                                    Call NpcAtacaUser(NpcIndex, UI)

                                End If

                                theading = FindDirectionEAO(.Pos, UserList(UI).Pos, (Npclist(NpcIndex).flags.AguaValida))
                                Call MoveNPCChar(NpcIndex, theading)
                                 
                                Exit Sub

                            End If

                        End If
                        
                    End If

                End If
                
            Next i

        End If

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)

    With Npclist(NpcIndex)
        .Movement = .flags.OldMovement
        .Hostile = .flags.OldHostil
        .flags.AttackedBy = vbNullString

    End With

End Sub

Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)

    Dim UI       As Integer

    Dim theading As Byte

    Dim i        As Long
    
    With Npclist(NpcIndex)

        For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
            UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
            'Is it in it's range of vision??
            If Abs(UserList(UI).Pos.x - .Pos.x) <= RANGO_VISION_X Then
                If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
        
                    If Status(UI) = 1 Or Status(UI) = 3 Then
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                            If .flags.LanzaSpells > 0 Then
                                Call NpcLanzaUnSpell(NpcIndex, UI)

                            End If

                            theading = FindDirectionEAO(.Pos, UserList(UI).Pos, (Npclist(NpcIndex).flags.AguaValida))
                            Call MoveNPCChar(NpcIndex, theading)
                            Exit Sub

                        End If

                    End If
                    
                End If

            End If
            
        Next i

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub CuraResucita(ByVal NpcIndex As Integer)

    Dim UI       As Integer

    Dim theading As Byte

    Dim i        As Long
    
    With Npclist(NpcIndex)

        For i = 1 To ModAreas.ConnGroups(.Pos.Map).CountEntrys
            UI = ModAreas.ConnGroups(.Pos.Map).UserEntrys(i)
                
            'Is it in it's range of vision??
            If Abs(UserList(UI).Pos.x - .Pos.x) <= RANGO_VISION_X Then
                If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then

                    If Not UserList(UI).accion.Particula = ParticulasIndex.Resucitar Then
                        If Status(UI) < 2 Then
                            If UserList(UI).flags.Muerto = 1 Then
                                Call SendData(SendTarget.ToPCArea, UI, PrepareMessageParticleFX(UserList(UI).Char.CharIndex, ParticulasIndex.Resucitar, 250, False))
                                Call SendData(SendTarget.ToPCArea, UI, PrepareMessageBarFx(UserList(UI).Char.CharIndex, 250, Accion_Barra.Resucitar))
                                UserList(UI).accion.AccionPendiente = True
                                UserList(UI).accion.Particula = ParticulasIndex.Resucitar
                                UserList(UI).accion.TipoAccion = Accion_Barra.Resucitar
                            Else

                                If UserList(UI).Stats.MinHp <> UserList(UI).Stats.MaxHp Then
                                    UserList(UI).Stats.MinHp = UserList(UI).Stats.MaxHp
                                    Call WriteUpdateUserStats(UI)
                                    Call SendData(SendTarget.ToPCArea, UI, PrepareMessageParticleFX(UserList(UI).Char.CharIndex, ParticulasIndex.Curar, 100, False))

                                End If

                            End If

                        End If

                    End If

                    Exit Sub
                    ' End If
                    ' End If
                    
                End If

            End If
            
        Next i

    End With

End Sub

Private Sub PersigueCriminal(ByVal NpcIndex As Integer)

    Dim UI       As Integer

    Dim theading As Byte

    Dim i        As Long

    Dim SignoNS  As Integer

    Dim SignoEO  As Integer
    
    With Npclist(NpcIndex)

        If .flags.Inmovilizado = 1 Then

            Select Case .Char.heading

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
                If Abs(UserList(UI).Pos.x - .Pos.x) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.x - .Pos.x) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        If Status(UI) = 0 Or Status(UI) = 2 Then
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                                If .flags.LanzaSpells > 0 Then
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
                If Abs(UserList(UI).Pos.x - .Pos.x) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        If Status(UI) = 0 Or Status(UI) = 2 Then
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)

                                End If

                                If .flags.Inmovilizado = 1 Then Exit Sub
                                theading = FindDirectionEAO(.Pos, UserList(UI).Pos, (Npclist(NpcIndex).flags.AguaValida))
                                Call MoveNPCChar(NpcIndex, theading)
                                Exit Sub

                            End If

                        End If
                        
                    End If

                End If
                
            Next i

        End If

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)

    Dim theading As Byte

    Dim x        As Long

    Dim Y        As Long

    Dim NI       As Integer

    Dim bNoEsta  As Boolean
    
    Dim SignoNS  As Integer

    Dim SignoEO  As Integer
    
    With Npclist(NpcIndex)

        If .flags.Inmovilizado = 1 Then

            Select Case .Char.heading

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
            
            For Y = .Pos.Y To .Pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
                For x = .Pos.x To .Pos.x + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)

                    If x >= MinXBorder And x <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        NI = MapData(.Pos.Map, x, Y).NpcIndex

                        If NI > 0 Then
                            If .TargetNPC = NI Then
                                bNoEsta = True

                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)

                                    If Npclist(NI).NPCtype = DRAGON Then
                                        Npclist(NI).CanAttack = 1
                                        Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)

                                    End If

                                Else

                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)

                                    End If

                                End If

                                Exit Sub

                            End If

                        End If

                    End If

                Next x
            Next Y

        Else

            For Y = .Pos.Y - RANGO_VISION_Y To .Pos.Y + RANGO_VISION_Y
                For x = .Pos.x - RANGO_VISION_Y To .Pos.x + RANGO_VISION_Y

                    If x >= MinXBorder And x <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        NI = MapData(.Pos.Map, x, Y).NpcIndex

                        If NI > 0 Then
                            If .TargetNPC = NI Then
                                bNoEsta = True

                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)

                                    If Npclist(NI).NPCtype = DRAGON Then
                                        Npclist(NI).CanAttack = 1
                                        Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)

                                    End If

                                Else

                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)

                                    End If

                                End If

                                If .flags.Inmovilizado = 1 Then Exit Sub
                                If .TargetNPC = 0 Then Exit Sub
                                theading = FindDirectionEAO(.Pos, Npclist(MapData(.Pos.Map, x, Y).NpcIndex).Pos, (Npclist(NpcIndex).flags.AguaValida))
                                 
                                Call MoveNPCChar(NpcIndex, theading)
                                 
                                Exit Sub

                            End If

                        End If

                    End If

                Next x
            Next Y

        End If
        
        If Not bNoEsta Then
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil

        End If

    End With

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

        End Select

    End With

    Exit Sub

ErrorHandler:
    Call LogError("NPCAI " & Npclist(NpcIndex).name & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist(NpcIndex).Pos.Map & " x:" & Npclist(NpcIndex).Pos.x & " y:" & Npclist(NpcIndex).Pos.Y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).Target & " TargN:" & Npclist(NpcIndex).TargetNPC & falladesc)

    Dim MiNPC As npc

    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)

End Sub

Function UserNear(ByVal NpcIndex As Integer) As Boolean
    '#################################################################
    'Returns True if there is an user adjacent to the npc position.
    '#################################################################
    UserNear = Not Int(Distance(Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.x, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.Y)) > 1

End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean

    '#################################################################
    'Returns true if we have to seek a new path
    '#################################################################
    If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
        ReCalculatePath = True
    ElseIf Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1 Then
        ReCalculatePath = True

    End If

End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
    '#################################################################
    'Coded By Gulfas Morgolock
    'Returns if the npc has arrived to the end of its path
    '#################################################################
    PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght

End Function
 
Function FollowPath(NpcIndex As Integer) As Boolean

    Dim tmpPos   As WorldPos

    Dim theading As Byte
 
    tmpPos.Map = Npclist(NpcIndex).Pos.Map
    tmpPos.x = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).Y
    tmpPos.Y = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).x
 
    theading = FindDirectionEAO(Npclist(NpcIndex).Pos, tmpPos, (Npclist(NpcIndex).flags.AguaValida))
 
    MoveNPCChar NpcIndex, theading
 
    Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.CurPos + 1
 
End Function

Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean

    '#################################################################
    'Coded By Gulfas Morgolock / 11-07-02
    'www.geocities.com/gmorgolock
    'morgolock@speedy.com.ar
    'This function seeks the shortest path from the Npc
    'to the user's location.
    '#################################################################
    Dim Y As Long

    Dim x As Long
    
    For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10    'Makes a loop that looks at
        For x = Npclist(NpcIndex).Pos.x - 10 To Npclist(NpcIndex).Pos.x + 10   '5 tiles in every direction
            
            'Make sure tile is legal
            If x > MinXBorder And x < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                
                'look for a user
                If MapData(Npclist(NpcIndex).Pos.Map, x, Y).UserIndex > 0 Then

                    'Move towards user
                    Dim tmpUserIndex As Integer

                    tmpUserIndex = MapData(Npclist(NpcIndex).Pos.Map, x, Y).UserIndex

                    With UserList(tmpUserIndex)

                        If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible Then
                            'We have to invert the coordinates, this is because
                            'ORE refers to maps in converse way of my pathfinding
                            'routines.
                            Npclist(NpcIndex).PFINFO.Target.x = UserList(tmpUserIndex).Pos.Y
                            Npclist(NpcIndex).PFINFO.Target.Y = UserList(tmpUserIndex).Pos.x 'ops!
                            Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
                            Call SeekPath(NpcIndex)
                            Exit Function

                        End If

                    End With

                End If

            End If

        Next x
    Next Y

End Function

Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    If Not IntervaloPermiteLanzarHechizo(NpcIndex) Then Exit Sub
    If Npclist(NpcIndex).Pos.Map <> UserList(UserIndex).Pos.Map Then Exit Sub
    If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Or UserList(UserIndex).flags.NoMagiaEfeceto = 1 Then Exit Sub
    
    Dim K As Integer

    K = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(K))

    If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = UserIndex
    If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex

End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)

    Dim K As Integer

    K = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Npclist(NpcIndex).Spells(K))

End Sub
