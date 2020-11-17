Attribute VB_Name = "modHechizos"
'Argentum Online 0.11.6
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'for more information about ORE please visit http://www.baronsoft.com/C
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public Const HELEMENTAL_FUEGO  As Integer = 26

Public Const HELEMENTAL_TIERRA As Integer = 28

Public Const SUPERANILLO       As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal Userindex As Integer, ByVal Spell As Integer, Optional ByVal IgnoreVisibilityCheck As Boolean = False)
    'Guardia caos
        
    On Error GoTo NpcLanzaSpellSobreUser_Err
        
    With UserList(Userindex)
        
        '�NPC puede ver a trav�s de la invisibilidad?
        If Not IgnoreVisibilityCheck Then
            If .flags.invisible = 1 Or .flags.Oculto = 1 Then Exit Sub
        End If

        'Npclist(NpcIndex).CanAttack = 0
        Dim da�o As Integer

        If Hechizos(Spell).SubeHP = 1 Then

            da�o = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

            .Stats.MinHp = .Stats.MinHp + da�o

            If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
    
            Call WriteConsoleMsg(Userindex, Npclist(NpcIndex).name & " te ha quitado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteUpdateHP(Userindex)
            Call SubirSkill(Userindex, Resistencia)

        ElseIf Hechizos(Spell).SubeHP = 2 Then
        
            da�o = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
        
            If .Invent.CascoEqpObjIndex > 0 Then
                da�o = da�o - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)

            End If
        
            If .Invent.AnilloEqpObjIndex > 0 Then
                da�o = da�o - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)

            End If
        
            If da�o < 0 Then da�o = 0
        
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
                
            If Hechizos(Spell).Particle > 0 Then '�Envio Particula?

                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(.Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))
          
            End If

            .Stats.MinHp = .Stats.MinHp - da�o
        
            Call WriteConsoleMsg(Userindex, Npclist(NpcIndex).name & " te ha quitado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteUpdateHP(Userindex)
            Call SubirSkill(Userindex, Resistencia)
        
            'Muere
            If .Stats.MinHp < 1 Then
                .Stats.MinHp = 0
                'If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                'restarCriminalidad (UserIndex)
                ' End If
                Call UserDie(Userindex)
                '[Barrin 1-12-03]
                ' If Npclist(NpcIndex).MaestroUser > 0 Then
                'Store it!
                'Call Statistics.StoreFrag(Npclist(NpcIndex).MaestroUser, UserIndex)
                
                'Call ContarMuerte(UserIndex, Npclist(NpcIndex).MaestroUser)
                ' Call ActStats(UserIndex, Npclist(NpcIndex).MaestroUser)
                '  End If
                '[/Barrin]
            End If
    
        ElseIf Hechizos(Spell).Paraliza = 1 Then

            If .flags.Paralizado = 0 Then
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

                .flags.Paralizado = 1
                .Counters.Paralisis = Hechizos(Spell).Duration / 2
          
                Call WriteParalizeOK(Userindex)
                Call WritePosUpdate(Userindex)

            End If

        ElseIf Hechizos(Spell).incinera = 1 Then
            Debug.Print "incinerar"

            If .flags.Incinerado = 0 Then
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.X, .Pos.Y))

                If Hechizos(Spell).Particle > 0 Then '�Envio Particula?
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(.Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))

                End If

                .flags.Incinerado = 1
                Call WriteConsoleMsg(Userindex, "Has sido incinerado por " & Npclist(NpcIndex).name & ".", FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If
    
    End With

    Exit Sub

NpcLanzaSpellSobreUser_Err:
    Call RegistrarError(Err.Number, Err.description, "modHechizos.NpcLanzaSpellSobreUser", Erl)

    Resume Next
        
End Sub

Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
        'solo hechizos ofensivos!
        
        On Error GoTo NpcLanzaSpellSobreNpc_Err
        

100     If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub

102     Npclist(NpcIndex).CanAttack = 0

        Dim da�o As Integer

104     If Hechizos(Spell).SubeHP = 2 Then
    
106         da�o = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
108         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, Npclist(TargetNPC).Pos.X, Npclist(TargetNPC).Pos.Y))
110         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
        
112         Npclist(TargetNPC).Stats.MinHp = Npclist(TargetNPC).Stats.MinHp - da�o
        
            'Call CalcularDarExp(Npclist(NpcIndex).MaestroUser, TargetNPC, da�o)
        
            'Muere
114         If Npclist(TargetNPC).Stats.MinHp < 1 Then
116             Npclist(TargetNPC).Stats.MinHp = 0
                ' If Npclist(NpcIndex).MaestroUser > 0 Then
                '  Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
                '  Else
118             Call MuereNpc(TargetNPC, 0)

                '  End If
            End If
    
        End If
    
        
        Exit Sub

NpcLanzaSpellSobreNpc_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.NpcLanzaSpellSobreNpc", Erl)
        Resume Next
        
End Sub

Function TieneHechizo(ByVal i As Integer, ByVal Userindex As Integer) As Boolean

    On Error GoTo errHandler
    
    Dim j As Integer

    For j = 1 To MAXUSERHECHIZOS

        If UserList(Userindex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function

        End If

    Next

    Exit Function
errHandler:

End Function

Sub AgregarHechizo(ByVal Userindex As Integer, ByVal Slot As Integer)
        
        On Error GoTo AgregarHechizo_Err
        

        Dim hIndex As Integer

        Dim j      As Integer

100     hIndex = ObjData(UserList(Userindex).Invent.Object(Slot).ObjIndex).HechizoIndex

102     If Not TieneHechizo(hIndex, Userindex) Then

            'Buscamos un slot vacio
104         For j = 1 To MAXUSERHECHIZOS

106             If UserList(Userindex).Stats.UserHechizos(j) = 0 Then Exit For
108         Next j
        
110         If UserList(Userindex).Stats.UserHechizos(j) <> 0 Then
112             Call WriteConsoleMsg(Userindex, "No tenes espacio para mas hechizos.", FontTypeNames.FONTTYPE_INFO)
            Else
114             UserList(Userindex).Stats.UserHechizos(j) = hIndex
116             Call UpdateUserHechizos(False, Userindex, CByte(j))
                'Quitamos del inv el item
118             Call QuitarUserInvItem(Userindex, CByte(Slot), 1)

            End If

        Else
120         Call WriteConsoleMsg(Userindex, "Ya tenes ese hechizo.", FontTypeNames.FONTTYPE_INFO)

        End If

        
        Exit Sub

AgregarHechizo_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.AgregarHechizo", Erl)
        Resume Next
        
End Sub
            
Sub DecirPalabrasMagicas(ByVal Hechizo As Byte, ByVal Userindex As Integer)

    On Error Resume Next

    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead("PMAG*" & Hechizo, UserList(Userindex).Char.CharIndex, vbCyan))
    Exit Sub

End Sub

Function PuedeLanzar(ByVal Userindex As Integer, ByVal HechizoIndex As Integer, Optional ByVal Slot As Integer = 0) As Boolean
        
        On Error GoTo PuedeLanzar_Err
        

100     If UserList(Userindex).flags.Muerto = 0 Then

            Dim wp2 As WorldPos

102         wp2.Map = UserList(Userindex).flags.TargetMap
104         wp2.X = UserList(Userindex).flags.TargetX
106         wp2.Y = UserList(Userindex).flags.TargetY
    
108         If Hechizos(HechizoIndex).NecesitaObj > 0 Then
110             If TieneObjEnInv(Userindex, Hechizos(HechizoIndex).NecesitaObj, Hechizos(HechizoIndex).NecesitaObj2) Then
112                 PuedeLanzar = True
                    'Exit Function
               
                Else
114                 Call WriteConsoleMsg(Userindex, "Necesitas un " & ObjData(Hechizos(HechizoIndex).NecesitaObj).name & " para lanzar el hechizo.", FontTypeNames.FONTTYPE_INFO)
116                 PuedeLanzar = False
                    Exit Function

                End If

            End If
    
118         If Hechizos(HechizoIndex).CoolDown > 0 Then

                Dim actual            As Long

                Dim segundosFaltantes As Long

120             actual = GetTickCount() And &H7FFFFFFF

122             If UserList(Userindex).Counters.UserHechizosInterval(Slot) + (Hechizos(HechizoIndex).CoolDown * 1000) > actual Then
124                 segundosFaltantes = Int((UserList(Userindex).Counters.UserHechizosInterval(Slot) + (Hechizos(HechizoIndex).CoolDown * 1000) - actual) / 1000)
126                 Call WriteConsoleMsg(Userindex, "Debes esperar " & segundosFaltantes & " segundos para volver a tirar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
128                 PuedeLanzar = False
                    Exit Function
                Else
130                 PuedeLanzar = True

                End If

            End If
    
132         If UserList(Userindex).Stats.MinHp > Hechizos(HechizoIndex).RequiredHP Then
134             If UserList(Userindex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
136                 If UserList(Userindex).Stats.UserSkills(eSkill.magia) >= Hechizos(HechizoIndex).MinSkill Then
138                     If UserList(Userindex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
140                         PuedeLanzar = True
                        Else
142                         Call WriteLocaleMsg(Userindex, "93", FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "Est�s muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
144                         PuedeLanzar = False

                        End If
                    
                    Else
146                     Call WriteConsoleMsg(Userindex, "No tenes suficientes puntos de magia para lanzar este hechizo, necesitas " & Hechizos(HechizoIndex).MinSkill & " puntos.", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteLocaleMsg(UserIndex, "221", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Necesitas " & Hechizos(HechizoIndex).MinSkill & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
148                     PuedeLanzar = False

                    End If

                Else
150                 Call WriteLocaleMsg(Userindex, "222", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "No tenes suficiente mana. Necesitas " & Hechizos(HechizoIndex).ManaRequerido & " puntos de mana.", FontTypeNames.FONTTYPE_INFO)
152                 PuedeLanzar = False

                End If

            Else
154             Call WriteConsoleMsg(Userindex, "No tenes suficiente vida. Necesitas " & Hechizos(HechizoIndex).RequiredHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
156             PuedeLanzar = False

            End If

        Else
            'Call WriteConsoleMsg(UserIndex, "No podes lanzar hechizos porque estas muerto.", FontTypeNames.FONTTYPE_INFO)
158         Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
160         PuedeLanzar = False

        End If

        
        Exit Function

PuedeLanzar_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.PuedeLanzar", Erl)
        Resume Next
        
End Function

Sub HechizoTerrenoEstado(ByVal Userindex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoTerrenoEstado_Err
        

        Dim PosCasteadaX As Integer

        Dim PosCasteadaY As Integer

        Dim PosCasteadaM As Integer

        Dim h            As Integer

        Dim TempX        As Integer

        Dim TempY        As Integer

100     PosCasteadaX = UserList(Userindex).flags.TargetX
102     PosCasteadaY = UserList(Userindex).flags.TargetY
104     PosCasteadaM = UserList(Userindex).flags.TargetMap
    
106     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
        
108     If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
110         b = True

            'Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(H).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
112         For TempX = PosCasteadaX - 11 To PosCasteadaX + 11
114             For TempY = PosCasteadaY - 11 To PosCasteadaY + 11

116                 If InMapBounds(PosCasteadaM, TempX, TempY) Then
118                     If MapData(PosCasteadaM, TempX, TempY).Userindex > 0 Then

                            'hay un user
120                         If UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).flags.NoDetectable = 0 Then
122                             UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).flags.invisible = 0
124                             Call WriteConsoleMsg(MapData(PosCasteadaM, TempX, TempY).Userindex, "Tu invisibilidad ya no tiene efecto.", FontTypeNames.FONTTYPE_INFOIAO)
126                             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageSetInvisible(UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).Char.CharIndex, False))

                            End If

                        End If

                    End If

128             Next TempY
130         Next TempX
    
132         Call InfoHechizo(Userindex)

        End If

        
        Exit Sub

HechizoTerrenoEstado_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoTerrenoEstado", Erl)
        Resume Next
        
End Sub

Sub HechizoSobreArea(ByVal Userindex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoSobreArea_Err
        

        Dim PosCasteadaX As Byte

        Dim PosCasteadaY As Byte

        Dim PosCasteadaM As Integer

        Dim h            As Integer

        Dim TempX        As Integer

        Dim TempY        As Integer

100     PosCasteadaX = UserList(Userindex).flags.TargetX
102     PosCasteadaY = UserList(Userindex).flags.TargetY
104     PosCasteadaM = UserList(Userindex).flags.TargetMap
 
106     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)

        Dim X         As Long

        Dim Y         As Long
    
        Dim NPCIndex2 As Integer

        Dim Cuantos   As Long
    
        'Envio Palabras magicas, wavs y fxs.
108     If UserList(Userindex).flags.NoPalabrasMagicas = 0 Then
110         Call DecirPalabrasMagicas(h, Userindex)

        End If
    
112     If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
    
114         If Hechizos(h).ParticleViaje > 0 Then
116             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFXWithDestinoXY(UserList(Userindex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, 1, Hechizos(h).wav, 1, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))
                
            Else
118             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))

            End If

        End If
    
120     If Hechizos(h).Particle > 0 Then 'Envio Particula?
122         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFXToFloor(UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))

        End If
    
124     If Hechizos(h).ParticleViaje = 0 Then
126         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(h).wav, PosCasteadaX, PosCasteadaY))  'Esta linea faltaba. Pablo (ToxicWaste)

        End If

        Dim cuantosuser As Byte

        Dim nameuser    As String
       
128     Select Case Hechizos(h).AreaAfecta

            Case 1

130             For X = 1 To Hechizos(h).AreaRadio
132                 For Y = 1 To Hechizos(h).AreaRadio

134                     If MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).Userindex > 0 Then
136                         NPCIndex2 = MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).Userindex

                            'If NPCIndex2 <> UserIndex Then
138                         If UserList(NPCIndex2).flags.Muerto = 0 Then
                                        
140                             AreaHechizo Userindex, NPCIndex2, PosCasteadaX, PosCasteadaY, False
142                             cuantosuser = cuantosuser + 1
                                ' nameuser = nameuser & "," & Npclist(NPCIndex2).Name
                                            
                            End If

                            ' End If
                        End If

                    Next
                Next
                    
                ' If cuantosuser > 0 Then
                '     Call WriteConsoleMsg(UserIndex, "Has alcanzado a " & cuantosuser & " usuarios.", FontTypeNames.FONTTYPE_FIGHT)
                ' End If
144         Case 2

146             For X = 1 To Hechizos(h).AreaRadio
148                 For Y = 1 To Hechizos(h).AreaRadio

150                     If MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex > 0 Then
152                         NPCIndex2 = MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex

154                         If Npclist(NPCIndex2).Attackable Then
156                             AreaHechizo Userindex, NPCIndex2, PosCasteadaX, PosCasteadaY, True
158                             Cuantos = Cuantos + 1

                            End If

                        End If

                    Next
                Next
                
                ' If Cuantos > 0 Then
                '  Call WriteConsoleMsg(UserIndex, "Has alcanzado a " & Cuantos & " criaturas.", FontTypeNames.FONTTYPE_FIGHT)
                '  End If
160         Case 3

162             For X = 1 To Hechizos(h).AreaRadio
164                 For Y = 1 To Hechizos(h).AreaRadio

166                     If MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).Userindex > 0 Then
168                         NPCIndex2 = MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).Userindex

                            'If NPCIndex2 <> UserIndex Then
170                         If UserList(NPCIndex2).flags.Muerto = 0 Then
172                             AreaHechizo Userindex, NPCIndex2, PosCasteadaX, PosCasteadaY, False
174                             cuantosuser = cuantosuser + 1

                            End If

                            ' End If
                        End If
                            
176                     If MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex > 0 Then
178                         NPCIndex2 = MapData(UserList(Userindex).Pos.Map, X + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex

180                         If Npclist(NPCIndex2).Attackable Then
182                             AreaHechizo Userindex, NPCIndex2, PosCasteadaX, PosCasteadaY, True
184                             Cuantos = Cuantos + 1
            
                            End If

                        End If
                            
                    Next
                Next
                
                ' If Cuantos > 0 Then
                '   Call WriteConsoleMsg(UserIndex, "Has alcanzado a " & Cuantos & " criaturas", FontTypeNames.FONTTYPE_FIGHT)
                ' End If
                'If cuantosuser > 0 Then
                ' Call WriteConsoleMsg(UserIndex, "Has alcanzado a a " & cuantosuser & " usuarios.", FontTypeNames.FONTTYPE_FIGHT)
                ' End If
        End Select

186     b = True

        
        Exit Sub

HechizoSobreArea_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoSobreArea", Erl)
        Resume Next
        
End Sub

Sub HechizoPortal(ByVal Userindex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoPortal_Err
        

100     If UserList(Userindex).flags.BattleModo = 1 Then
102         b = False
            'Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
104         Call WriteLocaleMsg(Userindex, "262", FontTypeNames.FONTTYPE_INFO)
        Else

            Dim PosCasteadaX As Byte

            Dim PosCasteadaY As Byte

            Dim PosCasteadaM As Integer

            Dim uh           As Integer

            Dim TempX        As Integer

            Dim TempY        As Integer

106         PosCasteadaX = UserList(Userindex).flags.TargetX
108         PosCasteadaY = UserList(Userindex).flags.TargetY
110         PosCasteadaM = UserList(Userindex).flags.TargetMap
 
112         uh = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
    
            'Envio Palabras magicas, wavs y fxs.
   
114         If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY).ObjInfo.Amount > 0 Or MapData(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY).Blocked Or MapData(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY).TileExit.Map > 0 Or UserList(Userindex).flags.TargetUser <> 0 Then
116             b = False
                'Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
118             Call WriteLocaleMsg(Userindex, "262", FontTypeNames.FONTTYPE_INFO)

            Else

120             If Hechizos(uh).TeleportX = 1 Then

122                 If UserList(Userindex).flags.Portal = 0 Then

124                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFX(UserList(Userindex).Char.CharIndex, ParticulasIndex.Runa, -1, False))
            
126                     UserList(Userindex).flags.PortalM = UserList(Userindex).Pos.Map
128                     UserList(Userindex).flags.PortalX = UserList(Userindex).flags.TargetX
130                     UserList(Userindex).flags.PortalY = UserList(Userindex).flags.TargetY
            
132                     Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageBarFx(UserList(Userindex).Char.CharIndex, 600, Accion_Barra.Intermundia))

134                     UserList(Userindex).Accion.AccionPendiente = True
136                     UserList(Userindex).Accion.Particula = ParticulasIndex.Runa
138                     UserList(Userindex).Accion.TipoAccion = Accion_Barra.Intermundia
140                     UserList(Userindex).Accion.HechizoPendiente = uh
            
142                     If UserList(Userindex).flags.NoPalabrasMagicas = 0 Then
144                         Call DecirPalabrasMagicas(uh, Userindex)

                        End If

146                     b = True
                    Else
148                     Call WriteConsoleMsg(Userindex, "No pod�s lanzar mas de un portal a la vez.", FontTypeNames.FONTTYPE_INFO)
150                     b = False

                    End If

                End If

            End If

        End If

        
        Exit Sub

HechizoPortal_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoPortal", Erl)
        Resume Next
        
End Sub

Sub HechizoMaterializacion(ByVal Userindex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoMaterializacion_Err
        

        Dim h   As Integer

        Dim MAT As obj

100     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
 
102     If MapData(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY).ObjInfo.Amount > 0 Or MapData(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY).Blocked Then
104         b = False
106         Call WriteLocaleMsg(Userindex, "262", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
        Else
108         MAT.Amount = Hechizos(h).MaterializaCant
110         MAT.ObjIndex = Hechizos(h).MaterializaObj
112         Call MakeObj(MAT, UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY)
            'Call WriteConsoleMsg(UserIndex, "Has materializado un objeto!!", FontTypeNames.FONTTYPE_INFO)
114         Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFXToFloor(UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))
116         b = True

        End If

        
        Exit Sub

HechizoMaterializacion_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoMaterializacion", Erl)
        Resume Next
        
End Sub

Sub HandleHechizoTerreno(ByVal Userindex As Integer, ByVal uh As Integer)
        
        On Error GoTo HandleHechizoTerreno_Err
        

        '***************************************************
        'Author: Unknown
        'Last Modification: 01/10/07
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
        'usuario
        '***************************************************
        Dim b As Boolean

100     Select Case Hechizos(uh).Tipo
        
            Case TipoHechizo.uInvocacion 'Tipo 1

                'Call HechizoInvocacion(UserIndex, b)
102         Case TipoHechizo.uEstado 'Tipo 2
104             Call HechizoTerrenoEstado(Userindex, b)

106         Case TipoHechizo.uMaterializa 'Tipo 3
108             Call HechizoMaterializacion(Userindex, b)
            
110         Case TipoHechizo.uArea 'Tipo 5
112             Call HechizoSobreArea(Userindex, b)
            
114         Case TipoHechizo.uPortal 'Tipo 6
116             Call HechizoPortal(Userindex, b)

118         Case TipoHechizo.UFamiliar

                ' Call InvocarFamiliar(UserIndex, b)
        End Select

        'If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.Amount > 0 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Or UserList(UserIndex).flags.TargetUser <> 0 Then
        '  b = False
        '  Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)

        'Else

120     If b Then
122         Call SubirSkill(Userindex, magia)

            'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
124         If UserList(Userindex).clase = eClass.Druid And UserList(Userindex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
126             UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido * 0.7
            Else
128             UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido

            End If

130         If UserList(Userindex).Stats.MinMAN < 0 Then UserList(Userindex).Stats.MinMAN = 0
132         UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Hechizos(uh).StaRequerido

134         If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0
136         Call WriteUpdateMana(Userindex)
            Call WriteUpdateSta(Userindex)

        End If

        
        Exit Sub

HandleHechizoTerreno_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HandleHechizoTerreno", Erl)
        Resume Next
        
End Sub

Sub HandleHechizoUsuario(ByVal Userindex As Integer, ByVal uh As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 01/10/07
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
        'usuario
        '***************************************************
        
        On Error GoTo HandleHechizoUsuario_Err
        

        Dim b As Boolean

100     Select Case Hechizos(uh).Tipo

            Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
102             Call HechizoEstadoUsuario(Userindex, b)

104         Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
106             Call HechizoPropUsuario(Userindex, b)

108         Case TipoHechizo.uCombinados
110             Call HechizoCombinados(Userindex, b)
    
        End Select

112     If b Then
114         Call SubirSkill(Userindex, magia)
            'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
116         UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido
118         If UserList(Userindex).Stats.MinMAN < 0 Then UserList(Userindex).Stats.MinMAN = 0

            If Hechizos(uh).RequiredHP > 0 Then
120             UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MinHp - Hechizos(uh).RequiredHP
122             If UserList(Userindex).Stats.MinHp < 0 Then UserList(Userindex).Stats.MinHp = 1
                Call WriteUpdateHP(Userindex)
            End If

124         UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Hechizos(uh).StaRequerido
126         If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0
128
            Call WriteUpdateMana(Userindex)
            Call WriteUpdateSta(Userindex)
132         UserList(Userindex).flags.TargetUser = 0

        End If

        
        Exit Sub

HandleHechizoUsuario_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HandleHechizoUsuario", Erl)
        Resume Next
        
End Sub

Sub HandleHechizoNPC(ByVal Userindex As Integer, ByVal uh As Integer)
        
        On Error GoTo HandleHechizoNPC_Err
        

        '***************************************************
        'Author: Unknown
        'Last Modification: 01/10/07
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
        'usuario
        '***************************************************
        Dim b As Boolean

100     Select Case Hechizos(uh).Tipo

            Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
102             Call HechizoEstadoNPC(UserList(Userindex).flags.TargetNPC, uh, b, Userindex)

104         Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
106             Call HechizoPropNPC(uh, UserList(Userindex).flags.TargetNPC, Userindex, b)

        End Select

108     If b Then
110         Call SubirSkill(Userindex, magia)
112         UserList(Userindex).flags.TargetNPC = 0
114         UserList(Userindex).Stats.MinMAN = UserList(Userindex).Stats.MinMAN - Hechizos(uh).ManaRequerido

            If Hechizos(uh).RequiredHP > 0 Then
116             If UserList(Userindex).Stats.MinMAN < 0 Then UserList(Userindex).Stats.MinMAN = 0
118             UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MinHp - Hechizos(uh).RequiredHP
                Call WriteUpdateHP(Userindex)
            End If

120         If UserList(Userindex).Stats.MinHp < 0 Then UserList(Userindex).Stats.MinHp = 1
122         UserList(Userindex).Stats.MinSta = UserList(Userindex).Stats.MinSta - Hechizos(uh).StaRequerido

124         If UserList(Userindex).Stats.MinSta < 0 Then UserList(Userindex).Stats.MinSta = 0
126         Call WriteUpdateMana(Userindex)
            Call WriteUpdateSta(Userindex)

        End If

        
        Exit Sub

HandleHechizoNPC_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HandleHechizoNPC", Erl)
        Resume Next
        
End Sub

Sub LanzarHechizo(Index As Integer, Userindex As Integer)
        
        On Error GoTo LanzarHechizo_Err
        

        Dim uh As Integer

100     uh = UserList(Userindex).Stats.UserHechizos(Index)

102     If PuedeLanzar(Userindex, uh, Index) Then

104         Select Case Hechizos(uh).Target

                Case TargetType.uUsuarios

106                 If UserList(Userindex).flags.TargetUser > 0 Then
108                     If Abs(UserList(UserList(Userindex).flags.TargetUser).Pos.Y - UserList(Userindex).Pos.Y) <= RANGO_VISION_Y Then
110                         Call HandleHechizoUsuario(Userindex, uh)
                    
112                         If Hechizos(uh).CoolDown > 0 Then
114                             UserList(Userindex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF

                            End If

                        Else
116                         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)

                            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
118                     Call WriteConsoleMsg(Userindex, "Este hechizo actua solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
120             Case TargetType.uNPC

122                 If UserList(Userindex).flags.TargetNPC > 0 Then
124                     If Abs(Npclist(UserList(Userindex).flags.TargetNPC).Pos.Y - UserList(Userindex).Pos.Y) <= RANGO_VISION_Y Then
126                         Call HandleHechizoNPC(Userindex, uh)

128                         If Hechizos(uh).CoolDown > 0 Then
130                             UserList(Userindex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF
                    
                            End If
                    
                        Else
132                         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)

                            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
134                     Call WriteConsoleMsg(Userindex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
136             Case TargetType.uUsuariosYnpc

138                 If UserList(Userindex).flags.TargetUser > 0 Then
140                     If Abs(UserList(UserList(Userindex).flags.TargetUser).Pos.Y - UserList(Userindex).Pos.Y) <= RANGO_VISION_Y Then
142                         Call HandleHechizoUsuario(Userindex, uh)
                    
144                         If Hechizos(uh).CoolDown > 0 Then
146                             UserList(Userindex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF

                            End If

                        Else
148                         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)

                            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

150                 ElseIf UserList(Userindex).flags.TargetNPC > 0 Then

152                     If Abs(Npclist(UserList(Userindex).flags.TargetNPC).Pos.Y - UserList(Userindex).Pos.Y) <= RANGO_VISION_Y Then
154                         If Hechizos(uh).CoolDown > 0 Then
156                             UserList(Userindex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF

                            End If

158                         Call HandleHechizoNPC(Userindex, uh)
                        Else
160                         Call WriteLocaleMsg(Userindex, "8", FontTypeNames.FONTTYPE_INFO)

                            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
162                     Call WriteConsoleMsg(Userindex, "Target invalido.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
164             Case TargetType.uTerreno

166                 If Hechizos(uh).CoolDown > 0 Then
168                     UserList(Userindex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF

                    End If

170                 Call HandleHechizoTerreno(Userindex, uh)

            End Select
    
        End If

172     If UserList(Userindex).Counters.Trabajando Then
174         Call WriteMacroTrabajoToggle(Userindex, False)

        End If

176     If UserList(Userindex).Counters.Ocultando Then UserList(Userindex).Counters.Ocultando = UserList(Userindex).Counters.Ocultando - 1
    
        
        Exit Sub

LanzarHechizo_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.LanzarHechizo", Erl)
        Resume Next
        
End Sub

Sub HechizoEstadoUsuario(ByVal Userindex As Integer, ByRef b As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 02/01/2008
        'Handles the Spells that afect the Stats of an User
        '24/01/2007 Pablo (ToxicWaste) - Invisibilidad no permitida en Mapas con InviSinEfecto
        '26/01/2007 Pablo (ToxicWaste) - Cambios que permiten mejor manejo de ataques en los rings.
        '26/01/2007 Pablo (ToxicWaste) - Revivir no permitido en Mapas con ResuSinEfecto
        '02/01/2008 Marcos (ByVal) - Curar Veneno no permitido en usuarios muertos.
        '***************************************************
        
        On Error GoTo HechizoEstadoUsuario_Err
        

        Dim h As Integer, tU As Integer

100     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
102     tU = UserList(Userindex).flags.TargetUser

104     If Hechizos(h).Invisibilidad = 1 Then
   
106         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "�Est� muerto!", FontTypeNames.FONTTYPE_INFO)
108             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
110             b = False
                Exit Sub

            End If
    
112         If UserList(tU).Counters.Saliendo Then
114             If Userindex <> tU Then
116                 Call WriteConsoleMsg(Userindex, "�El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
118                 b = False
                    Exit Sub
                Else
120                 Call WriteConsoleMsg(Userindex, "�No pod�s ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
122                 b = False
                    Exit Sub

                End If

            End If
    
            'No usar invi mapas InviSinEfecto
            ' If MapInfo(UserList(tU).Pos.map).InviSinEfecto > 0 Then
            '  Call WriteConsoleMsg(UserIndex, "�La invisibilidad no funciona aqu�!", FontTypeNames.FONTTYPE_INFO)
            '  b = False
            '   Exit Sub
            '  End If
    
            'Para poder tirar invi a un pk en el ring
124         If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
126             If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
128                 If esArmada(Userindex) Then
130                     Call WriteConsoleMsg(Userindex, "Los miembros de la armada real no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
132                     b = False
                        Exit Sub

                    End If

134                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
136                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
138                     b = False
                        Exit Sub
                    Else
140                     Call VolverCriminal(Userindex)

                    End If

                End If

            End If
    
            'Si sos user, no uses este hechizo con GMS.
142         If UserList(Userindex).flags.Privilegios And PlayerType.user Then
144             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
            
            If UserList(tU).flags.invisible = 1 Then
                If tU = Userindex Then
                    Call WriteConsoleMsg(Userindex, "�Ya est�s invisible!", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "�El objetivo ya se encuentra invisible!", FontTypeNames.FONTTYPE_INFO)
                End If
                b = False
                Exit Sub
            End If
   
146         UserList(tU).flags.invisible = 1
            'Ladder
            'Reseteamos el contador de Invisibilidad
148         UserList(tU).Counters.Invisibilidad = Hechizos(h).Duration
150         Call WriteContadores(tU)
152         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

154         Call InfoHechizo(Userindex)
156         b = True

        End If

158     If Hechizos(h).Envenena > 0 Then
            ' If UserIndex = tU Then
            '    Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            '   Exit Sub
            'End If
    
160         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
162         If Userindex <> tU Then
164             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

166         UserList(tU).flags.Envenenado = Hechizos(h).Envenena
168         Call InfoHechizo(Userindex)
170         b = True

        End If

172     If Hechizos(h).desencantar = 1 Then
            ' Call WriteConsoleMsg(UserIndex, "Ningun efecto magico tiene efecto sobre ti ya.", FontTypeNames.FONTTYPE_INFOIAO)

174         UserList(Userindex).flags.Envenenado = 0
176         UserList(Userindex).flags.Incinerado = 0
    
178         If UserList(Userindex).flags.Inmovilizado = 1 Then
180             UserList(Userindex).Counters.Inmovilizado = 0
182             UserList(Userindex).flags.Inmovilizado = 0
184             Call WriteInmovilizaOK(Userindex)
            

            End If
    
186         If UserList(Userindex).flags.Paralizado = 1 Then
188             UserList(Userindex).flags.Paralizado = 0
190             Call WriteParalizeOK(Userindex)
            
           
            End If
        
192         If UserList(Userindex).flags.Ceguera = 1 Then
194             UserList(Userindex).flags.Ceguera = 0
196             Call WriteBlindNoMore(Userindex)
            

            End If
    
198         If UserList(Userindex).flags.Maldicion = 1 Then
200             UserList(Userindex).flags.Maldicion = 0
202             UserList(Userindex).Counters.Maldicion = 0

            End If
    
204         Call InfoHechizo(Userindex)
206         b = True

        End If

208     If Hechizos(h).incinera = 1 Then
210         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
212             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
214         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
216         If Userindex <> tU Then
218             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

220         UserList(tU).flags.Incinerado = 1
222         Call InfoHechizo(Userindex)
224         b = True

        End If

226     If Hechizos(h).CuraVeneno = 1 Then

            'Verificamos que el usuario no este muerto
228         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "�Est� muerto!", FontTypeNames.FONTTYPE_INFO)
230             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
232             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar veneno a un pk en el ring
234         If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
236             If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
238                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
240                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
242                     b = False
                        Exit Sub

                    End If

244                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
246                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
248                     b = False
                        Exit Sub
                    Else

                        '    Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
        
            'Si sos user, no uses este hechizo con GMS.
250         If UserList(Userindex).flags.Privilegios And PlayerType.user Then
252             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
        
254         UserList(tU).flags.Envenenado = 0
256         Call InfoHechizo(Userindex)
258         b = True

        End If

260     If Hechizos(h).Maldicion = 1 Then
262         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
264             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
266         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
268         If Userindex <> tU Then
270             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

272         UserList(tU).flags.Maldicion = 1
274         UserList(tU).Counters.Maldicion = 200
    
276         Call InfoHechizo(Userindex)
278         b = True

        End If

280     If Hechizos(h).RemoverMaldicion = 1 Then
282         UserList(tU).flags.Maldicion = 0
284         Call InfoHechizo(Userindex)
286         b = True

        End If

288     If Hechizos(h).GolpeCertero = 1 Then
290         UserList(tU).flags.GolpeCertero = 1
292         Call InfoHechizo(Userindex)
294         b = True

        End If

296     If Hechizos(h).Bendicion = 1 Then
298         UserList(tU).flags.Bendicion = 1
300         Call InfoHechizo(Userindex)
302         b = True

        End If

304     If Hechizos(h).Paraliza = 1 Then
306         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
308             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            If UserList(tU).flags.Paralizado = 1 Then
309             Call WriteConsoleMsg(Userindex, UserList(tU).name & " ya est� paralizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
310         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
            
312         If Userindex <> tU Then
314             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If
            
316         Call InfoHechizo(Userindex)
318         b = True

320         If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
322             Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
324             Call WriteConsoleMsg(Userindex, " �El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            
                Exit Sub

            End If
            
326         UserList(tU).Counters.Paralisis = Hechizos(h).Duration

328         If UserList(tU).flags.Paralizado = 0 Then
330             UserList(tU).flags.Paralizado = 1
332             Call WriteParalizeOK(tU)
334             Call WritePosUpdate(tU)
            End If

        End If

336     If Hechizos(h).Velocidad > 0 Then
338         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
340             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
342         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
            
344         If Userindex <> tU Then
346             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If
            
348         Call InfoHechizo(Userindex)
350         b = True
                 
352         If UserList(tU).Counters.Velocidad = 0 Then
354             UserList(tU).flags.VelocidadBackup = UserList(tU).Char.speeding
            End If

356         UserList(tU).Char.speeding = Hechizos(h).Velocidad
358         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSpeedingACT(UserList(tU).Char.CharIndex, UserList(tU).Char.speeding))
            'End If
360         UserList(tU).Counters.Velocidad = Hechizos(h).Duration

        End If

362     If Hechizos(h).Inmoviliza = 1 Then
364         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
366             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
            If UserList(tU).flags.Paralizado = 1 Then
                Call WriteConsoleMsg(Userindex, UserList(tU).name & " ya est� paralizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
368         ElseIf UserList(tU).flags.Inmovilizado = 1 Then
370             Call WriteConsoleMsg(Userindex, UserList(tU).name & " ya est� inmovilizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
372         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
            
374         If Userindex <> tU Then
376             Call UsuarioAtacadoPorUsuario(Userindex, tU)
            End If
            
378         Call InfoHechizo(Userindex)
380         b = True
            '  If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
            '   Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
            '   Call WriteConsoleMsg(UserIndex, " �El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            '
            '    Exit Sub
            ' End If
            
382         UserList(tU).Counters.Inmovilizado = Hechizos(h).Duration

386         UserList(tU).flags.Inmovilizado = 1
388         Call WriteInmovilizaOK(tU)
390         Call WritePosUpdate(tU)
            

        End If

392     If Hechizos(h).RemoverParalisis = 1 Then
        
            'Para poder tirar remo a un pk en el ring
394         If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
396             If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
398                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
400                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
402                     b = False
                        Exit Sub

                    End If

404                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
406                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
408                     b = False
                        Exit Sub
                    Else
410                     Call VolverCriminal(Userindex)

                    End If

                End If

            End If
        
412         If UserList(tU).flags.Inmovilizado = 0 And UserList(tU).flags.Paralizado = 0 Then
414             Call WriteConsoleMsg(Userindex, "El objetivo no esta paralizado.", FontTypeNames.FONTTYPE_INFO)
416             b = False
                Exit Sub

            End If
        
418         If UserList(tU).flags.Inmovilizado = 1 Then
420             UserList(tU).Counters.Inmovilizado = 0
422             UserList(tU).flags.Inmovilizado = 0
424             Call WriteInmovilizaOK(tU)
426             Call WritePosUpdate(tU)
                ' Call InfoHechizo(UserIndex)
            

                'b = True
            End If
    
428         If UserList(tU).flags.Paralizado = 1 Then
430             UserList(tU).flags.Paralizado = 0
                'no need to crypt this
432             Call WriteParalizeOK(tU)
            

                '  b = True
            End If

434         b = True
436         Call InfoHechizo(Userindex)

        End If

438     If Hechizos(h).RemoverEstupidez = 1 Then
440         If UserList(tU).flags.Estupidez = 1 Then

                'Para poder tirar remo estu a un pk en el ring
442             If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
444                 If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
446                     If esArmada(Userindex) Then
                            'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
448                         Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
450                         b = False
                            Exit Sub

                        End If

452                     If UserList(Userindex).flags.Seguro Then
                            'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
454                         Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
456                         b = False
                            Exit Sub
                        Else

                            ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                        End If

                    End If

                End If
    
458             UserList(tU).flags.Estupidez = 0
                'no need to crypt this
460             Call WriteDumbNoMore(tU)
            
462             Call InfoHechizo(Userindex)
464             b = True

            End If

        End If

466     If Hechizos(h).Revivir = 1 Then
468         If UserList(tU).flags.Muerto = 1 Then
    
                'No usar resu en mapas con ResuSinEfecto
                'If MapInfo(UserList(tU).Pos.map).ResuSinEfecto > 0 Then
                '   Call WriteConsoleMsg(UserIndex, "�Revivir no est� permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
                '   b = False
                '   Exit Sub
                ' End If
        
470             If UserList(tU).Accion.TipoAccion = Accion_Barra.Resucitar Then
472                 Call WriteConsoleMsg(Userindex, "El usuario ya esta siendo resucitado.", FontTypeNames.FONTTYPE_INFO)
474                 b = False
                    Exit Sub

                End If
        
                'Para poder tirar revivir a un pk en el ring
476             If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
478                 If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
480                     If esArmada(Userindex) Then
                            'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
482                         Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
484                         b = False
                            Exit Sub

                        End If

486                     If UserList(Userindex).flags.Seguro Then
                            'call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
488                         Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
490                         b = False
                            Exit Sub
                        Else
492                         Call VolverCriminal(Userindex)

                        End If

                    End If

                End If
                        
494             Call SendData(SendTarget.ToPCArea, tU, PrepareMessageParticleFX(UserList(tU).Char.CharIndex, ParticulasIndex.Resucitar, 600, False))
496             Call SendData(SendTarget.ToPCArea, tU, PrepareMessageBarFx(UserList(tU).Char.CharIndex, 600, Accion_Barra.Resucitar))
498             UserList(tU).Accion.AccionPendiente = True
500             UserList(tU).Accion.Particula = ParticulasIndex.Resucitar
502             UserList(tU).Accion.TipoAccion = Accion_Barra.Resucitar
                
                'Pablo Toxic Waste (GD: 29/04/07)
                'UserList(tU).Stats.MinAGU = 0
                'UserList(tU).flags.Sed = 1
                'UserList(tU).Stats.MinHam = 0
                'UserList(tU).flags.Hambre = 1
504             Call WriteUpdateHungerAndThirst(tU)
506             Call InfoHechizo(Userindex)
                'UserList(tU).Stats.MinMAN = 0
                'UserList(tU).Stats.MinSta = 0
508             b = True
                'Solo saco vida si es User. no quiero que exploten GMs por ahi.
        
                'Call RevivirUsuario(tU)
            Else
510             b = False

            End If

        End If

512     If Hechizos(h).Ceguera = 1 Then
514         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
516             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
518         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
520         If Userindex <> tU Then
522             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

524         UserList(tU).flags.Ceguera = 1
526         UserList(tU).Counters.Ceguera = Hechizos(h).Duration

528         Call WriteBlind(tU)
        
530         Call InfoHechizo(Userindex)
532         b = True

        End If

534     If Hechizos(h).Estupidez = 1 Then
536         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
538             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

540         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
542         If Userindex <> tU Then
544             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

546         If UserList(tU).flags.Estupidez = 0 Then
548             UserList(tU).flags.Estupidez = 1
550             UserList(tU).Counters.Estupidez = Hechizos(h).Duration

            End If

552         Call WriteDumb(tU)
        

554         Call InfoHechizo(Userindex)
556         b = True

        End If

        
        Exit Sub

HechizoEstadoUsuario_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoEstadoUsuario", Erl)
        Resume Next
        
End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal Userindex As Integer)
        
        On Error GoTo HechizoEstadoNPC_Err
        

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 04/13/2008
        'Handles the Spells that afect the Stats of an NPC
        '04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
        'removidos por users de su misma faccion.
        '***************************************************
100     If Hechizos(hIndex).Invisibilidad = 1 Then
102         Call InfoHechizo(Userindex)
104         Npclist(NpcIndex).flags.invisible = 1
106         b = True

        End If

108     If Hechizos(hIndex).Envenena > 0 Then
110         If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
112             b = False
                Exit Sub

            End If

114         Call NPCAtacado(NpcIndex, Userindex)
116         Call InfoHechizo(Userindex)
118         Npclist(NpcIndex).flags.Envenenado = Hechizos(hIndex).Envenena
120         b = True

        End If

122     If Hechizos(hIndex).CuraVeneno = 1 Then
124         Call InfoHechizo(Userindex)
126         Npclist(NpcIndex).flags.Envenenado = 0
128         b = True

        End If

130     If Hechizos(hIndex).RemoverMaldicion = 1 Then
132         Call InfoHechizo(Userindex)
            'Npclist(NpcIndex).flags.Maldicion = 0
134         b = True

        End If

136     If Hechizos(hIndex).Bendicion = 1 Then
138         Call InfoHechizo(Userindex)
140         Npclist(NpcIndex).flags.Bendicion = 1
142         b = True

        End If

144     If Hechizos(hIndex).Paraliza = 1 Then
146         If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
148             If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
150                 b = False
                    Exit Sub

                End If

152             Call NPCAtacado(NpcIndex, Userindex)
154             Call InfoHechizo(Userindex)
156             Npclist(NpcIndex).flags.Paralizado = 1
158             Npclist(NpcIndex).flags.Inmovilizado = 0
160             Npclist(NpcIndex).Contadores.Paralisis = (Hechizos(hIndex).Duration * 6.5) * 2
162             b = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
164             Call WriteLocaleMsg(Userindex, "381", FontTypeNames.FONTTYPE_INFO)
166             b = False
                Exit Sub

            End If

        End If

168     If Hechizos(hIndex).RemoverParalisis = 1 Then
170         If Npclist(NpcIndex).flags.Paralizado = 1 Or Npclist(NpcIndex).flags.Inmovilizado = 1 Then
172             If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
174                 If esArmada(Userindex) Then
176                     Call InfoHechizo(Userindex)
178                     Npclist(NpcIndex).flags.Paralizado = 0
180                     Npclist(NpcIndex).Contadores.Paralisis = 0
182                     b = True
                        Exit Sub
                    Else
184                     Call WriteConsoleMsg(Userindex, "Solo pod�s Remover la Par�lisis de los Guardias si perteneces a su facci�n.", FontTypeNames.FONTTYPE_INFO)
186                     b = False
                        Exit Sub

                    End If
                
188                 Call WriteConsoleMsg(Userindex, "Solo pod�s Remover la Par�lisis de los NPCs que te consideren su amo", FontTypeNames.FONTTYPE_INFO)
190                 b = False
                    Exit Sub
                Else

192                 If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
194                     If esCaos(Userindex) Then
196                         Call InfoHechizo(Userindex)
198                         Npclist(NpcIndex).flags.Paralizado = 0
200                         Npclist(NpcIndex).Contadores.Paralisis = 0
202                         b = True
                            Exit Sub
                        Else
204                         Call WriteConsoleMsg(Userindex, "Solo pod�s Remover la Par�lisis de los Guardias si perteneces a su facci�n.", FontTypeNames.FONTTYPE_INFO)
206                         b = False
                            Exit Sub

                        End If

                    End If

                End If

            Else
208             Call WriteConsoleMsg(Userindex, "Este NPC no esta Paralizado", FontTypeNames.FONTTYPE_INFO)
210             b = False
                Exit Sub

            End If

        End If
 
212     If Hechizos(hIndex).Inmoviliza = 1 And Npclist(NpcIndex).flags.Paralizado = 0 Then
214         If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
216             If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
218                 b = False
                    Exit Sub

                End If

220             Call NPCAtacado(NpcIndex, Userindex)
222             Npclist(NpcIndex).flags.Inmovilizado = 1
224             Npclist(NpcIndex).flags.Paralizado = 0
226             Npclist(NpcIndex).Contadores.Paralisis = (Hechizos(hIndex).Duration * 6.5) * 2
228             Call InfoHechizo(Userindex)
230             b = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
232             Call WriteLocaleMsg(Userindex, "381", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        
        Exit Sub

HechizoEstadoNPC_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoEstadoNPC", Erl)
        Resume Next
        
End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal Userindex As Integer, ByRef b As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 14/08/2007
        'Handles the Spells that afect the Life NPC
        '14/08/2007 Pablo (ToxicWaste) - Orden general.
        '***************************************************
        
        On Error GoTo HechizoPropNPC_Err
        

        Dim da�o As Long
    
        'Salud
100     If Hechizos(hIndex).SubeHP = 1 Then
102         da�o = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
            'da�o = da�o + Porcentaje(da�o, 3 * UserList(UserIndex).Stats.ELV)
        
104         Call InfoHechizo(Userindex)
106         Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp + da�o

108         If Npclist(NpcIndex).Stats.MinHp > Npclist(NpcIndex).Stats.MaxHp Then Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MaxHp
110         Call WriteConsoleMsg(Userindex, "Has curado " & da�o & " puntos de salud a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
112         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead(da�o, Npclist(NpcIndex).Char.CharIndex, &HFF00))
114         b = True
        
116     ElseIf Hechizos(hIndex).SubeHP = 2 Then

118         If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
120             b = False
                Exit Sub

            End If
        
122         Call NPCAtacado(NpcIndex, Userindex)
124         da�o = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
        
126         da�o = da�o + Porcentaje(da�o, 3 * UserList(Userindex).Stats.ELV)
    
            ' If Hechizos(hIndex).StaffAffected Then
            '     If UserList(UserIndex).clase = eClass.Mage Then
            '         If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
            '             da�o = (da�o * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            '             'Aumenta da�o segun el staff-
            '             'Da�o = (Da�o* (70 + BonifB�culo)) / 100
            '         Else
            '             da�o = da�o * 0.7 'Baja da�o a 70% del original
            '         End If
            '     End If
            ' End If
        
            'If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDMAGICO Or UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
            '    da�o = da�o * 1.04  'laud magico de los bardos
            'End If
    
128         If UserList(Userindex).flags.Da�oMagico > 0 Then
130             da�o = da�o + Porcentaje(da�o, UserList(Userindex).flags.Da�oMagico)

            End If
    
132         b = True
        
134         If Npclist(NpcIndex).flags.Snd2 > 0 Then
136             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))

            End If
        
            'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
138         da�o = da�o - Npclist(NpcIndex).Stats.defM
        
140         If da�o < 0 Then da�o = 0
        
142         Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp - da�o
144         Call InfoHechizo(Userindex)
        
146         If UserList(Userindex).ChatCombate = 1 Then
148             Call WriteConsoleMsg(Userindex, "Le has causado " & da�o & " puntos de da�o a la criatura!", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
150         Call CalcularDarExp(Userindex, NpcIndex, da�o)
    
152         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead(da�o, Npclist(NpcIndex).Char.CharIndex))
    
154         If Npclist(NpcIndex).Stats.MinHp < 1 Then
156             Npclist(NpcIndex).Stats.MinHp = 0
158             Call MuereNpc(NpcIndex, Userindex)

            End If

        End If

        
        Exit Sub

HechizoPropNPC_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoPropNPC", Erl)
        Resume Next
        
End Sub

Sub InfoHechizo(ByVal Userindex As Integer)
        
        On Error GoTo InfoHechizo_Err
        

        Dim h As Integer

100     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
    
102     If UserList(Userindex).flags.NoPalabrasMagicas = 0 Then
104         Call DecirPalabrasMagicas(h, Userindex)

        End If

106     If UserList(Userindex).flags.TargetUser > 0 Then '�El Hechizo fue tirado sobre un usuario?
108         If Hechizos(h).FXgrh > 0 Then '�Envio FX?
110             If Hechizos(h).ParticleViaje > 0 Then
112                 Call SendData(SendTarget.ToPCArea, UserList(Userindex).flags.TargetUser, PrepareMessageParticleFXWithDestino(UserList(Userindex).Char.CharIndex, UserList(UserList(Userindex).flags.TargetUser).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1))
                Else
114                 Call SendData(SendTarget.ToPCArea, UserList(Userindex).flags.TargetUser, PrepareMessageCreateFX(UserList(UserList(Userindex).flags.TargetUser).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))

                End If

            End If

116         If Hechizos(h).Particle > 0 Then '�Envio Particula?
118             If Hechizos(h).ParticleViaje > 0 Then
120                 Call SendData(SendTarget.ToPCArea, UserList(Userindex).flags.TargetUser, PrepareMessageParticleFXWithDestino(UserList(Userindex).Char.CharIndex, UserList(UserList(Userindex).flags.TargetUser).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0))
                Else
122                 Call SendData(SendTarget.ToPCArea, UserList(Userindex).flags.TargetUser, PrepareMessageParticleFX(UserList(UserList(Userindex).flags.TargetUser).Char.CharIndex, Hechizos(h).Particle, Hechizos(h).TimeParticula, False))

                End If

            End If
        
124         If Hechizos(h).ParticleViaje = 0 Then
126             Call SendData(SendTarget.ToPCArea, UserList(Userindex).flags.TargetUser, PrepareMessagePlayWave(Hechizos(h).wav, UserList(UserList(Userindex).flags.TargetUser).Pos.X, UserList(UserList(Userindex).flags.TargetUser).Pos.Y)) 'Esta linea faltaba. Pablo (ToxicWaste)

            End If
        
128         If Hechizos(h).TimeEfect <> 0 Then 'Envio efecto de screen
130             Call WriteEfectToScreen(Userindex, Hechizos(h).ScreenColor, Hechizos(h).TimeEfect)

            End If

132     ElseIf UserList(Userindex).flags.TargetNPC > 0 Then '�El Hechizo fue tirado sobre un npc?

134         If Hechizos(h).FXgrh > 0 Then '�Envio FX?
136             If Npclist(UserList(Userindex).flags.TargetNPC).Stats.MinHp < 1 Then

                    'Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(H).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
138                 If Hechizos(h).ParticleViaje > 0 Then
140                     Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageParticleFXWithDestinoXY(UserList(Userindex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))
                    Else
142                     Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))

                    End If

                Else

144                 If Hechizos(h).ParticleViaje > 0 Then
146                     Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageParticleFXWithDestino(UserList(Userindex).Char.CharIndex, Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1))
                    Else
148                     Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageCreateFX(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))

                    End If

                End If

            End If
        
150         If Hechizos(h).Particle > 0 Then '�Envio Particula?
152             If Npclist(UserList(Userindex).flags.TargetNPC).Stats.MinHp < 1 Then
154                 Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageParticleFXWithDestinoXY(UserList(Userindex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0, Npclist(UserList(Userindex).flags.TargetNPC).Pos.X, Npclist(UserList(Userindex).flags.TargetNPC).Pos.Y))
                    'Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXToFloor(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.X, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y, Hechizos(H).Particle, Hechizos(H).TimeParticula))
                Else

156                 If Hechizos(h).ParticleViaje > 0 Then
158                     Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageParticleFXWithDestino(UserList(Userindex).Char.CharIndex, Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0))
                    Else
160                     Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessageParticleFX(Npclist(UserList(Userindex).flags.TargetNPC).Char.CharIndex, Hechizos(h).Particle, Hechizos(h).TimeParticula, False))

                    End If

                End If

            End If

162         If Hechizos(h).ParticleViaje = 0 Then
164             Call SendData(SendTarget.ToNPCArea, UserList(Userindex).flags.TargetNPC, PrepareMessagePlayWave(Hechizos(h).wav, Npclist(UserList(Userindex).flags.TargetNPC).Pos.X, Npclist(UserList(Userindex).flags.TargetNPC).Pos.Y))

            End If

        Else ' Entonces debe ser sobre el terreno

166         If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
168             Call modSendData.SendToAreaByPos(UserList(Userindex).Pos.Map, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))

            End If
        
170         If Hechizos(h).Particle > 0 Then 'Envio Particula?
172             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageParticleFXToFloor(UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))

            End If
        
174         If Hechizos(h).wav <> 0 Then
176             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(h).wav, UserList(Userindex).flags.TargetX, UserList(Userindex).flags.TargetY))   'Esta linea faltaba. Pablo (ToxicWaste)

            End If
    
        End If
    
178     If UserList(Userindex).ChatCombate = 1 Then
180         If UserList(Userindex).flags.TargetUser > 0 Then

                'Optimizacion de protocolo por Ladder
182             If Userindex <> UserList(Userindex).flags.TargetUser Then
184                 Call WriteConsoleMsg(Userindex, "HecMSGU*" & h & "*" & UserList(UserList(Userindex).flags.TargetUser).name, FontTypeNames.FONTTYPE_FIGHT)
186                 Call WriteConsoleMsg(UserList(Userindex).flags.TargetUser, "HecMSGA*" & h & "*" & UserList(Userindex).name, FontTypeNames.FONTTYPE_FIGHT)
    
                Else
188                 Call WriteConsoleMsg(Userindex, "ProMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)

                End If

190         ElseIf UserList(Userindex).flags.TargetNPC > 0 Then
192             Call WriteConsoleMsg(Userindex, "HecMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If

        
        Exit Sub

InfoHechizo_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.InfoHechizo", Erl)
        Resume Next
        
End Sub

Sub HechizoPropUsuario(ByVal Userindex As Integer, ByRef b As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 02/01/2008
        '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
        '***************************************************
        
        On Error GoTo HechizoPropUsuario_Err
        

        Dim h As Integer

        Dim da�o As Integer

        Dim tempChr           As Integer

        Dim enviarInfoHechizo As Boolean

100     enviarInfoHechizo = False
    
102     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
104     tempChr = UserList(Userindex).flags.TargetUser
      
        'Hambre
106     If Hechizos(h).SubeHam = 1 Then
    
108         enviarInfoHechizo = True
    
110         da�o = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
112         UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + da�o

114         If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
116         If Userindex <> tempChr Then
118             Call WriteConsoleMsg(Userindex, "Le has restaurado " & da�o & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
120             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha restaurado " & da�o & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
122             Call WriteConsoleMsg(Userindex, "Te has restaurado " & da�o & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
124         Call WriteUpdateHungerAndThirst(tempChr)
126         b = True
    
128     ElseIf Hechizos(h).SubeHam = 2 Then

130         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
132         If Userindex <> tempChr Then
134             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)
            Else
                Exit Sub

            End If
    
136         enviarInfoHechizo = True
    
138         da�o = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
140         UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - da�o
    
142         If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
144         If Userindex <> tempChr Then
146             Call WriteConsoleMsg(Userindex, "Le has quitado " & da�o & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
148             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha quitado " & da�o & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
150             Call WriteConsoleMsg(Userindex, "Te has quitado " & da�o & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
152         Call WriteUpdateHungerAndThirst(tempChr)
    
154         b = True
    
156         If UserList(tempChr).Stats.MinHam < 1 Then
158             UserList(tempChr).Stats.MinHam = 0
160             UserList(tempChr).flags.Hambre = 1

            End If
    
        End If

        'Sed
162     If Hechizos(h).SubeSed = 1 Then
    
164         enviarInfoHechizo = True
    
166         da�o = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
    
168         UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + da�o

170         If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
172         If Userindex <> tempChr Then
174             Call WriteConsoleMsg(Userindex, "Le has restaurado " & da�o & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
176             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha restaurado " & da�o & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
178             Call WriteConsoleMsg(Userindex, "Te has restaurado " & da�o & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
            Call WriteUpdateHungerAndThirst(tempChr)
    
180         b = True
    
182     ElseIf Hechizos(h).SubeSed = 2 Then
    
184         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
186         If Userindex <> tempChr Then
188             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
190         enviarInfoHechizo = True
    
192         da�o = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
    
194         UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - da�o
    
196         If Userindex <> tempChr Then
198             Call WriteConsoleMsg(Userindex, "Le has quitado " & da�o & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
200             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha quitado " & da�o & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
202             Call WriteConsoleMsg(Userindex, "Te has quitado " & da�o & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
204         If UserList(tempChr).Stats.MinAGU < 1 Then
206             UserList(tempChr).Stats.MinAGU = 0
208             UserList(tempChr).flags.Sed = 1

            End If
            
            Call WriteUpdateHungerAndThirst(tempChr)
    
210         b = True

        End If

        ' <-------- Agilidad ---------->
212     If Hechizos(h).SubeAgilidad = 1 Then
    
            'Para poder tirar cl a un pk en el ring
214         If (TriggerZonaPelea(Userindex, tempChr) <> TRIGGER6_PERMITE) Then
216             If Status(tempChr) = 0 And Status(Userindex) = 1 Or Status(tempChr) = 2 And Status(Userindex) = 1 Then
218                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
220                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
222                     b = False
                        Exit Sub

                    End If

224                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
226                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
228                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
230         enviarInfoHechizo = True
232         da�o = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
234         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

236         UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + da�o
         
238         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)

240         UserList(tempChr).flags.TomoPocion = True
242         b = True
244         Call WriteFYA(tempChr)
    
246     ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
248         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
250         If Userindex <> tempChr Then
252             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
254         enviarInfoHechizo = True
    
256         UserList(tempChr).flags.TomoPocion = True
258         da�o = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
260         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

262         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - da�o < MINATRIBUTOS Then
264             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
            Else
266             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - da�o

            End If
    
268         b = True
270         Call WriteFYA(tempChr)

        End If

        ' <-------- Fuerza ---------->
272     If Hechizos(h).SubeFuerza = 1 Then

            'Para poder tirar fuerza a un pk en el ring
274         If (TriggerZonaPelea(Userindex, tempChr) <> TRIGGER6_PERMITE) Then
276             If Status(tempChr) = 0 And Status(Userindex) = 1 Or Status(tempChr) = 2 And Status(Userindex) = 1 Then
278                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
280                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
282                     b = False
                        Exit Sub

                    End If

284                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
286                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
288                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
290         da�o = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
292         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
      
294         UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + da�o

296         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)
    
298         UserList(tempChr).flags.TomoPocion = True
            
            Call WriteFYA(tempChr)

300         b = True
    
302         enviarInfoHechizo = True
304         Call WriteFYA(tempChr)

306     ElseIf Hechizos(h).SubeFuerza = 2 Then

308         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
310         If Userindex <> tempChr Then
312             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
314         UserList(tempChr).flags.TomoPocion = True
    
316         da�o = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
318         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

320         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - da�o < MINATRIBUTOS Then
322             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
            Else
324             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - da�o

            End If

326         b = True
328         enviarInfoHechizo = True
330         Call WriteFYA(tempChr)

        End If

        'Salud
332     If Hechizos(h).SubeHP = 1 Then
    
            'Verifica que el usuario no este muerto
334         If UserList(tempChr).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "�Est� muerto!", FontTypeNames.FONTTYPE_INFO)
336             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
338             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar a un pk en el ring
340         If (TriggerZonaPelea(Userindex, tempChr) <> TRIGGER6_PERMITE) Then
342             If Status(tempChr) = 0 And Status(Userindex) = 1 Or Status(tempChr) = 2 And Status(Userindex) = 1 Then
344                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
346                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
348                     b = False
                        Exit Sub

                    End If

350                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
352                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
354                     b = False
                        Exit Sub
                    Else

                        'Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
       
356         da�o = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
            ' da�o = da�o + Porcentaje(da�o, 2 * UserList(UserIndex).Stats.ELV)
    
358         enviarInfoHechizo = True

360         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp + da�o

362         If UserList(tempChr).Stats.MinHp > UserList(tempChr).Stats.MaxHp Then UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp
    
364         If Userindex <> tempChr Then
366             Call WriteConsoleMsg(Userindex, "Le has restaurado " & da�o & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
368             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha restaurado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Else
370             Call WriteConsoleMsg(Userindex, "Te has restaurado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
372         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(da�o, UserList(tempChr).Char.CharIndex, &HFF00))
            Call WriteUpdateHP(tempChr)
    
374         b = True
376     ElseIf Hechizos(h).SubeHP = 2 Then
    
378         If Userindex = tempChr Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
380             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

382         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
384         da�o = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
    
386         da�o = da�o + Porcentaje(da�o, 3 * UserList(Userindex).Stats.ELV)
    
            '
            ' If Hechizos(H).StaffAffected Then
            '     If UserList(UserIndex).clase = eClass.Mage Then
            '         If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
            '             da�o = (da�o * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            '         Else
            '             da�o = da�o * 0.7 'Baja da�o a 70% del original
            '         End If
            '     End If
            ' End If
    
            'If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDMAGICO Or UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
            '    da�o = da�o * 1.04  'laud magico de los bardos
            'End If
    
            'cascos antimagia
            'If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
            '    da�o = da�o - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
            'End If
    
            'anillos
            ' If (UserList(tempChr).Invent.AnilloEqpObjIndex > 0) Then
            '    da�o = da�o - RandomNumber(ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
            'End If

388         If UserList(Userindex).flags.Da�oMagico > 0 Then
390             da�o = da�o + Porcentaje(da�o, UserList(Userindex).flags.Da�oMagico)

            End If

392         If da�o < 0 Then da�o = 0
    
394         If Userindex <> tempChr Then
396             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
398         enviarInfoHechizo = True

            'Defensa Resistencia magica
408         If UserList(tempChr).flags.ResistenciaMagica > 0 And Hechizos(h).AntiRm = 0 Then
410             da�o = da�o - Porcentaje(da�o, UserList(tempChr).flags.ResistenciaMagica)

            End If
    
416         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp - da�o
    
418         Call WriteConsoleMsg(Userindex, "Le has quitado " & da�o & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
420         Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha quitado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    
422         Call SubirSkill(tempChr, Resistencia)
    
424         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(da�o, UserList(tempChr).Char.CharIndex))

            'Muere
426         If UserList(tempChr).Stats.MinHp < 1 Then
                'Store it!
428             Call Statistics.StoreFrag(Userindex, tempChr)
430             Call ContarMuerte(tempChr, Userindex)
432             UserList(tempChr).Stats.MinHp = 0
434             Call ActStats(tempChr, Userindex)

                '  Call UserDie(tempChr)
            End If
            
            Call WriteUpdateHP(tempChr)
    
436         b = True

        End If

        'Mana
438     If Hechizos(h).SubeMana = 1 Then
    
440         enviarInfoHechizo = True
442         UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + da�o

444         If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
446         If Userindex <> tempChr Then
448             Call WriteConsoleMsg(Userindex, "Le has restaurado " & da�o & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
450             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha restaurado " & da�o & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
452             Call WriteConsoleMsg(Userindex, "Te has restaurado " & da�o & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
            Call WriteUpdateMana(tempChr)
    
454         b = True
    
456     ElseIf Hechizos(h).SubeMana = 2 Then

458         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
460         If Userindex <> tempChr Then
462             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
464         enviarInfoHechizo = True
    
466         If Userindex <> tempChr Then
468             Call WriteConsoleMsg(Userindex, "Le has quitado " & da�o & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
470             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha quitado " & da�o & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
472             Call WriteConsoleMsg(Userindex, "Te has quitado " & da�o & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
474         UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - da�o

476         If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0

            Call WriteUpdateMana(tempChr)

478         b = True
    
        End If

        'Stamina
480     If Hechizos(h).SubeSta = 1 Then
482         Call InfoHechizo(Userindex)
484         UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + da�o

486         If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta

488         If Userindex <> tempChr Then
490             Call WriteConsoleMsg(Userindex, "Le has restaurado " & da�o & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
492             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha restaurado " & da�o & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            Else
494             Call WriteConsoleMsg(Userindex, "Te has restaurado " & da�o & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
            Call WriteUpdateSta(tempChr)

496         b = True
498     ElseIf Hechizos(h).SubeSta = 2 Then

500         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
502         If Userindex <> tempChr Then
504             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
506         enviarInfoHechizo = True
    
508         If Userindex <> tempChr Then
510             Call WriteConsoleMsg(Userindex, "Le has quitado " & da�o & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
512             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha quitado " & da�o & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            Else
514             Call WriteConsoleMsg(Userindex, "Te has quitado " & da�o & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
516         UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - da�o
    
518         If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0

            Call WriteUpdateSta(tempChr)

520         b = True

        End If

522     If enviarInfoHechizo Then
524         Call InfoHechizo(Userindex)

        End If

    

        
        Exit Sub

HechizoPropUsuario_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoPropUsuario", Erl)
        Resume Next
        
End Sub

Sub HechizoCombinados(ByVal Userindex As Integer, ByRef b As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 02/01/2008
        '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
        '***************************************************
        
        On Error GoTo HechizoCombinados_Err
        

        Dim h As Integer

        Dim da�o As Integer

        Dim tempChr           As Integer

        Dim enviarInfoHechizo As Boolean

100     enviarInfoHechizo = False
    
102     h = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)
104     tempChr = UserList(Userindex).flags.TargetUser
      
        ' <-------- Agilidad ---------->
106     If Hechizos(h).SubeAgilidad = 1 Then
    
            'Para poder tirar cl a un pk en el ring
108         If (TriggerZonaPelea(Userindex, tempChr) <> TRIGGER6_PERMITE) Then
110             If Status(tempChr) = 0 And Status(Userindex) = 1 Or Status(tempChr) = 2 And Status(Userindex) = 1 Then
112                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
114                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
116                     b = False
                        Exit Sub

                    End If

118                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
120                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
122                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
124         enviarInfoHechizo = True
126         da�o = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    
128         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
            'UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + da�o
            'If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then _
             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)

130         UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + da�o

132         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
        
134         UserList(tempChr).flags.TomoPocion = True
136         b = True
138         Call WriteFYA(tempChr)
    
140     ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
142         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
144         If Userindex <> tempChr Then
146             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
148         enviarInfoHechizo = True
    
150         UserList(tempChr).flags.TomoPocion = True
152         da�o = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
154         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

156         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - da�o < 6 Then
158             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
            Else
160             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - da�o

            End If

            'If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
162         b = True
164         Call WriteFYA(tempChr)

        End If

        ' <-------- Fuerza ---------->
166     If Hechizos(h).SubeFuerza = 1 Then

            'Para poder tirar fuerza a un pk en el ring
168         If (TriggerZonaPelea(Userindex, tempChr) <> TRIGGER6_PERMITE) Then
170             If Status(tempChr) = 0 And Status(Userindex) = 1 Or Status(tempChr) = 2 And Status(Userindex) = 1 Then
172                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
174                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
176                     b = False
                        Exit Sub

                    End If

178                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
180                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
182                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
184         da�o = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
186         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
      
188         UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + da�o

190         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
    
192         UserList(tempChr).flags.TomoPocion = True
194         b = True
    
196         enviarInfoHechizo = True
198         Call WriteFYA(tempChr)
200     ElseIf Hechizos(h).SubeFuerza = 2 Then

202         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
204         If Userindex <> tempChr Then
206             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
208         UserList(tempChr).flags.TomoPocion = True
    
210         da�o = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
212         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
        
214         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - da�o < 6 Then
216             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
            Else
218             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - da�o

            End If
   
220         b = True
222         enviarInfoHechizo = True
224         Call WriteFYA(tempChr)

        End If

        'Salud
226     If Hechizos(h).SubeHP = 1 Then
    
            'Verifica que el usuario no este muerto
228         If UserList(tempChr).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "�Est� muerto!", FontTypeNames.FONTTYPE_INFO)
230             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
232             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar a un pk en el ring
234         If (TriggerZonaPelea(Userindex, tempChr) <> TRIGGER6_PERMITE) Then
236             If Status(tempChr) = 0 And Status(Userindex) = 1 Or Status(tempChr) = 2 And Status(Userindex) = 1 Then
238                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
240                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
242                     b = False
                        Exit Sub

                    End If

244                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
246                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
248                     b = False
                        Exit Sub
                    Else

                        'Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
       
250         da�o = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
            'da�o = da�o + Porcentaje(da�o, 3 * UserList(UserIndex).Stats.ELV)
    
252         enviarInfoHechizo = True

254         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp + da�o

256         If UserList(tempChr).Stats.MinHp > UserList(tempChr).Stats.MaxHp Then UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp
    
258         If Userindex <> tempChr Then
260             Call WriteConsoleMsg(Userindex, "Le has restaurado " & da�o & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
262             Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha restaurado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Else
264             Call WriteConsoleMsg(Userindex, "Te has restaurado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
266         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(da�o, UserList(tempChr).Char.CharIndex, &HFF00))
    
268         b = True
270     ElseIf Hechizos(h).SubeHP = 2 Then
    
272         If Userindex = tempChr Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
274             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
276         da�o = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
    
278         da�o = da�o + Porcentaje(da�o, 3 * UserList(Userindex).Stats.ELV)
            '
            ' If Hechizos(H).StaffAffected Then
            '     If UserList(UserIndex).clase = eClass.Mage Then
            '         If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
            '             da�o = (da�o * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            '         Else
            '             da�o = da�o * 0.7 'Baja da�o a 70% del original
            '         End If
            '     End If
            ' End If
    
            'If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDMAGICO Or UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
            '    da�o = da�o * 1.04  'laud magico de los bardos
            'End If
    
            'cascos antimagia
            'If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
            '    da�o = da�o - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
            'End If
    
            'anillos
            ' If (UserList(tempChr).Invent.AnilloEqpObjIndex > 0) Then
            '    da�o = da�o - RandomNumber(ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
            'End If

280         If UserList(Userindex).flags.Da�oMagico > 0 Then
282             da�o = da�o + Porcentaje(da�o, UserList(Userindex).flags.Da�oMagico)

            End If

284         If da�o < 0 Then da�o = 0
    
286         If Not PuedeAtacar(Userindex, tempChr) Then Exit Sub
    
288         If Userindex <> tempChr Then
290             Call UsuarioAtacadoPorUsuario(Userindex, tempChr)

            End If
    
292         enviarInfoHechizo = True
        
302         If UserList(tempChr).flags.ResistenciaMagica > 0 And Hechizos(h).AntiRm = 0 Then
304             da�o = da�o - Porcentaje(da�o, UserList(tempChr).flags.ResistenciaMagica)
            End If
        
            'Resistencia Magica By ladder
    
306         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp - da�o
    
308         Call WriteConsoleMsg(Userindex, "Le has quitado " & da�o & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
310         Call WriteConsoleMsg(tempChr, UserList(Userindex).name & " te ha quitado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
312         Call SubirSkill(tempChr, Resistencia)
314         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(da�o, UserList(tempChr).Char.CharIndex))

            'Muere
316         If UserList(tempChr).Stats.MinHp < 1 Then
                'Store it!
318             Call Statistics.StoreFrag(Userindex, tempChr)
        
320             Call ContarMuerte(tempChr, Userindex)
322             UserList(tempChr).Stats.MinHp = 0
324             Call ActStats(tempChr, Userindex)

                'Call UserDie(tempChr)
            End If
    
326         b = True

        End If

        Dim tU As Integer

328     tU = tempChr

330     If Hechizos(h).Invisibilidad = 1 Then
   
332         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "�Est� muerto!", FontTypeNames.FONTTYPE_INFO)
334             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
336             b = False
                Exit Sub

            End If
    
338         If UserList(tU).Counters.Saliendo Then
340             If Userindex <> tU Then
342                 Call WriteConsoleMsg(Userindex, "�El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
344                 b = False
                    Exit Sub
                Else
346                 Call WriteConsoleMsg(Userindex, "�No pod�s ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
348                 b = False
                    Exit Sub

                End If

            End If
    
            'Para poder tirar invi a un pk en el ring
350         If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
352             If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
354                 If esArmada(Userindex) Then
356                     Call WriteConsoleMsg(Userindex, "Los miembros de la armada real no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
358                     b = False
                        Exit Sub

                    End If

360                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
362                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
364                     b = False
                        Exit Sub
                    Else
366                     Call VolverCriminal(Userindex)

                    End If

                End If

            End If
    
            'Si sos user, no uses este hechizo con GMS.
368         If UserList(Userindex).flags.Privilegios And PlayerType.user Then
370             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
   
372         UserList(tU).flags.invisible = 1
            'Ladder
            'Reseteamos el contador de Invisibilidad
374         UserList(tU).Counters.Invisibilidad = Hechizos(h).Duration
376         Call WriteContadores(tU)
378         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

380         enviarInfoHechizo = True
382         b = True

        End If

384     If Hechizos(h).Envenena > 0 Then
            ' If UserIndex = tU Then
            '    Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            '   Exit Sub
            'End If
    
386         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
388         If Userindex <> tU Then
390             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

392         UserList(tU).flags.Envenenado = Hechizos(h).Envenena
394         enviarInfoHechizo = True
396         b = True

        End If

398     If Hechizos(h).desencantar = 1 Then
400         Call WriteConsoleMsg(Userindex, "Has sido desencantado.", FontTypeNames.FONTTYPE_INFO)

402         UserList(Userindex).flags.Envenenado = 0
404         UserList(Userindex).flags.Incinerado = 0
    
406         If UserList(Userindex).flags.Inmovilizado = 1 Then
408             UserList(Userindex).Counters.Inmovilizado = 0
410             UserList(Userindex).flags.Inmovilizado = 0
412             Call WriteInmovilizaOK(Userindex)
            

            End If
    
414         If UserList(Userindex).flags.Paralizado = 1 Then
416             UserList(Userindex).flags.Paralizado = 0
418             Call WriteParalizeOK(Userindex)
            
           
            End If
        
420         If UserList(Userindex).flags.Ceguera = 1 Then
422             UserList(Userindex).flags.Ceguera = 0
424             Call WriteBlindNoMore(Userindex)
            

            End If
    
426         If UserList(Userindex).flags.Maldicion = 1 Then
428             UserList(Userindex).flags.Maldicion = 0
430             UserList(Userindex).Counters.Maldicion = 0

            End If
    
432         enviarInfoHechizo = True
434         b = True

        End If

436     If Hechizos(h).Sanacion = 1 Then

438         UserList(tU).flags.Envenenado = 0
440         UserList(tU).flags.Incinerado = 0
442         enviarInfoHechizo = True
444         b = True

        End If

446     If Hechizos(h).incinera = 1 Then
448         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
450             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
452         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
454         If Userindex <> tU Then
456             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

458         UserList(tU).flags.Incinerado = 1
460         enviarInfoHechizo = True
462         b = True

        End If

464     If Hechizos(h).CuraVeneno = 1 Then

            'Verificamos que el usuario no este muerto
466         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "�Est� muerto!", FontTypeNames.FONTTYPE_INFO)
468             Call WriteLocaleMsg(Userindex, "77", FontTypeNames.FONTTYPE_INFO)
470             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar veneno a un pk en el ring
472         If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
474             If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
476                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
478                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
480                     b = False
                        Exit Sub

                    End If

482                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
484                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
486                     b = False
                        Exit Sub
                    Else

                        '    Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
        
            'Si sos user, no uses este hechizo con GMS.
488         If UserList(Userindex).flags.Privilegios And PlayerType.user Then
490             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
        
492         UserList(tU).flags.Envenenado = 0
494         enviarInfoHechizo = True
496         b = True

        End If

498     If Hechizos(h).Maldicion = 1 Then
500         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
502             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
504         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
506         If Userindex <> tU Then
508             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

510         UserList(tU).flags.Maldicion = 1
512         UserList(tU).Counters.Maldicion = 200
    
514         enviarInfoHechizo = True
516         b = True

        End If

518     If Hechizos(h).RemoverMaldicion = 1 Then
520         UserList(tU).flags.Maldicion = 0
522         enviarInfoHechizo = True
524         b = True

        End If

526     If Hechizos(h).GolpeCertero = 1 Then
528         UserList(tU).flags.GolpeCertero = 1
530         enviarInfoHechizo = True
532         b = True

        End If

534     If Hechizos(h).Bendicion = 1 Then
536         UserList(tU).flags.Bendicion = 1
538         enviarInfoHechizo = True
540         b = True

        End If

542     If Hechizos(h).Paraliza = 1 Then
544         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
546             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
548         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
            
550         If Userindex <> tU Then
552             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If
            
554         enviarInfoHechizo = True
556         b = True

558         If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
560             Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
562             Call WriteConsoleMsg(Userindex, " �El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            
                Exit Sub

            End If
            
564         UserList(tU).Counters.Paralisis = Hechizos(h).Duration

566         If UserList(tU).flags.Paralizado = 0 Then
568             UserList(tU).flags.Paralizado = 1
570             Call WriteParalizeOK(tU)
572             Call WritePosUpdate(tU)
            End If

        End If

574     If Hechizos(h).Inmoviliza = 1 Then
576         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
578             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
580         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
            
582         If Userindex <> tU Then
584             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If
            
586         enviarInfoHechizo = True
588         b = True
            
590         UserList(tU).Counters.Inmovilizado = Hechizos(h).Duration

592         If UserList(tU).flags.Inmovilizado = 0 Then
594             UserList(tU).flags.Inmovilizado = 1
596             Call WriteInmovilizaOK(tU)
598             Call WritePosUpdate(tU)
            

            End If

        End If

600     If Hechizos(h).RemoverParalisis = 1 Then
        
            'Para poder tirar remo a un pk en el ring
602         If (TriggerZonaPelea(Userindex, tU) <> TRIGGER6_PERMITE) Then
604             If Status(tU) = 0 And Status(Userindex) = 1 Or Status(tU) = 2 And Status(Userindex) = 1 Then
606                 If esArmada(Userindex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
608                     Call WriteLocaleMsg(Userindex, "379", FontTypeNames.FONTTYPE_INFO)
610                     b = False
                        Exit Sub

                    End If

612                 If UserList(Userindex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volver�s criminal como ellos", FontTypeNames.FONTTYPE_INFO)
614                     Call WriteLocaleMsg(Userindex, "378", FontTypeNames.FONTTYPE_INFO)
616                     b = False
                        Exit Sub
                    Else
618                     Call VolverCriminal(Userindex)

                    End If

                End If
            
            End If

620         If UserList(tU).flags.Inmovilizado = 1 Then
622             UserList(tU).Counters.Inmovilizado = 0
624             UserList(tU).flags.Inmovilizado = 0
626             Call WriteInmovilizaOK(tU)
628             enviarInfoHechizo = True
            
630             b = True

            End If

632         If UserList(tU).flags.Paralizado = 1 Then
634             UserList(tU).flags.Paralizado = 0
                'no need to crypt this
636             Call WriteParalizeOK(tU)
638             enviarInfoHechizo = True
            
640             b = True

            End If

        End If

642     If Hechizos(h).Ceguera = 1 Then
644         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
646             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
648         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
650         If Userindex <> tU Then
652             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

654         UserList(tU).flags.Ceguera = 1
656         UserList(tU).Counters.Ceguera = Hechizos(h).Duration

658         Call WriteBlind(tU)
        
660         enviarInfoHechizo = True
662         b = True

        End If

664     If Hechizos(h).Estupidez = 1 Then
666         If Userindex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No pod�s atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
668             Call WriteLocaleMsg(Userindex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

670         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
672         If Userindex <> tU Then
674             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If

676         If UserList(tU).flags.Estupidez = 0 Then
678             UserList(tU).flags.Estupidez = 1
680             UserList(tU).Counters.Estupidez = Hechizos(h).Duration

            End If

682         Call WriteDumb(tU)
        

684         enviarInfoHechizo = True
686         b = True

        End If

688     If Hechizos(h).Velocidad > 0 Then

690         If Not PuedeAtacar(Userindex, tU) Then Exit Sub
            
692         If Userindex <> tU Then
694             Call UsuarioAtacadoPorUsuario(Userindex, tU)

            End If
            
696         enviarInfoHechizo = True
698         b = True
            
700         If UserList(tU).Counters.Velocidad = 0 Then
702             UserList(tU).flags.VelocidadBackup = UserList(tU).Char.speeding

            End If

704         UserList(tU).Char.speeding = Hechizos(h).Velocidad
706         Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSpeedingACT(UserList(tU).Char.CharIndex, UserList(tU).Char.speeding))
            
708         UserList(tU).Counters.Velocidad = Hechizos(h).Duration

        End If

710     If enviarInfoHechizo Then
712         Call InfoHechizo(Userindex)

        End If

    

        
        Exit Sub

HechizoCombinados_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoCombinados", Erl)
        Resume Next
        
End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal Userindex As Integer, ByVal Slot As Byte)
        
        On Error GoTo UpdateUserHechizos_Err
        

        'Call LogTarea("Sub UpdateUserHechizos")

        Dim LoopC As Byte

        'Actualiza un solo slot
100     If Not UpdateAll Then

            'Actualiza el inventario
102         If UserList(Userindex).Stats.UserHechizos(Slot) > 0 Then
104             Call ChangeUserHechizo(Userindex, Slot, UserList(Userindex).Stats.UserHechizos(Slot))
            Else
106             Call ChangeUserHechizo(Userindex, Slot, 0)

            End If

        Else

            'Actualiza todos los slots
108         For LoopC = 1 To MAXUSERHECHIZOS

                'Actualiza el inventario
110             If UserList(Userindex).Stats.UserHechizos(LoopC) > 0 Then
112                 Call ChangeUserHechizo(Userindex, LoopC, UserList(Userindex).Stats.UserHechizos(LoopC))
                Else
114                 Call ChangeUserHechizo(Userindex, LoopC, 0)

                End If

116         Next LoopC

        End If

        
        Exit Sub

UpdateUserHechizos_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.UpdateUserHechizos", Erl)
        Resume Next
        
End Sub

Sub ChangeUserHechizo(ByVal Userindex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)
        
        On Error GoTo ChangeUserHechizo_Err
        

        'Call LogTarea("ChangeUserHechizo")
    
100     UserList(Userindex).Stats.UserHechizos(Slot) = Hechizo
    
102     If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
104         Call WriteChangeSpellSlot(Userindex, Slot)
        Else
106         Call WriteChangeSpellSlot(Userindex, Slot)

        End If

        
        Exit Sub

ChangeUserHechizo_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.ChangeUserHechizo", Erl)
        Resume Next
        
End Sub

Public Sub DesplazarHechizo(ByVal Userindex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)
        
        On Error GoTo DesplazarHechizo_Err
        

100     If (Dire <> 1 And Dire <> -1) Then Exit Sub
102     If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

        Dim TempHechizo As Integer

104     If Dire = 1 Then 'Mover arriba
106         If CualHechizo = 1 Then
108             Call WriteConsoleMsg(Userindex, "No pod�s mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
110             TempHechizo = UserList(Userindex).Stats.UserHechizos(CualHechizo)
112             UserList(Userindex).Stats.UserHechizos(CualHechizo) = UserList(Userindex).Stats.UserHechizos(CualHechizo - 1)
114             UserList(Userindex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

                'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
116             If UserList(Userindex).flags.Hechizo > 0 Then
118                 UserList(Userindex).flags.Hechizo = UserList(Userindex).flags.Hechizo - 1

                End If

            End If

        Else 'mover abajo

120         If CualHechizo = MAXUSERHECHIZOS Then
122             Call WriteConsoleMsg(Userindex, "No pod�s mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
124             TempHechizo = UserList(Userindex).Stats.UserHechizos(CualHechizo)
126             UserList(Userindex).Stats.UserHechizos(CualHechizo) = UserList(Userindex).Stats.UserHechizos(CualHechizo + 1)
128             UserList(Userindex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

                'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
130             If UserList(Userindex).flags.Hechizo > 0 Then
132                 UserList(Userindex).flags.Hechizo = UserList(Userindex).flags.Hechizo + 1

                End If

            End If

        End If

        
        Exit Sub

DesplazarHechizo_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.DesplazarHechizo", Erl)
        Resume Next
        
End Sub

Sub AreaHechizo(Userindex As Integer, NpcIndex As Integer, X As Byte, Y As Byte, npc As Boolean)
        
        On Error GoTo AreaHechizo_Err
        

        Dim calculo      As Integer

        Dim TilesDifUser As Integer

        Dim TilesDifNpc  As Integer

        Dim tilDif       As Integer

        Dim h2           As Integer

        Dim Hit          As Integer

        Dim da�o As Integer

        Dim porcentajeDesc As Integer

100     h2 = UserList(Userindex).Stats.UserHechizos(UserList(Userindex).flags.Hechizo)

        'Calculo de descuesto de golpe por cercania.
102     TilesDifUser = X + Y

104     If npc Then
106         If Hechizos(h2).SubeHP = 2 Then
108             TilesDifNpc = Npclist(NpcIndex).Pos.X + Npclist(NpcIndex).Pos.Y
            
110             tilDif = TilesDifUser - TilesDifNpc
            
112             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

114             If tilDif <> 0 Then
116                 porcentajeDesc = Abs(tilDif) * 20
118                 da�o = Hit / 100 * porcentajeDesc
120                 da�o = Hit - da�o
                Else
122                 da�o = Hit

                End If
            
124             If UserList(Userindex).flags.Da�oMagico > 0 Then
126                 da�o = da�o + Porcentaje(da�o, UserList(Userindex).flags.Da�oMagico)

                End If
            
128             Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp - da�o
            
130             If UserList(Userindex).ChatCombate = 1 Then
132                 Call WriteConsoleMsg(Userindex, "Le has causado " & da�o & " puntos de da�o a " & Npclist(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)

                End If
            
134             Call CalcularDarExp(Userindex, NpcIndex, da�o)
                
136             If Npclist(NpcIndex).Stats.MinHp <= 0 Then
                    'UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + Npclist(NpcIndex).GiveEXP
                    'UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Npclist(NpcIndex).GiveGLD
138                 Call MuereNpc(NpcIndex, Userindex)

                End If

                Exit Sub

            End If

        Else

140         TilesDifNpc = UserList(NpcIndex).Pos.X + UserList(NpcIndex).Pos.Y
142         tilDif = TilesDifUser - TilesDifNpc

144         If Hechizos(h2).SubeHP = 2 Then
146             If Userindex = NpcIndex Then
                    Exit Sub

                End If

148             If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
                
150             If Userindex <> NpcIndex Then
152                 Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

                End If
                
154             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

156             If tilDif <> 0 Then
158                 porcentajeDesc = Abs(tilDif) * 20
160                 da�o = Hit / 100 * porcentajeDesc
162                 da�o = Hit - da�o
                Else
164                 da�o = Hit

                End If
                        
166             If UserList(Userindex).flags.Da�oMagico > 0 Then
168                 da�o = da�o + Porcentaje(da�o, UserList(Userindex).flags.Da�oMagico)

                End If
                        
180             UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MinHp - da�o
                    
182             Call WriteConsoleMsg(Userindex, "Le has quitado " & da�o & " puntos de vida a " & UserList(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)
184             Call WriteConsoleMsg(NpcIndex, UserList(Userindex).name & " te ha quitado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
186             Call SubirSkill(NpcIndex, Resistencia)
188             Call WriteUpdateUserStats(NpcIndex)
                
                'Muere
190             If UserList(NpcIndex).Stats.MinHp < 1 Then
                    'Store it!
192                 Call Statistics.StoreFrag(Userindex, NpcIndex)
                        
194                 Call ContarMuerte(NpcIndex, Userindex)
196                 UserList(NpcIndex).Stats.MinHp = 0
198                 Call ActStats(NpcIndex, Userindex)

                    'Call UserDie(NpcIndex)
                End If

            End If
                
200         If Hechizos(h2).SubeHP = 1 Then
202             If (TriggerZonaPelea(Userindex, NpcIndex) <> TRIGGER6_PERMITE) Then
204                 If Status(Userindex) = 1 And Status(NpcIndex) <> 1 Then
                        Exit Sub

                    End If

                End If

206             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

208             If tilDif <> 0 Then
210                 porcentajeDesc = Abs(tilDif) * 20
212                 da�o = Hit / 100 * porcentajeDesc
214                 da�o = Hit - da�o
                Else
216                 da�o = Hit

                End If
 
218             UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MinHp + da�o

220             If UserList(NpcIndex).Stats.MinHp > UserList(NpcIndex).Stats.MaxHp Then UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MaxHp

            End If
 
222         If Userindex <> NpcIndex Then
224             Call WriteConsoleMsg(Userindex, "Le has restaurado " & da�o & " puntos de vida a " & UserList(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)
226             Call WriteConsoleMsg(NpcIndex, UserList(Userindex).name & " te ha restaurado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Else
228             Call WriteConsoleMsg(Userindex, "Te has restaurado " & da�o & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

            End If
                    
230         Call WriteUpdateUserStats(NpcIndex)

        End If
                
232     If Hechizos(h2).Envenena > 0 Then
234         If Userindex = NpcIndex Then
                Exit Sub

            End If
                    
236         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
                
238         If Userindex <> NpcIndex Then
240             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

            End If
                    
242         UserList(NpcIndex).flags.Envenenado = Hechizos(h2).Envenena
244         Call WriteConsoleMsg(NpcIndex, UserList(Userindex).name & " te ha envenenado.", FontTypeNames.FONTTYPE_FIGHT)

        End If
                
246     If Hechizos(h2).Paraliza = 1 Then
248         If Userindex = NpcIndex Then
                Exit Sub

            End If
    
250         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
            
252         If Userindex <> NpcIndex Then
254             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

            End If
            
256         Call WriteConsoleMsg(NpcIndex, "Has sido paralizado.", FontTypeNames.FONTTYPE_INFO)
258         UserList(NpcIndex).Counters.Paralisis = Hechizos(h2).Duration

260         If UserList(NpcIndex).flags.Paralizado = 0 Then
262             UserList(NpcIndex).flags.Paralizado = 1
264             Call WriteParalizeOK(NpcIndex)
            

            End If
            
        End If
                
266     If Hechizos(h2).Inmoviliza = 1 Then
268         If Userindex = NpcIndex Then
                Exit Sub

            End If
    
270         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
            
272         If Userindex <> NpcIndex Then
274             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

            End If
                    
276         Call WriteConsoleMsg(NpcIndex, "Has sido inmovilizado.", FontTypeNames.FONTTYPE_INFO)
278         UserList(NpcIndex).Counters.Inmovilizado = Hechizos(h2).Duration

280         If UserList(NpcIndex).flags.Inmovilizado = 0 Then
282             UserList(NpcIndex).flags.Inmovilizado = 1
284             Call WriteInmovilizaOK(NpcIndex)
286             Call WritePosUpdate(NpcIndex)
            
            End If

        End If
                
288     If Hechizos(h2).Ceguera = 1 Then
290         If Userindex = NpcIndex Then
                Exit Sub

            End If
    
292         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
            
294         If Userindex <> NpcIndex Then
296             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

            End If
                    
298         UserList(NpcIndex).flags.Ceguera = 1
300         UserList(NpcIndex).Counters.Ceguera = Hechizos(h2).Duration
302         Call WriteConsoleMsg(NpcIndex, "Te han cegado.", FontTypeNames.FONTTYPE_INFO)
            
304         Call WriteBlind(NpcIndex)
        

        End If
                
306     If Hechizos(h2).Velocidad > 0 Then
    
308         If Userindex = NpcIndex Then
                Exit Sub

            End If
    
310         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
            
312         If Userindex <> NpcIndex Then
314             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

            End If

316         If UserList(NpcIndex).Counters.Velocidad = 0 Then
318             UserList(NpcIndex).flags.VelocidadBackup = UserList(NpcIndex).Char.speeding

            End If

320         UserList(NpcIndex).Char.speeding = Hechizos(h2).Velocidad
322         Call SendData(SendTarget.ToPCArea, NpcIndex, PrepareMessageSpeedingACT(UserList(NpcIndex).Char.CharIndex, UserList(NpcIndex).Char.speeding))
324         UserList(NpcIndex).Counters.Velocidad = Hechizos(h2).Duration

        End If
                
326     If Hechizos(h2).Maldicion = 1 Then
328         If Userindex = NpcIndex Then
                Exit Sub

            End If
    
330         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
            
332         If Userindex <> NpcIndex Then
334             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

            End If

336         Call WriteConsoleMsg(NpcIndex, "Ahora estas maldito. No podras Atacar", FontTypeNames.FONTTYPE_INFO)
338         UserList(NpcIndex).flags.Maldicion = 1
340         UserList(NpcIndex).Counters.Maldicion = Hechizos(h2).Duration

        End If
                
342     If Hechizos(h2).RemoverMaldicion = 1 Then
344         Call WriteConsoleMsg(NpcIndex, "Te han removido la maldicion.", FontTypeNames.FONTTYPE_INFO)
346         UserList(NpcIndex).flags.Maldicion = 0

        End If
                
348     If Hechizos(h2).GolpeCertero = 1 Then
350         Call WriteConsoleMsg(NpcIndex, "Tu proximo golpe sera certero.", FontTypeNames.FONTTYPE_INFO)
352         UserList(NpcIndex).flags.GolpeCertero = 1

        End If
                
354     If Hechizos(h2).Bendicion = 1 Then
356         Call WriteConsoleMsg(NpcIndex, "Has sido bendecido.", FontTypeNames.FONTTYPE_INFO)
358         UserList(NpcIndex).flags.Bendicion = 1

        End If
                  
360     If Hechizos(h2).incinera = 1 Then
362         If Userindex = NpcIndex Then
                Exit Sub

            End If
    
364         If Not PuedeAtacar(Userindex, NpcIndex) Then Exit Sub
            
366         If Userindex <> NpcIndex Then
368             Call UsuarioAtacadoPorUsuario(Userindex, NpcIndex)

            End If

370         UserList(NpcIndex).flags.Incinerado = 1
372         Call WriteConsoleMsg(NpcIndex, "Has sido Incinerado.", FontTypeNames.FONTTYPE_INFO)

        End If
                
374     If Hechizos(h2).Invisibilidad = 1 Then
376         Call WriteConsoleMsg(NpcIndex, "Ahora sos invisible.", FontTypeNames.FONTTYPE_INFO)
378         UserList(NpcIndex).flags.invisible = 1
380         UserList(NpcIndex).Counters.Invisibilidad = Hechizos(h2).Duration
382         Call WriteContadores(NpcIndex)
384         Call SendData(SendTarget.ToPCArea, NpcIndex, PrepareMessageSetInvisible(UserList(NpcIndex).Char.CharIndex, True))

        End If
                              
386     If Hechizos(h2).Sanacion = 1 Then
388         Call WriteConsoleMsg(NpcIndex, "Has sido sanado.", FontTypeNames.FONTTYPE_INFO)
390         UserList(NpcIndex).flags.Envenenado = 0
392         UserList(NpcIndex).flags.Incinerado = 0

        End If
                
394     If Hechizos(h2).RemoverParalisis = 1 Then
396         Call WriteConsoleMsg(NpcIndex, "Has sido removido.", FontTypeNames.FONTTYPE_INFO)

398         If UserList(NpcIndex).flags.Inmovilizado = 1 Then
400             UserList(NpcIndex).Counters.Inmovilizado = 0
402             UserList(NpcIndex).flags.Inmovilizado = 0
404             Call WriteInmovilizaOK(NpcIndex)
            

            End If

406         If UserList(NpcIndex).flags.Paralizado = 1 Then
408             UserList(NpcIndex).flags.Paralizado = 0
                'no need to crypt this
410             Call WriteParalizeOK(NpcIndex)
            

            End If

        End If
                
412     If Hechizos(h2).desencantar = 1 Then
414         Call WriteConsoleMsg(NpcIndex, "Has sido desencantado.", FontTypeNames.FONTTYPE_INFO)
                    
416         UserList(NpcIndex).flags.Envenenado = 0
418         UserList(NpcIndex).flags.Incinerado = 0
                    
420         If UserList(NpcIndex).flags.Inmovilizado = 1 Then
422             UserList(NpcIndex).Counters.Inmovilizado = 0
424             UserList(NpcIndex).flags.Inmovilizado = 0
426             Call WriteInmovilizaOK(NpcIndex)
            

            End If
                    
428         If UserList(NpcIndex).flags.Paralizado = 1 Then
430             UserList(NpcIndex).flags.Paralizado = 0
432             Call WriteParalizeOK(NpcIndex)
            
                       
            End If
                    
434         If UserList(NpcIndex).flags.Ceguera = 1 Then
436             UserList(NpcIndex).flags.Ceguera = 0
438             Call WriteBlindNoMore(NpcIndex)
            

            End If
                    
440         If UserList(NpcIndex).flags.Maldicion = 1 Then
442             UserList(NpcIndex).flags.Maldicion = 0
444             UserList(NpcIndex).Counters.Maldicion = 0

            End If

        End If
        
        
        Exit Sub

AreaHechizo_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.AreaHechizo", Erl)
        Resume Next
        
End Sub
