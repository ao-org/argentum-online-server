Attribute VB_Name = "modHechizos"
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
'for more information about ORE please visit http://www.baronsoft.com/C
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

Public Const HELEMENTAL_FUEGO  As Integer = 26

Public Const HELEMENTAL_TIERRA As Integer = 28

Public Const SUPERANILLO       As Integer = 700

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal Spell As Integer, Optional ByVal IgnoreVisibilityCheck As Boolean = False)
    'Guardia caos
        
    On Error GoTo NpcLanzaSpellSobreUser_Err
        
    With UserList(UserIndex)
        
        '¿NPC puede ver a través de la invisibilidad?
        If Not IgnoreVisibilityCheck Then
            If .flags.invisible = 1 Or .flags.Oculto = 1 Or .flags.Inmunidad = 1 Then Exit Sub
        End If

        'Npclist(NpcIndex).CanAttack = 0
        Dim daño As Integer

        If Hechizos(Spell).SubeHP = 1 Then

            daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.x, .Pos.Y))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

            .Stats.MinHp = .Stats.MinHp + daño

            If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
    
            Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteUpdateHP(UserIndex)
            Call SubirSkill(UserIndex, Resistencia)

        ElseIf Hechizos(Spell).SubeHP = 2 Then
        
            daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
        
            If .Invent.CascoEqpObjIndex > 0 Then
                daño = daño - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)

            End If
        
            If .Invent.AnilloEqpObjIndex > 0 Then
                daño = daño - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)

            End If
        
            If daño < 0 Then daño = 0
        
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.x, .Pos.Y))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageEfectOverHead(daño, .Char.CharIndex))
                
            If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?

                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))
          
            End If

            .Stats.MinHp = .Stats.MinHp - daño
        
            Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Call SubirSkill(UserIndex, Resistencia)
        
            'Muere
            If .Stats.MinHp < 1 Then
                .Stats.MinHp = 0
                'If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                'restarCriminalidad (UserIndex)
                ' End If
                Call UserDie(UserIndex)
                '[Barrin 1-12-03]
                ' If Npclist(NpcIndex).MaestroUser > 0 Then
                'Store it!
                'Call Statistics.StoreFrag(Npclist(NpcIndex).MaestroUser, UserIndex)
                
                'Call ContarMuerte(UserIndex, Npclist(NpcIndex).MaestroUser)
                ' Call ActStats(UserIndex, Npclist(NpcIndex).MaestroUser)
                '  End If
                '[/Barrin]
            End If
            
            Call WriteUpdateHP(UserIndex)
    
        ElseIf Hechizos(Spell).Paraliza = 1 Then

            If .flags.Paralizado = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.x, .Pos.Y))
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

                .flags.Paralizado = 1
                .Counters.Paralisis = Hechizos(Spell).Duration / 2
          
                Call WriteParalizeOK(UserIndex)
                Call WritePosUpdate(UserIndex)

            End If

        ElseIf Hechizos(Spell).incinera = 1 Then
            Debug.Print "incinerar"

            If .flags.Incinerado = 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(Spell).wav, .Pos.x, .Pos.Y))

                If Hechizos(Spell).Particle > 0 Then '¿Envio Particula?
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, Hechizos(Spell).Particle, Hechizos(Spell).TimeParticula, False))

                End If

                .flags.Incinerado = 1
                Call WriteConsoleMsg(UserIndex, "Has sido incinerado por " & Npclist(NpcIndex).name & ".", FontTypeNames.FONTTYPE_FIGHT)

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

        Dim daño As Integer

104     If Hechizos(Spell).SubeHP = 2 Then
    
106         daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
108         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).wav, Npclist(TargetNPC).Pos.x, Npclist(TargetNPC).Pos.Y))
110         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
111         Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageEfectOverHead(daño, Npclist(TargetNPC).Char.CharIndex))
        
112         Npclist(TargetNPC).Stats.MinHp = Npclist(TargetNPC).Stats.MinHp - daño

            ' Mascotas dan experiencia al amo
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call CalcularDarExp(Npclist(NpcIndex).MaestroUser, TargetNPC, daño)
            End If
        
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

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

    On Error GoTo ErrHandler
    
    Dim j As Integer

    For j = 1 To MAXUSERHECHIZOS

        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function

        End If

    Next

    Exit Function
ErrHandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal slot As Integer)
        
        On Error GoTo AgregarHechizo_Err
        

        Dim hIndex As Integer

        Dim j      As Integer

100     hIndex = ObjData(UserList(UserIndex).Invent.Object(slot).ObjIndex).HechizoIndex

102     If Not TieneHechizo(hIndex, UserIndex) Then

            'Buscamos un slot vacio
104         For j = 1 To MAXUSERHECHIZOS

106             If UserList(UserIndex).Stats.UserHechizos(j) = 0 Then Exit For
108         Next j
        
110         If UserList(UserIndex).Stats.UserHechizos(j) <> 0 Then
112             Call WriteConsoleMsg(UserIndex, "No tenes espacio para mas hechizos.", FontTypeNames.FONTTYPE_INFO)
            Else
114             UserList(UserIndex).Stats.UserHechizos(j) = hIndex
116             Call UpdateUserHechizos(False, UserIndex, CByte(j))
                'Quitamos del inv el item
118             Call QuitarUserInvItem(UserIndex, CByte(slot), 1)

            End If

        Else
120         Call WriteConsoleMsg(UserIndex, "Ya tenes ese hechizo.", FontTypeNames.FONTTYPE_INFO)

        End If

        
        Exit Sub

AgregarHechizo_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.AgregarHechizo", Erl)
        Resume Next
        
End Sub
            
Sub DecirPalabrasMagicas(ByVal Hechizo As Byte, ByVal UserIndex As Integer)

    On Error Resume Next

    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("PMAG*" & Hechizo, UserList(UserIndex).Char.CharIndex, vbCyan))
    Exit Sub

End Sub

Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer, Optional ByVal slot As Integer = 0) As Boolean
        
        On Error GoTo PuedeLanzar_Err
        

100     If UserList(UserIndex).flags.Muerto = 0 Then

            Dim wp2 As WorldPos

102         wp2.Map = UserList(UserIndex).flags.TargetMap
104         wp2.x = UserList(UserIndex).flags.TargetX
106         wp2.Y = UserList(UserIndex).flags.TargetY
    
108         If Hechizos(HechizoIndex).NecesitaObj > 0 Then
110             If TieneObjEnInv(UserIndex, Hechizos(HechizoIndex).NecesitaObj, Hechizos(HechizoIndex).NecesitaObj2) Then
112                 PuedeLanzar = True
                    'Exit Function
               
                Else
114                 Call WriteConsoleMsg(UserIndex, "Necesitas un " & ObjData(Hechizos(HechizoIndex).NecesitaObj).name & " para lanzar el hechizo.", FontTypeNames.FONTTYPE_INFO)
116                 PuedeLanzar = False
                    Exit Function

                End If

            End If
    
118         If Hechizos(HechizoIndex).CoolDown > 0 Then

                Dim actual            As Long

                Dim segundosFaltantes As Long

120             actual = GetTickCount() And &H7FFFFFFF

122             If UserList(UserIndex).Counters.UserHechizosInterval(slot) + (Hechizos(HechizoIndex).CoolDown * 1000) > actual Then
124                 segundosFaltantes = Int((UserList(UserIndex).Counters.UserHechizosInterval(slot) + (Hechizos(HechizoIndex).CoolDown * 1000) - actual) / 1000)
126                 Call WriteConsoleMsg(UserIndex, "Debes esperar " & segundosFaltantes & " segundos para volver a tirar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
128                 PuedeLanzar = False
                    Exit Function
                Else
130                 PuedeLanzar = True

                End If

            End If
    
132         If UserList(UserIndex).Stats.MinHp > Hechizos(HechizoIndex).RequiredHP Then
134             If UserList(UserIndex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
136                 If UserList(UserIndex).Stats.UserSkills(eSkill.magia) >= Hechizos(HechizoIndex).MinSkill Then
138                     If UserList(UserIndex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
140                         PuedeLanzar = True
                        Else
142                         Call WriteLocaleMsg(UserIndex, "93", FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
144                         PuedeLanzar = False

                        End If
                    
                    Else
146                     Call WriteConsoleMsg(UserIndex, "No tenes suficientes puntos de magia para lanzar este hechizo, necesitas " & Hechizos(HechizoIndex).MinSkill & " puntos.", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteLocaleMsg(UserIndex, "221", FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "Necesitas " & Hechizos(HechizoIndex).MinSkill & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
148                     PuedeLanzar = False

                    End If

                Else
150                 Call WriteLocaleMsg(UserIndex, "222", FontTypeNames.FONTTYPE_INFO)
                    'Call WriteConsoleMsg(UserIndex, "No tenes suficiente mana. Necesitas " & Hechizos(HechizoIndex).ManaRequerido & " puntos de mana.", FontTypeNames.FONTTYPE_INFO)
152                 PuedeLanzar = False

                End If

            Else
154             Call WriteConsoleMsg(UserIndex, "No tenes suficiente vida. Necesitas " & Hechizos(HechizoIndex).RequiredHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
156             PuedeLanzar = False

            End If

        Else
            'Call WriteConsoleMsg(UserIndex, "No podes lanzar hechizos porque estas muerto.", FontTypeNames.FONTTYPE_INFO)
158         Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
160         PuedeLanzar = False

        End If

        
        Exit Function

PuedeLanzar_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.PuedeLanzar", Erl)
        Resume Next
        
End Function

''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operación.

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef b As Boolean)
    '***************************************************
    'Author: Uknown
    'Modification: 06/15/2008 (NicoNZ)
    'Last modification: 01/12/2020 (WyroX)
    'Sale del sub si no hay una posición valida.
    '***************************************************
    
    On Error GoTo HechizoInvocacion_Err
    
    With UserList(UserIndex)
    
        Dim h As Integer, j As Integer, ind As Integer, Index As Integer
        Dim TargetPos As WorldPos
    
        TargetPos.Map = .flags.TargetMap
        TargetPos.x = .flags.TargetX
        TargetPos.Y = .flags.TargetY
        
        h = .Stats.UserHechizos(.flags.Hechizo)
    
        If Hechizos(h).Invoca = 1 Then
    
            If .NroMascotas >= MAXMASCOTAS Then Exit Sub
        
            'No deja invocar mas de 1 fatuo
            If Hechizos(h).NumNpc = FUEGOFATUO And .NroMascotas >= 1 Then
                Call WriteConsoleMsg(UserIndex, "Para invocar el fuego fatuo no debes tener otras criaturas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'No permitimos se invoquen criaturas en zonas seguras
            If MapInfo(.Pos.Map).Seguro Or MapData(.Pos.Map, .Pos.x, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
                Call WriteConsoleMsg(UserIndex, "En zona segura no puedes invocar criaturas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            For j = 1 To Hechizos(h).cant
                
                If .NroMascotas < MAXMASCOTAS Then
                    ind = SpawnNpc(Hechizos(h).NumNpc, TargetPos, True, False)
                    If ind > 0 Then
                        .NroMascotas = .NroMascotas + 1
                        
                        Index = FreeMascotaIndex(UserIndex)
                        
                        .MascotasIndex(Index) = ind
                        .MascotasType(Index) = Npclist(ind).Numero
                        
                        Npclist(ind).MaestroUser = UserIndex
                        Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
                        Npclist(ind).GiveGLD = 0
                        
                        Call FollowAmo(ind)
                    Else
                        Exit Sub
                    End If
                        
                Else
                    Exit For
                End If
                
            Next j
            
            Call InfoHechizo(UserIndex)
            b = True
        
        ElseIf Hechizos(h).Invoca = 2 Then
            
            ' Si tiene mascotas
            If .NroMascotas > 0 Then
                ' Tiene que estar en zona insegura
                If Not MapInfo(.Pos.Map).Seguro Then

                    Dim i As Integer
                    
                    ' Si no están guardadas las mascotas
                    If .flags.MascotasGuardadas = 0 Then
                        For i = 1 To MAXMASCOTAS
                            If .MascotasIndex(i) > 0 Then
                                ' Si no es un elemental, lo "guardamos"... lo matamos
                                If Npclist(.MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then
                                    ' Le saco el maestro, para que no me lo quite de mis mascotas
                                    Npclist(.MascotasIndex(i)).MaestroUser = 0
                                    ' Lo borro
                                    Call QuitarNPC(.MascotasIndex(i))
                                    ' Saco el índice
                                    .MascotasIndex(i) = 0
                                    
                                    b = True
                                End If
                            End If
                        Next
                        
                        .flags.MascotasGuardadas = 1

                    ' Ya están guardadas, así que las invocamos
                    Else
                        For i = 1 To MAXMASCOTAS
                            ' Si está guardada y no está ya en el mapa
                            If .MascotasType(i) > 0 And .MascotasIndex(i) = 0 Then
                                .MascotasIndex(i) = SpawnNpc(.MascotasType(i), TargetPos, True, True)

                                Npclist(.MascotasIndex(i)).MaestroUser = UserIndex
                                Call FollowAmo(.MascotasIndex(i))
                                
                                b = True
                            End If
                        Next
                        
                        .flags.MascotasGuardadas = 0
                    End If
                
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar tus mascotas en un mapa seguro.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes mascotas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            If b Then Call InfoHechizo(UserIndex)
            
        End If
    
    End With
    
    Exit Sub
    
HechizoInvocacion_Err:
    Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoTerrenoEstado")
    Resume Next

End Sub

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoTerrenoEstado_Err
        

        Dim PosCasteadaX As Integer

        Dim PosCasteadaY As Integer

        Dim PosCasteadaM As Integer

        Dim h            As Integer

        Dim TempX        As Integer

        Dim TempY        As Integer

100     PosCasteadaX = UserList(UserIndex).flags.TargetX
102     PosCasteadaY = UserList(UserIndex).flags.TargetY
104     PosCasteadaM = UserList(UserIndex).flags.TargetMap
    
106     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
        
108     If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
110         b = True

            'Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(H).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
112         For TempX = PosCasteadaX - 11 To PosCasteadaX + 11
114             For TempY = PosCasteadaY - 11 To PosCasteadaY + 11

116                 If InMapBounds(PosCasteadaM, TempX, TempY) Then
118                     If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then

                            'hay un user
120                         If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.NoDetectable = 0 Then
122                             UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 0
124                             Call WriteConsoleMsg(MapData(PosCasteadaM, TempX, TempY).UserIndex, "Tu invisibilidad ya no tiene efecto.", FontTypeNames.FONTTYPE_INFOIAO)
126                             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.CharIndex, False))

                            End If

                        End If

                    End If

128             Next TempY
130         Next TempX
    
132         Call InfoHechizo(UserIndex)

        End If

        
        Exit Sub

HechizoTerrenoEstado_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoTerrenoEstado", Erl)
        Resume Next
        
End Sub

Sub HechizoSobreArea(ByVal UserIndex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoSobreArea_Err
        

        Dim PosCasteadaX As Byte

        Dim PosCasteadaY As Byte

        Dim PosCasteadaM As Integer

        Dim h            As Integer

        Dim TempX        As Integer

        Dim TempY        As Integer

100     PosCasteadaX = UserList(UserIndex).flags.TargetX
102     PosCasteadaY = UserList(UserIndex).flags.TargetY
104     PosCasteadaM = UserList(UserIndex).flags.TargetMap
 
106     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

        Dim x         As Long

        Dim Y         As Long
    
        Dim NPCIndex2 As Integer

        Dim Cuantos   As Long
    
        'Envio Palabras magicas, wavs y fxs.
108     If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
110         Call DecirPalabrasMagicas(h, UserIndex)

        End If
    
112     If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
    
114         If Hechizos(h).ParticleViaje > 0 Then
116             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, 1, Hechizos(h).wav, 1, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
                
            Else
118             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))

            End If

        End If
    
120     If Hechizos(h).Particle > 0 Then 'Envio Particula?
122         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))

        End If
    
124     If Hechizos(h).ParticleViaje = 0 Then
126         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).wav, PosCasteadaX, PosCasteadaY))  'Esta linea faltaba. Pablo (ToxicWaste)

        End If

        Dim cuantosuser As Byte

        Dim nameuser    As String
       
128     Select Case Hechizos(h).AreaAfecta

            Case 1

130             For x = 1 To Hechizos(h).AreaRadio
132                 For Y = 1 To Hechizos(h).AreaRadio

134                     If MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).UserIndex > 0 Then
136                         NPCIndex2 = MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).UserIndex

                            'If NPCIndex2 <> UserIndex Then
138                         If UserList(NPCIndex2).flags.Muerto = 0 Then
                                        
140                             AreaHechizo UserIndex, NPCIndex2, PosCasteadaX, PosCasteadaY, False
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

146             For x = 1 To Hechizos(h).AreaRadio
148                 For Y = 1 To Hechizos(h).AreaRadio

150                     If MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex > 0 Then
152                         NPCIndex2 = MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex

154                         If Npclist(NPCIndex2).Attackable Then
156                             AreaHechizo UserIndex, NPCIndex2, PosCasteadaX, PosCasteadaY, True
158                             Cuantos = Cuantos + 1

                            End If

                        End If

                    Next
                Next
                
                ' If Cuantos > 0 Then
                '  Call WriteConsoleMsg(UserIndex, "Has alcanzado a " & Cuantos & " criaturas.", FontTypeNames.FONTTYPE_FIGHT)
                '  End If
160         Case 3

162             For x = 1 To Hechizos(h).AreaRadio
164                 For Y = 1 To Hechizos(h).AreaRadio

166                     If MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).UserIndex > 0 Then
168                         NPCIndex2 = MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).UserIndex

                            'If NPCIndex2 <> UserIndex Then
170                         If UserList(NPCIndex2).flags.Muerto = 0 Then
172                             AreaHechizo UserIndex, NPCIndex2, PosCasteadaX, PosCasteadaY, False
174                             cuantosuser = cuantosuser + 1

                            End If

                            ' End If
                        End If
                            
176                     If MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex > 0 Then
178                         NPCIndex2 = MapData(UserList(UserIndex).Pos.Map, x + PosCasteadaX - CInt(Hechizos(h).AreaRadio / 2), PosCasteadaY + Y - CInt(Hechizos(h).AreaRadio / 2)).NpcIndex

180                         If Npclist(NPCIndex2).Attackable Then
182                             AreaHechizo UserIndex, NPCIndex2, PosCasteadaX, PosCasteadaY, True
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

Sub HechizoPortal(ByVal UserIndex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoPortal_Err
        

100     If UserList(UserIndex).flags.BattleModo = 1 Then
102         b = False
            'Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
104         Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)
        Else

            Dim PosCasteadaX As Byte

            Dim PosCasteadaY As Byte

            Dim PosCasteadaM As Integer

            Dim uh           As Integer

            Dim TempX        As Integer

            Dim TempY        As Integer

106         PosCasteadaX = UserList(UserIndex).flags.TargetX
108         PosCasteadaY = UserList(UserIndex).flags.TargetY
110         PosCasteadaM = UserList(UserIndex).flags.TargetMap
 
112         uh = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
            'Envio Palabras magicas, wavs y fxs.
   
114         If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.Amount > 0 Or (MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked And eBlock.ALL_SIDES) <> eBlock.ALL_SIDES Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).TileExit.Map > 0 Or UserList(UserIndex).flags.TargetUser <> 0 Then
116             b = False
                'Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
118             Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)

            Else

120             If Hechizos(uh).TeleportX = 1 Then

122                 If UserList(UserIndex).flags.Portal = 0 Then

124                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, ParticulasIndex.Runa, -1, False))
            
126                     UserList(UserIndex).flags.PortalM = UserList(UserIndex).Pos.Map
128                     UserList(UserIndex).flags.PortalX = UserList(UserIndex).flags.TargetX
130                     UserList(UserIndex).flags.PortalY = UserList(UserIndex).flags.TargetY
            
132                     Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageBarFx(UserList(UserIndex).Char.CharIndex, 600, Accion_Barra.Intermundia))

134                     UserList(UserIndex).Accion.AccionPendiente = True
136                     UserList(UserIndex).Accion.Particula = ParticulasIndex.Runa
138                     UserList(UserIndex).Accion.TipoAccion = Accion_Barra.Intermundia
140                     UserList(UserIndex).Accion.HechizoPendiente = uh
            
142                     If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
144                         Call DecirPalabrasMagicas(uh, UserIndex)

                        End If

146                     b = True
                    Else
148                     Call WriteConsoleMsg(UserIndex, "No podés lanzar mas de un portal a la vez.", FontTypeNames.FONTTYPE_INFO)
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

Sub HechizoMaterializacion(ByVal UserIndex As Integer, ByRef b As Boolean)
        
        On Error GoTo HechizoMaterializacion_Err
        

        Dim h   As Integer

        Dim MAT As obj

100     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
 
102     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.Amount > 0 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Then
104         b = False
106         Call WriteLocaleMsg(UserIndex, "262", FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)
        Else
108         MAT.Amount = Hechizos(h).MaterializaCant
110         MAT.ObjIndex = Hechizos(h).MaterializaObj
112         Call MakeObj(MAT, UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY)
            'Call WriteConsoleMsg(UserIndex, "Has materializado un objeto!!", FontTypeNames.FONTTYPE_INFO)
114         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))
116         b = True

        End If

        
        Exit Sub

HechizoMaterializacion_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoMaterializacion", Erl)
        Resume Next
        
End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal uh As Integer)
        
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
                Call HechizoInvocacion(UserIndex, b)

102         Case TipoHechizo.uEstado 'Tipo 2
104             Call HechizoTerrenoEstado(UserIndex, b)

106         Case TipoHechizo.uMaterializa 'Tipo 3
108             Call HechizoMaterializacion(UserIndex, b)
            
110         Case TipoHechizo.uArea 'Tipo 5
112             Call HechizoSobreArea(UserIndex, b)
            
114         Case TipoHechizo.uPortal 'Tipo 6
116             Call HechizoPortal(UserIndex, b)

118         Case TipoHechizo.UFamiliar

                ' Call InvocarFamiliar(UserIndex, b)
        End Select

        'If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).ObjInfo.Amount > 0 Or MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY).Blocked Or UserList(UserIndex).flags.TargetUser <> 0 Then
        '  b = False
        '  Call WriteConsoleMsg(UserIndex, "Area invalida para lanzar este Hechizo!", FontTypeNames.FONTTYPE_INFO)

        'Else

120     If b Then
122         Call SubirSkill(UserIndex, magia)

            'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
124         If UserList(UserIndex).clase = eClass.Druid And UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
126             UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido * 0.7
            Else
128             UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

            End If

130         If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
132         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido

134         If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
136         Call WriteUpdateMana(UserIndex)
            Call WriteUpdateSta(UserIndex)

        End If

        
        Exit Sub

HandleHechizoTerreno_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HandleHechizoTerreno", Erl)
        Resume Next
        
End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal uh As Integer)
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
102             Call HechizoEstadoUsuario(UserIndex, b)

104         Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
106             Call HechizoPropUsuario(UserIndex, b)

108         Case TipoHechizo.uCombinados
110             Call HechizoCombinados(UserIndex, b)
    
        End Select

112     If b Then
114         Call SubirSkill(UserIndex, magia)
            'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
116         UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
118         If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0

            If Hechizos(uh).RequiredHP > 0 Then
120             UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - Hechizos(uh).RequiredHP
122             If UserList(UserIndex).Stats.MinHp < 0 Then UserList(UserIndex).Stats.MinHp = 1
                Call WriteUpdateHP(UserIndex)
            End If

124         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido
126         If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
128
            Call WriteUpdateMana(UserIndex)
            Call WriteUpdateSta(UserIndex)
132         UserList(UserIndex).flags.TargetUser = 0

        End If

        
        Exit Sub

HandleHechizoUsuario_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HandleHechizoUsuario", Erl)
        Resume Next
        
End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal uh As Integer)
        
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
102             Call HechizoEstadoNPC(UserList(UserIndex).flags.TargetNPC, uh, b, UserIndex)

104         Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
106             Call HechizoPropNPC(uh, UserList(UserIndex).flags.TargetNPC, UserIndex, b)

        End Select

108     If b Then
110         Call SubirSkill(UserIndex, magia)
112         UserList(UserIndex).flags.TargetNPC = 0
114         UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido

            If Hechizos(uh).RequiredHP > 0 Then
116             If UserList(UserIndex).Stats.MinMAN < 0 Then UserList(UserIndex).Stats.MinMAN = 0
118             UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp - Hechizos(uh).RequiredHP
                Call WriteUpdateHP(UserIndex)
            End If

120         If UserList(UserIndex).Stats.MinHp < 0 Then UserList(UserIndex).Stats.MinHp = 1
122         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Hechizos(uh).StaRequerido

124         If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
126         Call WriteUpdateMana(UserIndex)
            Call WriteUpdateSta(UserIndex)

        End If

        
        Exit Sub

HandleHechizoNPC_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HandleHechizoNPC", Erl)
        Resume Next
        
End Sub

Sub LanzarHechizo(Index As Integer, UserIndex As Integer)
        
        On Error GoTo LanzarHechizo_Err
        

        Dim uh As Integer

100     uh = UserList(UserIndex).Stats.UserHechizos(Index)

102     If PuedeLanzar(UserIndex, uh, Index) Then

104         Select Case Hechizos(uh).Target

                Case TargetType.uUsuarios

106                 If UserList(UserIndex).flags.TargetUser > 0 Then
108                     If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
110                         Call HandleHechizoUsuario(UserIndex, uh)
                    
112                         If Hechizos(uh).CoolDown > 0 Then
114                             UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF

                            End If

                        Else
116                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
118                     Call WriteConsoleMsg(UserIndex, "Este hechizo actua solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
120             Case TargetType.uNPC

122                 If UserList(UserIndex).flags.TargetNPC > 0 Then
124                     If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
126                         Call HandleHechizoNPC(UserIndex, uh)

128                         If Hechizos(uh).CoolDown > 0 Then
130                             UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF
                    
                            End If
                    
                        Else
132                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
134                     Call WriteConsoleMsg(UserIndex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
136             Case TargetType.uUsuariosYnpc

138                 If UserList(UserIndex).flags.TargetUser > 0 Then
140                     If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
142                         Call HandleHechizoUsuario(UserIndex, uh)
                    
144                         If Hechizos(uh).CoolDown > 0 Then
146                             UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF

                            End If

                        Else
148                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                            'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

150                 ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then

152                     If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y Then
154                         If Hechizos(uh).CoolDown > 0 Then
156                             UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF

                            End If

158                         Call HandleHechizoNPC(UserIndex, uh)
                        Else
160                         Call WriteLocaleMsg(UserIndex, "8", FontTypeNames.FONTTYPE_INFO)

                            ' Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                        End If

                    Else
162                     Call WriteConsoleMsg(UserIndex, "Target invalido.", FontTypeNames.FONTTYPE_INFO)

                    End If
        
164             Case TargetType.uTerreno

166                 If Hechizos(uh).CoolDown > 0 Then
168                     UserList(UserIndex).Counters.UserHechizosInterval(Index) = GetTickCount() And &H7FFFFFFF

                    End If

170                 Call HandleHechizoTerreno(UserIndex, uh)

            End Select
    
        End If

172     If UserList(UserIndex).Counters.Trabajando Then
174         Call WriteMacroTrabajoToggle(UserIndex, False)

        End If

176     If UserList(UserIndex).Counters.Ocultando Then UserList(UserIndex).Counters.Ocultando = UserList(UserIndex).Counters.Ocultando - 1
    
        
        Exit Sub

LanzarHechizo_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.LanzarHechizo", Erl)
        Resume Next
        
End Sub

Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
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

100     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
102     tU = UserList(UserIndex).flags.TargetUser

104     If Hechizos(h).Invisibilidad = 1 Then
   
106         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
108             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
110             b = False
                Exit Sub

            End If
    
112         If UserList(tU).Counters.Saliendo Then
114             If UserIndex <> tU Then
116                 Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
118                 b = False
                    Exit Sub
                Else
120                 Call WriteConsoleMsg(UserIndex, "¡No podés ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
122                 b = False
                    Exit Sub

                End If

            End If
    
            'No usar invi mapas InviSinEfecto
            ' If MapInfo(UserList(tU).Pos.map).InviSinEfecto > 0 Then
            '  Call WriteConsoleMsg(UserIndex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
            '  b = False
            '   Exit Sub
            '  End If
    
            'Para poder tirar invi a un pk en el ring
124         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
126             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
128                 If esArmada(UserIndex) Then
130                     Call WriteConsoleMsg(UserIndex, "Los miembros de la armada real no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
132                     b = False
                        Exit Sub

                    End If

134                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
136                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
138                     b = False
                        Exit Sub
                    Else
140                     Call VolverCriminal(UserIndex)

                    End If

                End If

            End If
    
            'Si sos user, no uses este hechizo con GMS.
142         If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
144             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
            
            If UserList(tU).flags.invisible = 1 Then
                If tU = UserIndex Then
                    Call WriteConsoleMsg(UserIndex, "¡Ya estás invisible!", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "¡El objetivo ya se encuentra invisible!", FontTypeNames.FONTTYPE_INFO)
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

154         Call InfoHechizo(UserIndex)
156         b = True

        End If
        
        If Hechizos(h).Mimetiza = 1 Then
            If UserList(tU).flags.Muerto = 1 Then
                Exit Sub
            End If
            
            If UserList(tU).flags.Navegando = 1 Then
                Exit Sub
            End If
            If UserList(UserIndex).flags.Navegando = 1 Then
                Exit Sub
            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
                If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub
                End If
            End If
            
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                Call WriteConsoleMsg(UserIndex, "Ya te encuentras transformado. El hechizo no ha tenido efecto", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
            
            'copio el char original al mimetizado
            
            With UserList(UserIndex)
                .CharMimetizado.Body = .Char.Body
                .CharMimetizado.Head = .Char.Head
                .CharMimetizado.CascoAnim = .Char.CascoAnim
                .CharMimetizado.ShieldAnim = .Char.ShieldAnim
                .CharMimetizado.WeaponAnim = .Char.WeaponAnim
                
                .flags.Mimetizado = 1
                
                'ahora pongo local el del enemigo
                .Char.Body = UserList(tU).Char.Body
                .Char.Head = UserList(tU).Char.Head
                .Char.CascoAnim = UserList(tU).Char.CascoAnim
                .Char.ShieldAnim = UserList(tU).Char.ShieldAnim
                .Char.WeaponAnim = UserList(tU).Char.WeaponAnim
            
                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
            End With
           
           Call InfoHechizo(UserIndex)
           b = True
        End If

158     If Hechizos(h).Envenena > 0 Then
            ' If UserIndex = tU Then
            '    Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            '   Exit Sub
            'End If
    
160         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
162         If UserIndex <> tU Then
164             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

166         UserList(tU).flags.Envenenado = Hechizos(h).Envenena
168         Call InfoHechizo(UserIndex)
170         b = True

        End If

172     If Hechizos(h).desencantar = 1 Then
            ' Call WriteConsoleMsg(UserIndex, "Ningun efecto magico tiene efecto sobre ti ya.", FontTypeNames.FONTTYPE_INFOIAO)

174         UserList(UserIndex).flags.Envenenado = 0
176         UserList(UserIndex).flags.Incinerado = 0
    
178         If UserList(UserIndex).flags.Inmovilizado = 1 Then
180             UserList(UserIndex).Counters.Inmovilizado = 0
182             UserList(UserIndex).flags.Inmovilizado = 0
184             Call WriteInmovilizaOK(UserIndex)
            

            End If
    
186         If UserList(UserIndex).flags.Paralizado = 1 Then
188             UserList(UserIndex).flags.Paralizado = 0
190             Call WriteParalizeOK(UserIndex)
            
           
            End If
        
192         If UserList(UserIndex).flags.Ceguera = 1 Then
194             UserList(UserIndex).flags.Ceguera = 0
196             Call WriteBlindNoMore(UserIndex)
            

            End If
    
198         If UserList(UserIndex).flags.Maldicion = 1 Then
200             UserList(UserIndex).flags.Maldicion = 0
202             UserList(UserIndex).Counters.Maldicion = 0

            End If
    
204         Call InfoHechizo(UserIndex)
206         b = True

        End If

208     If Hechizos(h).incinera = 1 Then
210         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
212             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
214         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
216         If UserIndex <> tU Then
218             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

220         UserList(tU).flags.Incinerado = 1
222         Call InfoHechizo(UserIndex)
224         b = True

        End If

226     If Hechizos(h).CuraVeneno = 1 Then

            'Verificamos que el usuario no este muerto
228         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
230             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
232             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar veneno a un pk en el ring
234         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
236             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
238                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
240                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
242                     b = False
                        Exit Sub

                    End If

244                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
246                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
248                     b = False
                        Exit Sub
                    Else

                        '    Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
        
            'Si sos user, no uses este hechizo con GMS.
250         If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
252             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
        
254         UserList(tU).flags.Envenenado = 0
256         Call InfoHechizo(UserIndex)
258         b = True

        End If

260     If Hechizos(h).Maldicion = 1 Then
262         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
264             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
266         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
268         If UserIndex <> tU Then
270             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

272         UserList(tU).flags.Maldicion = 1
274         UserList(tU).Counters.Maldicion = 200
    
276         Call InfoHechizo(UserIndex)
278         b = True

        End If

280     If Hechizos(h).RemoverMaldicion = 1 Then
282         UserList(tU).flags.Maldicion = 0
284         Call InfoHechizo(UserIndex)
286         b = True

        End If

288     If Hechizos(h).GolpeCertero = 1 Then
290         UserList(tU).flags.GolpeCertero = 1
292         Call InfoHechizo(UserIndex)
294         b = True

        End If

296     If Hechizos(h).Bendicion = 1 Then
298         UserList(tU).flags.Bendicion = 1
300         Call InfoHechizo(UserIndex)
302         b = True

        End If

304     If Hechizos(h).Paraliza = 1 Then
306         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
308             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            If UserList(tU).flags.Paralizado = 1 Then
309             Call WriteConsoleMsg(UserIndex, UserList(tU).name & " ya está paralizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
310         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
312         If UserIndex <> tU Then
314             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If
            
316         Call InfoHechizo(UserIndex)
318         b = True

320         If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
322             Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
324             Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            
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
338         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
340             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
342         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
344         If UserIndex <> tU Then
346             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If
            
348         Call InfoHechizo(UserIndex)
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
364         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
366             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
            If UserList(tU).flags.Paralizado = 1 Then
                Call WriteConsoleMsg(UserIndex, UserList(tU).name & " ya está paralizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
368         ElseIf UserList(tU).flags.Inmovilizado = 1 Then
370             Call WriteConsoleMsg(UserIndex, UserList(tU).name & " ya está inmovilizado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
    
372         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
374         If UserIndex <> tU Then
376             Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            End If
            
378         Call InfoHechizo(UserIndex)
380         b = True
            '  If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
            '   Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
            '   Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
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
394         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
396             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
398                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
400                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
402                     b = False
                        Exit Sub

                    End If

404                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
406                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
408                     b = False
                        Exit Sub
                    Else
410                     Call VolverCriminal(UserIndex)

                    End If

                End If

            End If
        
412         If UserList(tU).flags.Inmovilizado = 0 And UserList(tU).flags.Paralizado = 0 Then
414             Call WriteConsoleMsg(UserIndex, "El objetivo no esta paralizado.", FontTypeNames.FONTTYPE_INFO)
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
436         Call InfoHechizo(UserIndex)

        End If

438     If Hechizos(h).RemoverEstupidez = 1 Then
440         If UserList(tU).flags.Estupidez = 1 Then

                'Para poder tirar remo estu a un pk en el ring
442             If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
444                 If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
446                     If esArmada(UserIndex) Then
                            'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
448                         Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
450                         b = False
                            Exit Sub

                        End If

452                     If UserList(UserIndex).flags.Seguro Then
                            'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
454                         Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
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
            
462             Call InfoHechizo(UserIndex)
464             b = True

            End If

        End If

466     If Hechizos(h).Revivir = 1 Then
468         If UserList(tU).flags.Muerto = 1 Then
    
                'No usar resu en mapas con ResuSinEfecto
                'If MapInfo(UserList(tU).Pos.map).ResuSinEfecto > 0 Then
                '   Call WriteConsoleMsg(UserIndex, "¡Revivir no está permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
                '   b = False
                '   Exit Sub
                ' End If
        
470             If UserList(tU).Accion.TipoAccion = Accion_Barra.Resucitar Then
472                 Call WriteConsoleMsg(UserIndex, "El usuario ya esta siendo resucitado.", FontTypeNames.FONTTYPE_INFO)
474                 b = False
                    Exit Sub

                End If
        
                'Para poder tirar revivir a un pk en el ring
476             If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
478                 If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
480                     If esArmada(UserIndex) Then
                            'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
482                         Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
484                         b = False
                            Exit Sub

                        End If

486                     If UserList(UserIndex).flags.Seguro Then
                            'call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
488                         Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
490                         b = False
                            Exit Sub
                        Else
492                         Call VolverCriminal(UserIndex)

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
506             Call InfoHechizo(UserIndex)
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
514         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
516             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
518         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
520         If UserIndex <> tU Then
522             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

524         UserList(tU).flags.Ceguera = 1
526         UserList(tU).Counters.Ceguera = Hechizos(h).Duration

528         Call WriteBlind(tU)
        
530         Call InfoHechizo(UserIndex)
532         b = True

        End If

534     If Hechizos(h).Estupidez = 1 Then
536         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
538             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

540         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
542         If UserIndex <> tU Then
544             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

546         If UserList(tU).flags.Estupidez = 0 Then
548             UserList(tU).flags.Estupidez = 1
550             UserList(tU).Counters.Estupidez = Hechizos(h).Duration

            End If

552         Call WriteDumb(tU)
        

554         Call InfoHechizo(UserIndex)
556         b = True

        End If

        
        Exit Sub

HechizoEstadoUsuario_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoEstadoUsuario", Erl)
        Resume Next
        
End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal UserIndex As Integer)
        
        On Error GoTo HechizoEstadoNPC_Err
        

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 04/13/2008
        'Handles the Spells that afect the Stats of an NPC
        '04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
        'removidos por users de su misma faccion.
        '***************************************************
100     If Hechizos(hIndex).Invisibilidad = 1 Then
102         Call InfoHechizo(UserIndex)
104         Npclist(NpcIndex).flags.invisible = 1
106         b = True

        End If

108     If Hechizos(hIndex).Envenena > 0 Then
110         If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
112             b = False
                Exit Sub

            End If

114         Call NPCAtacado(NpcIndex, UserIndex)
116         Call InfoHechizo(UserIndex)
118         Npclist(NpcIndex).flags.Envenenado = Hechizos(hIndex).Envenena
120         b = True

        End If

122     If Hechizos(hIndex).CuraVeneno = 1 Then
124         Call InfoHechizo(UserIndex)
126         Npclist(NpcIndex).flags.Envenenado = 0
128         b = True

        End If

130     If Hechizos(hIndex).RemoverMaldicion = 1 Then
132         Call InfoHechizo(UserIndex)
            'Npclist(NpcIndex).flags.Maldicion = 0
134         b = True

        End If

136     If Hechizos(hIndex).Bendicion = 1 Then
138         Call InfoHechizo(UserIndex)
140         Npclist(NpcIndex).flags.Bendicion = 1
142         b = True

        End If

144     If Hechizos(hIndex).Paraliza = 1 Then
146         If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
148             If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
150                 b = False
                    Exit Sub

                End If

152             Call NPCAtacado(NpcIndex, UserIndex)
154             Call InfoHechizo(UserIndex)
156             Npclist(NpcIndex).flags.Paralizado = 1
158             Npclist(NpcIndex).flags.Inmovilizado = 0
160             Npclist(NpcIndex).Contadores.Paralisis = (Hechizos(hIndex).Duration * 6.5) * 2
162             b = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
164             Call WriteLocaleMsg(UserIndex, "381", FontTypeNames.FONTTYPE_INFO)
166             b = False
                Exit Sub

            End If

        End If

168     If Hechizos(hIndex).RemoverParalisis = 1 Then
170         If Npclist(NpcIndex).flags.Paralizado = 1 Or Npclist(NpcIndex).flags.Inmovilizado = 1 Then
172             If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
174                 If esArmada(UserIndex) Then
176                     Call InfoHechizo(UserIndex)
178                     Npclist(NpcIndex).flags.Paralizado = 0
180                     Npclist(NpcIndex).Contadores.Paralisis = 0
182                     b = True
                        Exit Sub
                    Else
184                     Call WriteConsoleMsg(UserIndex, "Solo podés Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
186                     b = False
                        Exit Sub

                    End If
                
188                 Call WriteConsoleMsg(UserIndex, "Solo podés Remover la Parálisis de los NPCs que te consideren su amo", FontTypeNames.FONTTYPE_INFO)
190                 b = False
                    Exit Sub
                Else

192                 If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
194                     If esCaos(UserIndex) Then
196                         Call InfoHechizo(UserIndex)
198                         Npclist(NpcIndex).flags.Paralizado = 0
200                         Npclist(NpcIndex).Contadores.Paralisis = 0
202                         b = True
                            Exit Sub
                        Else
204                         Call WriteConsoleMsg(UserIndex, "Solo podés Remover la Parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
206                         b = False
                            Exit Sub

                        End If

                    End If

                End If

            Else
208             Call WriteConsoleMsg(UserIndex, "Este NPC no esta Paralizado", FontTypeNames.FONTTYPE_INFO)
210             b = False
                Exit Sub

            End If

        End If
 
212     If Hechizos(hIndex).Inmoviliza = 1 And Npclist(NpcIndex).flags.Paralizado = 0 Then
214         If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
216             If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
218                 b = False
                    Exit Sub

                End If

220             Call NPCAtacado(NpcIndex, UserIndex)
222             Npclist(NpcIndex).flags.Inmovilizado = 1
224             Npclist(NpcIndex).flags.Paralizado = 0
226             Npclist(NpcIndex).Contadores.Paralisis = (Hechizos(hIndex).Duration * 6.5) * 2
228             Call InfoHechizo(UserIndex)
230             b = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
232             Call WriteLocaleMsg(UserIndex, "381", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        If Hechizos(hIndex).Mimetiza = 1 Then
    
            If UserList(UserIndex).flags.Mimetizado = 1 Then
                Call WriteConsoleMsg(UserIndex, "Ya te encuentras transformado. El hechizo no ha tenido efecto", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.AdminInvisible = 1 Then Exit Sub
            
                
            If UserList(UserIndex).clase = eClass.Druid Then
                'copio el char original al mimetizado
                With UserList(UserIndex)
                    .CharMimetizado.Body = .Char.Body
                    .CharMimetizado.Head = .Char.Head
                    .CharMimetizado.CascoAnim = .Char.CascoAnim
                    .CharMimetizado.ShieldAnim = .Char.ShieldAnim
                    .CharMimetizado.WeaponAnim = .Char.WeaponAnim
                    
                    .flags.Mimetizado = 1
                    
                    'ahora pongo lo del NPC.
                    .Char.Body = Npclist(NpcIndex).Char.Body
                    .Char.Head = Npclist(NpcIndex).Char.Head
                    .Char.CascoAnim = NingunCasco
                    .Char.ShieldAnim = NingunEscudo
                    .Char.WeaponAnim = NingunArma
                
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                End With
            Else
                Call WriteConsoleMsg(UserIndex, "Solo los druidas pueden mimetizarse con criaturas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
           Call InfoHechizo(UserIndex)
           b = True
        End If
        
        Exit Sub

HechizoEstadoNPC_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoEstadoNPC", Erl)
        Resume Next
        
End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef b As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 14/08/2007
        'Handles the Spells that afect the Life NPC
        '14/08/2007 Pablo (ToxicWaste) - Orden general.
        '***************************************************
        
        On Error GoTo HechizoPropNPC_Err
        

        Dim daño As Long
    
        'Salud
100     If Hechizos(hIndex).SubeHP = 1 Then
102         daño = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
            'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
        
104         Call InfoHechizo(UserIndex)
106         Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp + daño

108         If Npclist(NpcIndex).Stats.MinHp > Npclist(NpcIndex).Stats.MaxHp Then Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MaxHp
110         Call WriteConsoleMsg(UserIndex, "Has curado " & daño & " puntos de salud a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
112         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead(daño, Npclist(NpcIndex).Char.CharIndex, vbGreen))
114         b = True
        
116     ElseIf Hechizos(hIndex).SubeHP = 2 Then

118         If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
120             b = False
                Exit Sub

            End If
        
122         Call NPCAtacado(NpcIndex, UserIndex)
124         daño = RandomNumber(Hechizos(hIndex).MinHp, Hechizos(hIndex).MaxHp)
        
126         daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
            ' Daño mágico arma
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                daño = daño + Porcentaje(daño, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
            End If
            
            ' Daño mágico anillo
            If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                daño = daño + Porcentaje(daño, ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MagicDamageBonus)
            End If

132         b = True
        
134         If Npclist(NpcIndex).flags.Snd2 > 0 Then
136             Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.x, Npclist(NpcIndex).Pos.Y))
            End If
        
            'Quizas tenga defenza magica el NPC.
            If Hechizos(hIndex).AntiRm = 0 Then
138             daño = daño - Npclist(NpcIndex).Stats.defM
            End If
        
140         If daño < 0 Then daño = 0
        
142         Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp - daño
144         Call InfoHechizo(UserIndex)
        
146         If UserList(UserIndex).ChatCombate = 1 Then
148             Call WriteConsoleMsg(UserIndex, "Le has causado " & daño & " puntos de daño a la criatura!", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
150         Call CalcularDarExp(UserIndex, NpcIndex, daño)
    
152         Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageEfectOverHead(daño, Npclist(NpcIndex).Char.CharIndex))
    
154         If Npclist(NpcIndex).Stats.MinHp < 1 Then
156             Npclist(NpcIndex).Stats.MinHp = 0
158             Call MuereNpc(NpcIndex, UserIndex)

            End If

        End If

        
        Exit Sub

HechizoPropNPC_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoPropNPC", Erl)
        Resume Next
        
End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)
        
        On Error GoTo InfoHechizo_Err
        

        Dim h As Integer

100     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
    
102     If UserList(UserIndex).flags.NoPalabrasMagicas = 0 Then
104         Call DecirPalabrasMagicas(h, UserIndex)

        End If

106     If UserList(UserIndex).flags.TargetUser > 0 Then '¿El Hechizo fue tirado sobre un usuario?
108         If Hechizos(h).FXgrh > 0 Then '¿Envio FX?
110             If Hechizos(h).ParticleViaje > 0 Then
112                 Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.CharIndex, UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1))
                Else
114                 Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageCreateFX(UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))

                End If

            End If

116         If Hechizos(h).Particle > 0 Then '¿Envio Particula?
118             If Hechizos(h).ParticleViaje > 0 Then
120                 Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.CharIndex, UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0))
                Else
122                 Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessageParticleFX(UserList(UserList(UserIndex).flags.TargetUser).Char.CharIndex, Hechizos(h).Particle, Hechizos(h).TimeParticula, False))

                End If

            End If
        
124         If Hechizos(h).ParticleViaje = 0 Then
126             Call SendData(SendTarget.ToPCArea, UserList(UserIndex).flags.TargetUser, PrepareMessagePlayWave(Hechizos(h).wav, UserList(UserList(UserIndex).flags.TargetUser).Pos.x, UserList(UserList(UserIndex).flags.TargetUser).Pos.Y)) 'Esta linea faltaba. Pablo (ToxicWaste)

            End If
        
128         If Hechizos(h).TimeEfect <> 0 Then 'Envio efecto de screen
130             Call WriteEfectToScreen(UserIndex, Hechizos(h).ScreenColor, Hechizos(h).TimeEfect)

            End If

132     ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then '¿El Hechizo fue tirado sobre un npc?

134         If Hechizos(h).FXgrh > 0 Then '¿Envio FX?
136             If Npclist(UserList(UserIndex).flags.TargetNPC).Stats.MinHp < 1 Then

                    'Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(H).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
138                 If Hechizos(h).ParticleViaje > 0 Then
140                     Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))
                    Else
142                     Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))

                    End If

                Else

144                 If Hechizos(h).ParticleViaje > 0 Then
146                     Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.CharIndex, Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).FXgrh, Hechizos(h).TimeParticula, Hechizos(h).wav, 1))
                    Else
148                     Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageCreateFX(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))

                    End If

                End If

            End If
        
150         If Hechizos(h).Particle > 0 Then '¿Envio Particula?
152             If Npclist(UserList(UserIndex).flags.TargetNPC).Stats.MinHp < 1 Then
154                 Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXWithDestinoXY(UserList(UserIndex).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.x, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y))
                    'Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXToFloor(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.X, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y, Hechizos(H).Particle, Hechizos(H).TimeParticula))
                Else

156                 If Hechizos(h).ParticleViaje > 0 Then
158                     Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFXWithDestino(UserList(UserIndex).Char.CharIndex, Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).ParticleViaje, Hechizos(h).Particle, Hechizos(h).TimeParticula, Hechizos(h).wav, 0))
                    Else
160                     Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessageParticleFX(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, Hechizos(h).Particle, Hechizos(h).TimeParticula, False))

                    End If

                End If

            End If

162         If Hechizos(h).ParticleViaje = 0 Then
164             Call SendData(SendTarget.ToNPCArea, UserList(UserIndex).flags.TargetNPC, PrepareMessagePlayWave(Hechizos(h).wav, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.x, Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y))

            End If

        Else ' Entonces debe ser sobre el terreno

166         If Hechizos(h).FXgrh > 0 Then 'Envio Fx?
168             Call modSendData.SendToAreaByPos(UserList(UserIndex).Pos.Map, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, PrepareMessageFxPiso(Hechizos(h).FXgrh, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))

            End If
        
170         If Hechizos(h).Particle > 0 Then 'Envio Particula?
172             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFXToFloor(UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY, Hechizos(h).Particle, Hechizos(h).TimeParticula))

            End If
        
174         If Hechizos(h).wav <> 0 Then
176             Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(h).wav, UserList(UserIndex).flags.TargetX, UserList(UserIndex).flags.TargetY))   'Esta linea faltaba. Pablo (ToxicWaste)

            End If
    
        End If
    
178     If UserList(UserIndex).ChatCombate = 1 Then
180         If Hechizos(h).Target = TargetType.uTerreno Then
                Call WriteConsoleMsg(UserIndex, "ProMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)
            
            ElseIf UserList(UserIndex).flags.TargetUser > 0 Then

                'Optimizacion de protocolo por Ladder
182             If UserIndex <> UserList(UserIndex).flags.TargetUser Then
184                 Call WriteConsoleMsg(UserIndex, "HecMSGU*" & h & "*" & UserList(UserList(UserIndex).flags.TargetUser).name, FontTypeNames.FONTTYPE_FIGHT)
186                 Call WriteConsoleMsg(UserList(UserIndex).flags.TargetUser, "HecMSGA*" & h & "*" & UserList(UserIndex).name, FontTypeNames.FONTTYPE_FIGHT)
    
                Else
188                 Call WriteConsoleMsg(UserIndex, "ProMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)

                End If

190         ElseIf UserList(UserIndex).flags.TargetNPC > 0 Then
192             Call WriteConsoleMsg(UserIndex, "HecMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)

            Else
                Call WriteConsoleMsg(UserIndex, "ProMSG*" & h, FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If

        
        Exit Sub

InfoHechizo_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.InfoHechizo", Erl)
        Resume Next
        
End Sub

Sub HechizoPropUsuario(ByVal UserIndex As Integer, ByRef b As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 02/01/2008
        '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
        '***************************************************
        
        On Error GoTo HechizoPropUsuario_Err
        

        Dim h As Integer

        Dim daño As Integer

        Dim tempChr           As Integer

        Dim enviarInfoHechizo As Boolean

100     enviarInfoHechizo = False
    
102     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
104     tempChr = UserList(UserIndex).flags.TargetUser
      
        'Hambre
106     If Hechizos(h).SubeHam = 1 Then
    
108         enviarInfoHechizo = True
    
110         daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
112         UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + daño

114         If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
116         If UserIndex <> tempChr Then
118             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
120             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
122             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
124         Call WriteUpdateHungerAndThirst(tempChr)
126         b = True
    
128     ElseIf Hechizos(h).SubeHam = 2 Then

130         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
132         If UserIndex <> tempChr Then
134             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
            Else
                Exit Sub

            End If
    
136         enviarInfoHechizo = True
    
138         daño = RandomNumber(Hechizos(h).MinHam, Hechizos(h).MaxHam)
    
140         UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - daño
    
142         If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
144         If UserIndex <> tempChr Then
146             Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
148             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
150             Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

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
    
166         daño = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
    
168         UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + daño

170         If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
172         If UserIndex <> tempChr Then
174             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
176             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
178             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
            Call WriteUpdateHungerAndThirst(tempChr)
    
180         b = True
    
182     ElseIf Hechizos(h).SubeSed = 2 Then
    
184         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
186         If UserIndex <> tempChr Then
188             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
190         enviarInfoHechizo = True
    
192         daño = RandomNumber(Hechizos(h).MinSed, Hechizos(h).MaxSed)
    
194         UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - daño
    
196         If UserIndex <> tempChr Then
198             Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
200             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
202             Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

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
214         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
216             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
218                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
220                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
222                     b = False
                        Exit Sub

                    End If

224                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
226                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
228                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
230         enviarInfoHechizo = True
232         daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
234         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

236         UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño
         
238         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)

240         UserList(tempChr).flags.TomoPocion = True
242         b = True
244         Call WriteFYA(tempChr)
    
246     ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
248         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
250         If UserIndex <> tempChr Then
252             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
254         enviarInfoHechizo = True
    
256         UserList(tempChr).flags.TomoPocion = True
258         daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
260         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

262         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño < MINATRIBUTOS Then
264             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
            Else
266             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño

            End If
    
268         b = True
270         Call WriteFYA(tempChr)

        End If

        ' <-------- Fuerza ---------->
272     If Hechizos(h).SubeFuerza = 1 Then

            'Para poder tirar fuerza a un pk en el ring
274         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
276             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
278                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
280                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
282                     b = False
                        Exit Sub

                    End If

284                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
286                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
288                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
290         daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
292         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
      
294         UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + daño

296         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)
    
298         UserList(tempChr).flags.TomoPocion = True
            
            Call WriteFYA(tempChr)

300         b = True
    
302         enviarInfoHechizo = True
304         Call WriteFYA(tempChr)

306     ElseIf Hechizos(h).SubeFuerza = 2 Then

308         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
310         If UserIndex <> tempChr Then
312             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
314         UserList(tempChr).flags.TomoPocion = True
    
316         daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
318         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

320         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño < MINATRIBUTOS Then
322             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
            Else
324             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño

            End If

326         b = True
328         enviarInfoHechizo = True
330         Call WriteFYA(tempChr)

        End If

        'Salud
332     If Hechizos(h).SubeHP = 1 Then
    
            'Verifica que el usuario no este muerto
334         If UserList(tempChr).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
336             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
338             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar a un pk en el ring
340         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
342             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
344                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
346                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
348                     b = False
                        Exit Sub

                    End If

350                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
352                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
354                     b = False
                        Exit Sub
                    Else

                        'Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
       
356         daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
            ' daño = daño + Porcentaje(daño, 2 * UserList(UserIndex).Stats.ELV)
    
358         enviarInfoHechizo = True

360         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp + daño

362         If UserList(tempChr).Stats.MinHp > UserList(tempChr).Stats.MaxHp Then UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp
    
364         If UserIndex <> tempChr Then
366             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
368             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Else
370             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
372         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(daño, UserList(tempChr).Char.CharIndex, vbGreen))
            Call WriteUpdateHP(tempChr)
    
374         b = True

376     ElseIf Hechizos(h).SubeHP = 2 Then
    
378         If UserIndex = tempChr Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
380             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

382         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
384         daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
    
386         daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)

            ' Daño mágico arma
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                daño = daño + Porcentaje(daño, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
            End If
            
            ' Daño mágico anillo
            If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                daño = daño + Porcentaje(daño, ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MagicDamageBonus)
            End If
            
            ' Si el hechizo no ignora la RM
            If Hechizos(h).AntiRm = 0 Then
                ' Resistencia mágica armadura
                If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
                    daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica anillo
                If UserList(tempChr).Invent.AnilloEqpObjIndex > 0 Then
                    daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica escudo
                If UserList(tempChr).Invent.EscudoEqpObjIndex > 0 Then
                    daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.EscudoEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica casco
                If UserList(tempChr).Invent.CascoEqpObjIndex > 0 Then
                    daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).ResistenciaMagica)
                End If
            End If

            ' Prevengo daño negativo
            If daño < 0 Then daño = 0
    
394         If UserIndex <> tempChr Then
396             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)
            End If
    
398         enviarInfoHechizo = True
    
416         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp - daño
    
418         Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
420         Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    
422         Call SubirSkill(tempChr, Resistencia)
    
424         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(daño, UserList(tempChr).Char.CharIndex))

            'Muere
426         If UserList(tempChr).Stats.MinHp < 1 Then
                'Store it!
428             Call Statistics.StoreFrag(UserIndex, tempChr)
430             Call ContarMuerte(tempChr, UserIndex)
432             UserList(tempChr).Stats.MinHp = 0
434             Call ActStats(tempChr, UserIndex)

                '  Call UserDie(tempChr)
            End If
            
            Call WriteUpdateHP(tempChr)
    
436         b = True

        End If

        'Mana
438     If Hechizos(h).SubeMana = 1 Then
    
440         enviarInfoHechizo = True
442         UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + daño

444         If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
446         If UserIndex <> tempChr Then
448             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
450             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
452             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
            Call WriteUpdateMana(tempChr)
    
454         b = True
    
456     ElseIf Hechizos(h).SubeMana = 2 Then

458         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
460         If UserIndex <> tempChr Then
462             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
464         enviarInfoHechizo = True
    
466         If UserIndex <> tempChr Then
468             Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
470             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
472             Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
474         UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - daño

476         If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0

            Call WriteUpdateMana(tempChr)

478         b = True
    
        End If

        'Stamina
480     If Hechizos(h).SubeSta = 1 Then
482         Call InfoHechizo(UserIndex)
484         UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + daño

486         If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta

488         If UserIndex <> tempChr Then
490             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
492             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            Else
494             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
            Call WriteUpdateSta(tempChr)

496         b = True
498     ElseIf Hechizos(h).SubeSta = 2 Then

500         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
502         If UserIndex <> tempChr Then
504             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
506         enviarInfoHechizo = True
    
508         If UserIndex <> tempChr Then
510             Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
512             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
            Else
514             Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
516         UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - daño
    
518         If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0

            Call WriteUpdateSta(tempChr)

520         b = True

        End If

522     If enviarInfoHechizo Then
524         Call InfoHechizo(UserIndex)

        End If

    

        
        Exit Sub

HechizoPropUsuario_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoPropUsuario", Erl)
        Resume Next
        
End Sub

Sub HechizoCombinados(ByVal UserIndex As Integer, ByRef b As Boolean)
        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 02/01/2008
        '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
        '***************************************************
        
        On Error GoTo HechizoCombinados_Err
        

        Dim h As Integer

        Dim daño As Integer

        Dim tempChr           As Integer

        Dim enviarInfoHechizo As Boolean

100     enviarInfoHechizo = False
    
102     h = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)
104     tempChr = UserList(UserIndex).flags.TargetUser
      
        ' <-------- Agilidad ---------->
106     If Hechizos(h).SubeAgilidad = 1 Then
    
            'Para poder tirar cl a un pk en el ring
108         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
110             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
112                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
114                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
116                     b = False
                        Exit Sub

                    End If

118                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
120                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
122                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
124         enviarInfoHechizo = True
126         daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
    
128         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
            'UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño
            'If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then _
             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)

130         UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño

132         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
        
134         UserList(tempChr).flags.TomoPocion = True
136         b = True
138         Call WriteFYA(tempChr)
    
140     ElseIf Hechizos(h).SubeAgilidad = 2 Then
    
142         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
144         If UserIndex <> tempChr Then
146             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
148         enviarInfoHechizo = True
    
150         UserList(tempChr).flags.TomoPocion = True
152         daño = RandomNumber(Hechizos(h).MinAgilidad, Hechizos(h).MaxAgilidad)
154         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration

156         If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño < 6 Then
158             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
            Else
160             UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño

            End If

            'If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
162         b = True
164         Call WriteFYA(tempChr)

        End If

        ' <-------- Fuerza ---------->
166     If Hechizos(h).SubeFuerza = 1 Then

            'Para poder tirar fuerza a un pk en el ring
168         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
170             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
172                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
174                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
176                     b = False
                        Exit Sub

                    End If

178                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
180                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
182                     b = False
                        Exit Sub
                    Else

                        ' Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
    
184         daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
    
186         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
      
188         UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + daño

190         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
    
192         UserList(tempChr).flags.TomoPocion = True
194         b = True
    
196         enviarInfoHechizo = True
198         Call WriteFYA(tempChr)
200     ElseIf Hechizos(h).SubeFuerza = 2 Then

202         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
204         If UserIndex <> tempChr Then
206             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
208         UserList(tempChr).flags.TomoPocion = True
    
210         daño = RandomNumber(Hechizos(h).MinFuerza, Hechizos(h).MaxFuerza)
212         UserList(tempChr).flags.DuracionEfecto = Hechizos(h).Duration
        
214         If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño < 6 Then
216             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
            Else
218             UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño

            End If
   
220         b = True
222         enviarInfoHechizo = True
224         Call WriteFYA(tempChr)

        End If

        'Salud
226     If Hechizos(h).SubeHP = 1 Then
    
            'Verifica que el usuario no este muerto
228         If UserList(tempChr).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
230             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
232             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar a un pk en el ring
234         If (TriggerZonaPelea(UserIndex, tempChr) <> TRIGGER6_PERMITE) Then
236             If Status(tempChr) = 0 And Status(UserIndex) = 1 Or Status(tempChr) = 2 And Status(UserIndex) = 1 Then
238                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
240                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
242                     b = False
                        Exit Sub

                    End If

244                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
246                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
248                     b = False
                        Exit Sub
                    Else

                        'Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
       
250         daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
            'daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
    
252         enviarInfoHechizo = True

254         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp + daño

256         If UserList(tempChr).Stats.MinHp > UserList(tempChr).Stats.MaxHp Then UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MaxHp
    
258         If UserIndex <> tempChr Then
260             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
262             Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Else
264             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
266         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(daño, UserList(tempChr).Char.CharIndex, vbGreen))
    
268         b = True
270     ElseIf Hechizos(h).SubeHP = 2 Then
    
272         If UserIndex = tempChr Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
274             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
276         daño = RandomNumber(Hechizos(h).MinHp, Hechizos(h).MaxHp)
    
278         daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
            
            ' Daño mágico arma
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                daño = daño + Porcentaje(daño, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
            End If
            
            ' Daño mágico anillo
            If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                daño = daño + Porcentaje(daño, ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MagicDamageBonus)
            End If
            
            ' Si el hechizo no ignora la RM
            If Hechizos(h).AntiRm = 0 Then
                ' Resistencia mágica armadura
                If UserList(tempChr).Invent.ArmourEqpObjIndex > 0 Then
                    daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.ArmourEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica anillo
                If UserList(tempChr).Invent.AnilloEqpObjIndex > 0 Then
                    daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica escudo
                If UserList(tempChr).Invent.EscudoEqpObjIndex > 0 Then
                    daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.EscudoEqpObjIndex).ResistenciaMagica)
                End If
                
                ' Resistencia mágica casco
                If UserList(tempChr).Invent.CascoEqpObjIndex > 0 Then
                    daño = daño - Porcentaje(daño, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).ResistenciaMagica)
                End If
            End If

            ' Prevengo daño negativo
            If daño < 0 Then daño = 0
    
286         If Not PuedeAtacar(UserIndex, tempChr) Then Exit Sub
    
288         If UserIndex <> tempChr Then
290             Call UsuarioAtacadoPorUsuario(UserIndex, tempChr)

            End If
    
292         enviarInfoHechizo = True
    
306         UserList(tempChr).Stats.MinHp = UserList(tempChr).Stats.MinHp - daño
    
308         Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
310         Call WriteConsoleMsg(tempChr, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
312         Call SubirSkill(tempChr, Resistencia)
314         Call SendData(SendTarget.ToPCArea, tempChr, PrepareMessageEfectOverHead(daño, UserList(tempChr).Char.CharIndex))

            'Muere
316         If UserList(tempChr).Stats.MinHp < 1 Then
                'Store it!
318             Call Statistics.StoreFrag(UserIndex, tempChr)
        
320             Call ContarMuerte(tempChr, UserIndex)
322             UserList(tempChr).Stats.MinHp = 0
324             Call ActStats(tempChr, UserIndex)

                'Call UserDie(tempChr)
            End If
    
326         b = True

        End If

        Dim tU As Integer

328     tU = tempChr

330     If Hechizos(h).Invisibilidad = 1 Then
   
332         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
334             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
336             b = False
                Exit Sub

            End If
    
338         If UserList(tU).Counters.Saliendo Then
340             If UserIndex <> tU Then
342                 Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
344                 b = False
                    Exit Sub
                Else
346                 Call WriteConsoleMsg(UserIndex, "¡No podés ponerte invisible mientras te encuentres saliendo!", FontTypeNames.FONTTYPE_WARNING)
348                 b = False
                    Exit Sub

                End If

            End If
    
            'Para poder tirar invi a un pk en el ring
350         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
352             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
354                 If esArmada(UserIndex) Then
356                     Call WriteConsoleMsg(UserIndex, "Los miembros de la armada real no pueden ayudar a los criminales", FontTypeNames.FONTTYPE_INFO)
358                     b = False
                        Exit Sub

                    End If

360                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
362                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
364                     b = False
                        Exit Sub
                    Else
366                     Call VolverCriminal(UserIndex)

                    End If

                End If

            End If
    
            'Si sos user, no uses este hechizo con GMS.
368         If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
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
            '    Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            '   Exit Sub
            'End If
    
386         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
388         If UserIndex <> tU Then
390             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

392         UserList(tU).flags.Envenenado = Hechizos(h).Envenena
394         enviarInfoHechizo = True
396         b = True

        End If

398     If Hechizos(h).desencantar = 1 Then
400         Call WriteConsoleMsg(UserIndex, "Has sido desencantado.", FontTypeNames.FONTTYPE_INFO)

402         UserList(UserIndex).flags.Envenenado = 0
404         UserList(UserIndex).flags.Incinerado = 0
    
406         If UserList(UserIndex).flags.Inmovilizado = 1 Then
408             UserList(UserIndex).Counters.Inmovilizado = 0
410             UserList(UserIndex).flags.Inmovilizado = 0
412             Call WriteInmovilizaOK(UserIndex)
            

            End If
    
414         If UserList(UserIndex).flags.Paralizado = 1 Then
416             UserList(UserIndex).flags.Paralizado = 0
418             Call WriteParalizeOK(UserIndex)
            
           
            End If
        
420         If UserList(UserIndex).flags.Ceguera = 1 Then
422             UserList(UserIndex).flags.Ceguera = 0
424             Call WriteBlindNoMore(UserIndex)
            

            End If
    
426         If UserList(UserIndex).flags.Maldicion = 1 Then
428             UserList(UserIndex).flags.Maldicion = 0
430             UserList(UserIndex).Counters.Maldicion = 0

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
448         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
450             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
452         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
454         If UserIndex <> tU Then
456             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

458         UserList(tU).flags.Incinerado = 1
460         enviarInfoHechizo = True
462         b = True

        End If

464     If Hechizos(h).CuraVeneno = 1 Then

            'Verificamos que el usuario no este muerto
466         If UserList(tU).flags.Muerto = 1 Then
                'Call WriteConsoleMsg(UserIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
468             Call WriteLocaleMsg(UserIndex, "77", FontTypeNames.FONTTYPE_INFO)
470             b = False
                Exit Sub

            End If
    
            'Para poder tirar curar veneno a un pk en el ring
472         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
474             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
476                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
478                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
480                     b = False
                        Exit Sub

                    End If

482                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
484                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
486                     b = False
                        Exit Sub
                    Else

                        '    Call DisNobAuBan(UserIndex, UserList(UserIndex).Reputacion.NobleRep * 0.5, 10000)
                    End If

                End If

            End If
        
            'Si sos user, no uses este hechizo con GMS.
488         If UserList(UserIndex).flags.Privilegios And PlayerType.user Then
490             If Not UserList(tU).flags.Privilegios And PlayerType.user Then
                    Exit Sub

                End If

            End If
        
492         UserList(tU).flags.Envenenado = 0
494         enviarInfoHechizo = True
496         b = True

        End If

498     If Hechizos(h).Maldicion = 1 Then
500         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
502             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
504         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
506         If UserIndex <> tU Then
508             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

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
544         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
546             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
548         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
550         If UserIndex <> tU Then
552             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If
            
554         enviarInfoHechizo = True
556         b = True

558         If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
560             Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
562             Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
            
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
576         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
578             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
580         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
582         If UserIndex <> tU Then
584             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

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
602         If (TriggerZonaPelea(UserIndex, tU) <> TRIGGER6_PERMITE) Then
604             If Status(tU) = 0 And Status(UserIndex) = 1 Or Status(tU) = 2 And Status(UserIndex) = 1 Then
606                 If esArmada(UserIndex) Then
                        'Call WriteConsoleMsg(UserIndex, "Los Armadas no pueden ayudar a los Criminales", FontTypeNames.FONTTYPE_INFO)
608                     Call WriteLocaleMsg(UserIndex, "379", FontTypeNames.FONTTYPE_INFO)
610                     b = False
                        Exit Sub

                    End If

612                 If UserList(UserIndex).flags.Seguro Then
                        'Call WriteConsoleMsg(UserIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos", FontTypeNames.FONTTYPE_INFO)
614                     Call WriteLocaleMsg(UserIndex, "378", FontTypeNames.FONTTYPE_INFO)
616                     b = False
                        Exit Sub
                    Else
618                     Call VolverCriminal(UserIndex)

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
644         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
646             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
    
648         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
650         If UserIndex <> tU Then
652             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

            End If

654         UserList(tU).flags.Ceguera = 1
656         UserList(tU).Counters.Ceguera = Hechizos(h).Duration

658         Call WriteBlind(tU)
        
660         enviarInfoHechizo = True
662         b = True

        End If

664     If Hechizos(h).Estupidez = 1 Then
666         If UserIndex = tU Then
                'Call WriteConsoleMsg(UserIndex, "No podés atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
668             Call WriteLocaleMsg(UserIndex, "380", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

670         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
672         If UserIndex <> tU Then
674             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

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

690         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
692         If UserIndex <> tU Then
694             Call UsuarioAtacadoPorUsuario(UserIndex, tU)

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
712         Call InfoHechizo(UserIndex)

        End If

    

        
        Exit Sub

HechizoCombinados_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.HechizoCombinados", Erl)
        Resume Next
        
End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal slot As Byte)
        
        On Error GoTo UpdateUserHechizos_Err
        

        'Call LogTarea("Sub UpdateUserHechizos")

        Dim LoopC As Byte

        'Actualiza un solo slot
100     If Not UpdateAll Then

            'Actualiza el inventario
102         If UserList(UserIndex).Stats.UserHechizos(slot) > 0 Then
104             Call ChangeUserHechizo(UserIndex, slot, UserList(UserIndex).Stats.UserHechizos(slot))
            Else
106             Call ChangeUserHechizo(UserIndex, slot, 0)

            End If

        Else

            'Actualiza todos los slots
108         For LoopC = 1 To MAXUSERHECHIZOS

                'Actualiza el inventario
110             If UserList(UserIndex).Stats.UserHechizos(LoopC) > 0 Then
112                 Call ChangeUserHechizo(UserIndex, LoopC, UserList(UserIndex).Stats.UserHechizos(LoopC))
                Else
114                 Call ChangeUserHechizo(UserIndex, LoopC, 0)

                End If

116         Next LoopC

        End If

        
        Exit Sub

UpdateUserHechizos_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.UpdateUserHechizos", Erl)
        Resume Next
        
End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal slot As Byte, ByVal Hechizo As Integer)
        
        On Error GoTo ChangeUserHechizo_Err
        

        'Call LogTarea("ChangeUserHechizo")
    
100     UserList(UserIndex).Stats.UserHechizos(slot) = Hechizo
    
102     If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
104         Call WriteChangeSpellSlot(UserIndex, slot)
        Else
106         Call WriteChangeSpellSlot(UserIndex, slot)

        End If

        
        Exit Sub

ChangeUserHechizo_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.ChangeUserHechizo", Erl)
        Resume Next
        
End Sub

Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)
        
        On Error GoTo DesplazarHechizo_Err
        

100     If (Dire <> 1 And Dire <> -1) Then Exit Sub
102     If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

        Dim TempHechizo As Integer

104     If Dire = 1 Then 'Mover arriba
106         If CualHechizo = 1 Then
108             Call WriteConsoleMsg(UserIndex, "No podés mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
110             TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
112             UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1)
114             UserList(UserIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

                'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
116             If UserList(UserIndex).flags.Hechizo > 0 Then
118                 UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo - 1

                End If

            End If

        Else 'mover abajo

120         If CualHechizo = MAXUSERHECHIZOS Then
122             Call WriteConsoleMsg(UserIndex, "No podés mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
124             TempHechizo = UserList(UserIndex).Stats.UserHechizos(CualHechizo)
126             UserList(UserIndex).Stats.UserHechizos(CualHechizo) = UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1)
128             UserList(UserIndex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

                'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
130             If UserList(UserIndex).flags.Hechizo > 0 Then
132                 UserList(UserIndex).flags.Hechizo = UserList(UserIndex).flags.Hechizo + 1

                End If

            End If

        End If

        
        Exit Sub

DesplazarHechizo_Err:
        Call RegistrarError(Err.Number, Err.description, "modHechizos.DesplazarHechizo", Erl)
        Resume Next
        
End Sub

Sub AreaHechizo(UserIndex As Integer, NpcIndex As Integer, x As Byte, Y As Byte, npc As Boolean)
        
        On Error GoTo AreaHechizo_Err
        

        Dim calculo      As Integer

        Dim TilesDifUser As Integer

        Dim TilesDifNpc  As Integer

        Dim tilDif       As Integer

        Dim h2           As Integer

        Dim Hit          As Integer

        Dim daño As Integer

        Dim porcentajeDesc As Integer

100     h2 = UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)

        'Calculo de descuesto de golpe por cercania.
102     TilesDifUser = x + Y

104     If npc Then
106         If Hechizos(h2).SubeHP = 2 Then
108             TilesDifNpc = Npclist(NpcIndex).Pos.x + Npclist(NpcIndex).Pos.Y
            
110             tilDif = TilesDifUser - TilesDifNpc
            
112             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

                Hit = Hit + Porcentaje(Hit, 3 * UserList(UserIndex).Stats.ELV)
            
                ' Daño mágico arma
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                End If
                
                ' Daño mágico anillo
                If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                    Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MagicDamageBonus)
                End If

                ' Disminuir daño con distancia
114             If tilDif <> 0 Then
116                 porcentajeDesc = Abs(tilDif) * 20
118                 daño = Hit / 100 * porcentajeDesc
120                 daño = Hit - daño
                Else
122                 daño = Hit
                End If
                
                ' Si el hechizo no ignora la RM
                If Hechizos(h2).AntiRm = 0 Then
                    daño = daño - Npclist(NpcIndex).Stats.defM
                End If
                
                ' Prevengo daño negativo
                If daño < 0 Then daño = 0
            
128             Npclist(NpcIndex).Stats.MinHp = Npclist(NpcIndex).Stats.MinHp - daño
            
130             If UserList(UserIndex).ChatCombate = 1 Then
132                 Call WriteConsoleMsg(UserIndex, "Le has causado " & daño & " puntos de daño a " & Npclist(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)

                End If
            
134             Call CalcularDarExp(UserIndex, NpcIndex, daño)
                
136             If Npclist(NpcIndex).Stats.MinHp <= 0 Then
                    'UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + Npclist(NpcIndex).GiveEXP
                    'UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Npclist(NpcIndex).GiveGLD
138                 Call MuereNpc(NpcIndex, UserIndex)
                End If

                Exit Sub

            End If

        Else

140         TilesDifNpc = UserList(NpcIndex).Pos.x + UserList(NpcIndex).Pos.Y
142         tilDif = TilesDifUser - TilesDifNpc

144         If Hechizos(h2).SubeHP = 2 Then
146             If UserIndex = NpcIndex Then
                    Exit Sub
                End If

148             If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
                
150             If UserIndex <> NpcIndex Then
152                 Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

                End If
                
154             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

                Hit = Hit + Porcentaje(Hit, 3 * UserList(UserIndex).Stats.ELV)
            
                ' Daño mágico arma
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).MagicDamageBonus)
                End If
                
                ' Daño mágico anillo
                If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                    Hit = Hit + Porcentaje(Hit, ObjData(UserList(UserIndex).Invent.AnilloEqpObjIndex).MagicDamageBonus)
                End If

156             If tilDif <> 0 Then
158                 porcentajeDesc = Abs(tilDif) * 20
160                 daño = Hit / 100 * porcentajeDesc
162                 daño = Hit - daño
                Else
164                 daño = Hit
                End If
                
                ' Si el hechizo no ignora la RM
                If Hechizos(h2).AntiRm = 0 Then
                    ' Resistencia mágica armadura
                    If UserList(NpcIndex).Invent.ArmourEqpObjIndex > 0 Then
                        daño = daño - Porcentaje(daño, ObjData(UserList(NpcIndex).Invent.ArmourEqpObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica anillo
                    If UserList(NpcIndex).Invent.AnilloEqpObjIndex > 0 Then
                        daño = daño - Porcentaje(daño, ObjData(UserList(NpcIndex).Invent.AnilloEqpObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica escudo
                    If UserList(NpcIndex).Invent.EscudoEqpObjIndex > 0 Then
                        daño = daño - Porcentaje(daño, ObjData(UserList(NpcIndex).Invent.EscudoEqpObjIndex).ResistenciaMagica)
                    End If
                    
                    ' Resistencia mágica casco
                    If UserList(NpcIndex).Invent.CascoEqpObjIndex > 0 Then
                        daño = daño - Porcentaje(daño, ObjData(UserList(NpcIndex).Invent.CascoEqpObjIndex).ResistenciaMagica)
                    End If
                End If
                
                ' Prevengo daño negativo
                If daño < 0 Then daño = 0

180             UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MinHp - daño
                    
182             Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de vida a " & UserList(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)
184             Call WriteConsoleMsg(NpcIndex, UserList(UserIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
186             Call SubirSkill(NpcIndex, Resistencia)
188             Call WriteUpdateUserStats(NpcIndex)
                
                'Muere
190             If UserList(NpcIndex).Stats.MinHp < 1 Then
                    'Store it!
192                 Call Statistics.StoreFrag(UserIndex, NpcIndex)
                        
194                 Call ContarMuerte(NpcIndex, UserIndex)
196                 UserList(NpcIndex).Stats.MinHp = 0
198                 Call ActStats(NpcIndex, UserIndex)

                    'Call UserDie(NpcIndex)
                End If

            End If
                
200         If Hechizos(h2).SubeHP = 1 Then
202             If (TriggerZonaPelea(UserIndex, NpcIndex) <> TRIGGER6_PERMITE) Then
204                 If Status(UserIndex) = 1 And Status(NpcIndex) <> 1 Then
                        Exit Sub

                    End If

                End If

206             Hit = RandomNumber(Hechizos(h2).MinHp, Hechizos(h2).MaxHp)

208             If tilDif <> 0 Then
210                 porcentajeDesc = Abs(tilDif) * 20
212                 daño = Hit / 100 * porcentajeDesc
214                 daño = Hit - daño
                Else
216                 daño = Hit

                End If
 
218             UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MinHp + daño

220             If UserList(NpcIndex).Stats.MinHp > UserList(NpcIndex).Stats.MaxHp Then UserList(NpcIndex).Stats.MinHp = UserList(NpcIndex).Stats.MaxHp

            End If
 
222         If UserIndex <> NpcIndex Then
224             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de vida a " & UserList(NpcIndex).name, FontTypeNames.FONTTYPE_FIGHT)
226             Call WriteConsoleMsg(NpcIndex, UserList(UserIndex).name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Else
228             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)

            End If
                    
230         Call WriteUpdateUserStats(NpcIndex)

        End If
                
232     If Hechizos(h2).Envenena > 0 Then
234         If UserIndex = NpcIndex Then
                Exit Sub

            End If
                    
236         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
                
238         If UserIndex <> NpcIndex Then
240             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
                    
242         UserList(NpcIndex).flags.Envenenado = Hechizos(h2).Envenena
244         Call WriteConsoleMsg(NpcIndex, UserList(UserIndex).name & " te ha envenenado.", FontTypeNames.FONTTYPE_FIGHT)

        End If
                
246     If Hechizos(h2).Paraliza = 1 Then
248         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
250         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
252         If UserIndex <> NpcIndex Then
254             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
            
256         Call WriteConsoleMsg(NpcIndex, "Has sido paralizado.", FontTypeNames.FONTTYPE_INFO)
258         UserList(NpcIndex).Counters.Paralisis = Hechizos(h2).Duration

260         If UserList(NpcIndex).flags.Paralizado = 0 Then
262             UserList(NpcIndex).flags.Paralizado = 1
264             Call WriteParalizeOK(NpcIndex)
            

            End If
            
        End If
                
266     If Hechizos(h2).Inmoviliza = 1 Then
268         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
270         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
272         If UserIndex <> NpcIndex Then
274             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

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
290         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
292         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
294         If UserIndex <> NpcIndex Then
296             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If
                    
298         UserList(NpcIndex).flags.Ceguera = 1
300         UserList(NpcIndex).Counters.Ceguera = Hechizos(h2).Duration
302         Call WriteConsoleMsg(NpcIndex, "Te han cegado.", FontTypeNames.FONTTYPE_INFO)
            
304         Call WriteBlind(NpcIndex)
        

        End If
                
306     If Hechizos(h2).Velocidad > 0 Then
    
308         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
310         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
312         If UserIndex <> NpcIndex Then
314             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

            End If

316         If UserList(NpcIndex).Counters.Velocidad = 0 Then
318             UserList(NpcIndex).flags.VelocidadBackup = UserList(NpcIndex).Char.speeding

            End If

320         UserList(NpcIndex).Char.speeding = Hechizos(h2).Velocidad
322         Call SendData(SendTarget.ToPCArea, NpcIndex, PrepareMessageSpeedingACT(UserList(NpcIndex).Char.CharIndex, UserList(NpcIndex).Char.speeding))
324         UserList(NpcIndex).Counters.Velocidad = Hechizos(h2).Duration

        End If
                
326     If Hechizos(h2).Maldicion = 1 Then
328         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
330         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
332         If UserIndex <> NpcIndex Then
334             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

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
362         If UserIndex = NpcIndex Then
                Exit Sub

            End If
    
364         If Not PuedeAtacar(UserIndex, NpcIndex) Then Exit Sub
            
366         If UserIndex <> NpcIndex Then
368             Call UsuarioAtacadoPorUsuario(UserIndex, NpcIndex)

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
